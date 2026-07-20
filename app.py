"""Daily Operations Report (DOR) — Streamlit UI.

Run with:  streamlit run app.py

All inputs sit in one block at the top: the compiled orderbook, the compiled
user MTM, the Combined Max Loss file(s) and the outlier deviation. One
"Process" click computes BOTH reports — the trade value analysis and the
Int / Pos+Int segregation pivot — and offers exactly two downloads:

  * DOR_<date>.xlsx — all 8 data sheets in one workbook (tradevalue, summary,
    Strikes, Outlier Clients, Segregation, Raw_Data_Per_User, Worst 10%ile,
    Slippage)
  * DOR_<date>.html — the styled, client-shareable summary report
"""

import io
from decimal import Decimal

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl import Workbook

import segregate_int_pos_mtm2 as seg
from dor import build_dor_html
from tradevalue import (
    OUTLIER_CLIENTS_HEADER,
    REPORT_HEADER,
    STRIKE_ALGO_HEADER,
    STRIKE_CHAIN_HEADER,
    SUMMARY_HEADER,
    add_outlier_clients_sheet,
    add_report_sheets,
    add_strikes_sheet,
    aggregate,
    algo_summary,
    dedup_orders,
    format_lots_band,
    format_row,
    format_summary_row,
    outlier_clients,
    read_allocations,
    read_orderbook,
    report_totals,
    strike_chain,
    strike_report,
    user_lot_observations,
)

st.set_page_config(page_title="Daily Operations Report", layout="wide")
st.title("Daily Operations Report")
st.caption(
    "Upload all four inputs, set the outlier deviation, and click **Process** — "
    "both the Trade Value analysis and the Int / Pos+Int Segregation are computed "
    "in one go, with a single Excel workbook and a shareable DOR.html as output."
)

# ---------------------------------------------------------------------------
# Inputs — everything at the top
# ---------------------------------------------------------------------------
in1, in2 = st.columns(2)
with in1:
    orderbook_file = st.file_uploader("Orderbook (CSV / Excel)", type=["csv", "xlsx", "xlsm", "xls"])
with in2:
    summary_file = st.file_uploader("Compiled User MTM (CSV / Excel)", type=["csv", "xlsx", "xlsm", "xls"])

in3, in4 = st.columns(2)
with in3:
    maxloss_1dte = st.file_uploader("Combined Max Loss — 1DTE (required)", type=["xlsx", "xlsm", "xls"])
with in4:
    maxloss_4dte = st.file_uploader("Combined Max Loss — 4DTE (optional)", type=["xlsx", "xlsm", "xls"])

dev_col, dev2_col, btn_col = st.columns([1, 1, 2])
with dev_col:
    deviation = st.number_input(
        "Outlier deviation (× std dev)",
        min_value=0.1, max_value=10.0, value=1.0, step=0.5,
        help="Drives the box plot and the row-level outlier flags: a row is "
             "flagged when its lot pct is more than this many standard "
             "deviations away from its algo's average lot pct.",
    )
with dev2_col:
    client_deviation = st.number_input(
        "Client outlier deviation (× std dev)",
        min_value=0.1, max_value=10.0, value=2.0, step=0.5,
        help="Stricter threshold for the Outlier Clients table: lists the "
             "users whose lot pct is more than this many standard deviations "
             "from their algo's average, with the band converted into lots.",
    )
with btn_col:
    st.write("")  # aligns the button with the input
    st.write("")
    process_clicked = st.button(
        "Process",
        type="primary",
        disabled=(orderbook_file is None or summary_file is None or maxloss_1dte is None),
    )


# ---------------------------------------------------------------------------
# Processing — one pass computes both reports
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _process_all(ob_bytes, ob_name, mtm_bytes, mtm_name, one_bytes, four_bytes,
                 deviation_value, schema):
    # `schema` only serves as part of the cache key: when the report columns
    # change, stale cached rows from an older code version are not reused.
    multiplier = Decimal(str(deviation_value))

    # Trade value
    orders = read_orderbook(io.BytesIO(ob_bytes), ob_name)
    allocations = read_allocations(io.BytesIO(mtm_bytes), mtm_name)
    deduped = dedup_orders(orders)
    tv_rows = aggregate(deduped, allocations, multiplier)
    strikes = strike_report(deduped, allocations)

    # Segregation
    four = seg.load_combined(io.BytesIO(four_bytes)) if four_bytes else {}
    one = seg.load_combined(io.BytesIO(one_bytes))
    comp = seg.prepare_comp(io.BytesIO(mtm_bytes), four, one)
    report_date = seg.infer_report_date(comp)

    return {
        "tv_rows": tv_rows,
        "order_count": len(orders),
        "allocation_count": len(allocations),
        "strikes": strikes,
        "pivot_rows": seg.pivot_rows(comp),
        "pivot_stats": seg.comp_stats(comp),
        "comp": comp,
        "date": report_date,
    }


def _signature():
    return (
        orderbook_file.name if orderbook_file else None,
        summary_file.name if summary_file else None,
        maxloss_1dte.name if maxloss_1dte else None,
        maxloss_4dte.name if maxloss_4dte else None,
        deviation,
    )


if process_clicked:
    try:
        with st.spinner("Processing…"):
            result = _process_all(
                orderbook_file.getvalue(), orderbook_file.name,
                summary_file.getvalue(), summary_file.name,
                maxloss_1dte.getvalue(),
                maxloss_4dte.getvalue() if maxloss_4dte else None,
                deviation,
                tuple(REPORT_HEADER + STRIKE_ALGO_HEADER + STRIKE_CHAIN_HEADER
                      + OUTLIER_CLIENTS_HEADER),
            )
        st.session_state["dor"] = {**result, "deviation": deviation, "signature": _signature()}
    except Exception as exc:
        st.error(f"Processing failed: {exc}")


def _boxplot_chart(observations):
    """Box plot of per-USER lot pct per algo (one point per user — combined
    lots across segments ÷ combined allocation). The boxes show the
    distribution; the std-dev-flagged outliers are overlaid as points (blue
    below, orange above — a colorblind-safe pair) with user details on
    hover."""
    records = [
        {
            "Algo": str(r["algo"]),
            "Lot Pct": round(float(r["lot_pct"]), 4),
            "User ID": r["user_id"],
            "Server": r["server"],
            "Lots": int(round(float(r["lots"]))),
            "Outlier": r["outlier"],
        }
        for r in observations
        if r.get("lot_pct") is not None and r["algo"]
    ]
    if not records:
        return None
    chart_df = pd.DataFrame(records)
    algo_order = sorted(chart_df["Algo"].unique(),
                        key=lambda a: (0, int(a)) if a.isdigit() else (1, a))
    x = alt.X("Algo:N", sort=algo_order, title="Algo", axis=alt.Axis(labelAngle=0))
    y = alt.Y("Lot Pct:Q", title="Lot Pct")

    box = alt.Chart(chart_df).mark_boxplot(
        extent=1.5, outliers=False, size=28,
        box={"fill": "#9AA4B1", "fillOpacity": 0.45, "stroke": "#6B7684"},
        median={"stroke": "#40474F", "strokeWidth": 2},
        rule={"stroke": "#6B7684"},
        ticks={"stroke": "#6B7684"},
    ).encode(x=x, y=y)

    flagged = chart_df[chart_df["Outlier"].isin(["Below average range", "Above average range"])]
    points = alt.Chart(flagged).mark_point(filled=True, size=55, opacity=0.85).encode(
        x=x, y=y,
        color=alt.Color(
            "Outlier:N", title="Outlier",
            scale=alt.Scale(domain=["Below average range", "Above average range"],
                            range=["#0072B2", "#E69F00"]),
            legend=alt.Legend(orient="top"),
        ),
        tooltip=[
            alt.Tooltip("User ID:N"),
            alt.Tooltip("Server:N"),
            alt.Tooltip("Algo:N"),
            alt.Tooltip("Lot Pct:Q", format=".2f"),
            alt.Tooltip("Lots:Q", format=",.0f"),
            alt.Tooltip("Outlier:N", title="Flag"),
        ],
    )
    return (box + points).properties(height=420)


PIVOT_COLS = ["Section", "ALGO", "SERVER", "No. of Users", "No. of SL Hit Users",
              "MAX LOSS", "ALLOCATION", "Realized P&L", "Unrealized P&L", "MTM",
              "Return %", "95%", "5%"]

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
dor_state = st.session_state.get("dor")
if dor_state is None:
    st.info("Upload the orderbook, the Compiled User MTM and the 1DTE Max Loss file, then click **Process**.")
    st.stop()

if dor_state["signature"] != _signature():
    st.warning("Inputs changed since the last run — click **Process** to refresh the results.")

tv_rows = dor_state["tv_rows"]
used_deviation = dor_state["deviation"]
used_multiplier = Decimal(str(used_deviation))
report_date = dor_state["date"]

if not tv_rows:
    st.warning("No completed NIFTY / BANKNIFTY / SENSEX orders found in the orderbook.")
    st.stop()

# ---- Segregation (first) ----
st.header("Segregation Pivot (Int / Pos+Int)")
stats = dor_state["pivot_stats"]
p1, p2, p3, p4 = st.columns(4)
p1.metric("Accounts", stats["accounts"])
p2.metric("Positional", stats["positional"])
p3.metric("Intraday", stats["intraday"])
p4.metric("Noren", stats["noren"])
if report_date:
    st.caption(f"Report date: {report_date}")

pivot_df = pd.DataFrame(
    [
        {
            "Section": r["Section"],
            "ALGO": r["ALGO"],
            "SERVER": r["SERVER"],
            "No. of Users": r["Users"],
            "No. of SL Hit Users": r["SLHit"],
            "MAX LOSS": r["MaxLoss"],
            "ALLOCATION": r["Allocation"],
            "Realized P&L": r["Realized"],
            "Unrealized P&L": r["Unrealized"],
            "MTM": r["MTM"],
            "Return %": r["Return"],
            "95%": r["P95"],
            "5%": r["P5"],
        }
        for r in dor_state["pivot_rows"]
    ],
    columns=PIVOT_COLS,
)
row_kinds = [r["kind"] for r in dor_state["pivot_rows"]]


def _pivot_row_style(row):
    kind = row_kinds[row.name]
    if kind == "subtotal":
        style = "background-color:#EDF0F5; color:#1F2937; font-weight:600"
    elif kind == "algototal":
        style = "background-color:#DCE3EC; color:#1F2937; font-weight:700"
    elif kind == "grandtotal":
        style = "background-color:#D1C9E1; color:#1F2937; font-weight:700"
    else:
        return [""] * len(row)
    return [style] * len(row)


def _pnl_color(value):
    """Profit green / loss red for the Realized, Unrealized and MTM columns."""
    if pd.isna(value) or value == 0:
        return ""
    return ("color:#15803D;font-weight:600" if value > 0
            else "color:#DC2626;font-weight:600")


styled = (
    pivot_df.style
    .apply(_pivot_row_style, axis=1)
    .set_properties(**{"text-align": "center"})
    .map(_pnl_color, subset=["Realized P&L", "Unrealized P&L", "MTM"])
    .format({
        "MAX LOSS": "{:,.0f}", "ALLOCATION": "{:,.0f}", "Realized P&L": "{:,.0f}",
        "Unrealized P&L": "{:,.0f}", "MTM": "{:,.0f}",
        "Return %": "{:.2f}", "95%": "{:.2f}", "5%": "{:.2f}",
    })
)
st.dataframe(styled, width="stretch", hide_index=True, height=600)

# ---- Slippage (realized loss % beyond max-loss %) ----
st.header("Slippage")
st.caption(
    "ML % = MAX LOSS / ALLOCATION; Realized ML % = |Realized P&L| / ALLOCATION — plain ratios "
    "of allocation, same convention as Return %. An account has slippage only when Realized "
    "ML % exceeds ML % by at least 0.1 (1.00 → 1.09 is not slippage; 1.10 is). Avg Slippage = "
    "average Realized ML % of the algo's slippage accounts; Major = slippage accounts above "
    "their algo's average."
)
slip_rows = seg.slippage_rows(dor_state["comp"])
slip_algo, slip_overall = seg.slippage_summary(slip_rows, dor_state["comp"])
slip_majors = seg.major_slippages(slip_rows, slip_algo)
if not slip_rows:
    st.info("-- No slippage today --")
else:
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Accounts", slip_overall["accounts"])
    k2.metric("Slippage Accounts", slip_overall["slipped"])
    k3.metric("Avg Slippage %", f"{slip_overall['avg_slippage']:.2f}")
    k4.metric("Major Slippages", len(slip_majors))
    sl1, sl2 = st.columns(2)
    with sl1:
        st.subheader("Avg Slippage per Algo")
        df_slip_sum = pd.DataFrame(
            [
                [s["ALGO"], s["accounts"], s["slipped"],
                 round(s["avg_slippage"], 2) if s["avg_slippage"] is not None else None]
                for s in slip_algo + [slip_overall]
            ],
            columns=seg.SLIP_SUMMARY_HEADERS,
        )
        st.dataframe(df_slip_sum.style.set_properties(**{"text-align": "center"}),
                     width="stretch", hide_index=True)
    with sl2:
        st.subheader("Major Slippages")
        if slip_majors:
            df_majors = pd.DataFrame(
                [
                    [seg._algo_val(r["ALGO"]), r["SERVER"], r["UserID"], r["Alias"],
                     round(r["Allocation"]), round(r["MaxLoss"]), round(r["Realized"]),
                     round(r["MLPct"], 2), round(r["RealizedMLPct"], 2),
                     round(r["DiffPct"], 2), round(r["AlgoAvgSlippage"], 2)]
                    for r in slip_majors
                ],
                columns=seg.SLIP_MAJOR_HEADERS,
            )
            st.dataframe(df_majors.style.set_properties(**{"text-align": "center"}),
                         width="stretch", hide_index=True)
        else:
            st.info("No slippage account is above its algo's average.")

# ---- Trade Value (after the pivot) ----
st.header("Trade Value")
totals = report_totals(tv_rows)
matched = len({r["user_id"] for r in tv_rows if r["allocation"] is not None})

m1, m2, m3, m4 = st.columns(4)
m1.metric("Users", totals["users"])
m2.metric("Orders", f"{totals['orders']:,}")
m3.metric("Total Lots", f"{totals['lots']:,.0f}")
m4.metric("Total Trade Value", f"{totals['trade_value']:,.2f}")

st.caption(f"Outliers: beyond {used_deviation:g} standard deviation(s) from each algo's average lot pct.")
if matched < totals["users"]:
    st.caption(
        f"Allocation matched for {matched} of {totals['users']} users "
        f"({dor_state['allocation_count']} users in the summary); unmatched users have a blank Allocation."
    )

st.subheader("Algo Summary")
st.caption(
    "All columns count users. Outliers are judged per user — combined lots across their "
    "segments/servers ÷ combined allocation — so Below + In + Above = Total Users "
    "(users with no usable allocation carry no flag)."
)
summary_rows = algo_summary(tv_rows, used_multiplier)
df_summary = pd.DataFrame([format_summary_row(s) for s in summary_rows], columns=SUMMARY_HEADER)
st.dataframe(df_summary.style.set_properties(**{"text-align": "center"}),
             width="stretch", hide_index=True)

st.subheader("Lot Pct by Algo (Box Plot)")
user_obs = user_lot_observations(tv_rows, used_multiplier)
chart = _boxplot_chart(user_obs)
if chart is not None:
    st.altair_chart(chart, use_container_width=True)

# ---- Outlier clients (stricter, second deviation) ----
st.subheader(f"Outlier Clients (±{client_deviation:g}σ)")
outlier_rows = outlier_clients(tv_rows, Decimal(str(client_deviation)))
st.caption(
    f"One row per user: those whose combined lot pct is more than {client_deviation:g} "
    "standard deviation(s) from their algo's average. Average Lots = that band converted "
    "into lots for the user's allocation; Diff of Lots = lots fired beyond the nearest band "
    "edge. Changing the deviation above updates this table instantly — no need to reprocess."
)
if outlier_rows:
    df_outliers = pd.DataFrame(
        [
            [r["algo"], r["server"], r["user_id"], format_lots_band(r),
             r["outlier"], int(round(r["lots"])), int(round(r["diff_lots"]))]
            for r in outlier_rows
        ],
        columns=OUTLIER_CLIENTS_HEADER,
    )
    st.dataframe(
        df_outliers.style.set_properties(**{"text-align": "center"}),
        width="stretch", hide_index=True,
    )
else:
    st.info(f"No clients beyond ±{client_deviation:g} standard deviation(s).")

# ---- Strikes ----
st.header("Strikes Traded")
strikes = dor_state["strikes"]
seg_counts = {}
for r in strikes["per_strike"]:
    seg_counts[r["segment"]] = seg_counts.get(r["segment"], 0) + 1
st.caption(
    f"{len(strikes['per_strike'])} distinct strikes traded"
    + (" (" + " · ".join(f"{s}: {n}" for s, n in sorted(seg_counts.items())) + ")" if seg_counts else "")
    + " — a strike is one contract: index + expiry + strike + CE/PE; "
    "both symbol formats (e.g. NIFTY21JUL26… and NIFTY26721…) count as one."
)
chain = strike_chain(strikes["per_strike"])
expiries = sorted({expiry for _, expiry in chain})
expiry_labels = [e.strftime("%d%b%y").upper() for e in expiries]
segments = sorted({seg for seg, _ in chain})

s1, s2 = st.columns(2)
with s1:
    st.subheader("Strikes per Algo / Server")
    f1, f2 = st.columns(2)
    with f1:
        as_expiry = st.selectbox("Expiry", ["All"] + expiry_labels, key="as_expiry")
    with f2:
        as_index = st.selectbox("Index", ["All"] + segments, key="as_index")
    as_rows, as_distinct = [], set()
    for r in strikes["by_algo_server"]:
        selected = [
            c for c in r["contracts"]
            if (as_expiry == "All" or c[1].strftime("%d%b%y").upper() == as_expiry)
            and (as_index == "All" or c[0] == as_index)
        ]
        if selected:
            as_rows.append([r["algo"], r["server"], len(selected)])
            as_distinct.update(selected)
    df_algo_strikes = pd.DataFrame(as_rows, columns=STRIKE_ALGO_HEADER)
    st.dataframe(df_algo_strikes.style.set_properties(**{"text-align": "center"}),
                 width="stretch", hide_index=True)
    st.caption(f"Total (distinct) strikes in this view: {len(as_distinct):,}")
with s2:
    st.subheader("Lots per Strike — Option Chain")
    c_exp, c_idx = st.columns(2)
    with c_exp:
        expiry_label = st.selectbox("Expiry", expiry_labels, key="chain_expiry")
    expiry = expiries[expiry_labels.index(expiry_label)]
    with c_idx:
        index = st.selectbox("Index", sorted(seg for seg, e in chain if e == expiry),
                             key="chain_index")
    chain_rows = chain.get((index, expiry), [])
    df_chain = pd.DataFrame(
        [[int(round(ce)), strike, int(round(pe))] for ce, strike, pe in chain_rows],
        columns=STRIKE_CHAIN_HEADER,
    )
    st.dataframe(
        df_chain.style.set_properties(**{"text-align": "center"}),
        width="stretch", hide_index=True,
        column_config={
            "CE": st.column_config.NumberColumn(format="localized"),
            "Strike": st.column_config.NumberColumn(format="plain"),
            "PE": st.column_config.NumberColumn(format="localized"),
        },
    )
    ce_total = sum(int(round(ce)) for ce, _, _ in chain_rows)
    pe_total = sum(int(round(pe)) for _, _, pe in chain_rows)
    st.caption(f"Total lots — CE: {ce_total:,} · PE: {pe_total:,} · both: {ce_total + pe_total:,}")

st.subheader("Trade Value Rows")
df = pd.DataFrame([format_row(r) for r in tv_rows], columns=REPORT_HEADER)
st.dataframe(df.style.set_properties(**{"text-align": "center"}),
             width="stretch", hide_index=True)

# ---------------------------------------------------------------------------
# Downloads — one Excel (all 5 sheets) + the shareable DOR.html
# ---------------------------------------------------------------------------
st.header("Downloads")

workbook = Workbook()
workbook.remove(workbook.active)
add_report_sheets(workbook, tv_rows, used_multiplier)
add_strikes_sheet(workbook, strikes)
add_outlier_clients_sheet(workbook, outlier_rows)
seg.add_segregation_sheets(workbook, dor_state["comp"], report_date)
excel_buffer = io.BytesIO()
workbook.save(excel_buffer)

dor_html = build_dor_html(
    report_date=report_date or "—",
    deviation=float(used_deviation),
    tv_totals=totals,
    tv_summary=summary_rows,
    pivot_stats=stats,
    pivot_rows=dor_state["pivot_rows"],
    boxplot_rows=user_obs,
    strikes=strikes,
    outliers=outlier_rows,
    outlier_deviation=float(client_deviation),
    slippage={"summary": slip_algo, "overall": slip_overall, "majors": slip_majors},
    excel_bytes=excel_buffer.getvalue(),
    excel_filename=f"DOR_{report_date or 'report'}.xlsx",
)

d1, d2 = st.columns(2)
with d1:
    st.download_button(
        "Download Excel — all data (8 sheets)",
        data=excel_buffer.getvalue(),
        file_name=f"DOR_{report_date or 'report'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption("tradevalue · summary · Strikes · Outlier Clients · Segregation · "
               "Raw_Data_Per_User · Worst 10%ile · Slippage")
with d2:
    st.download_button(
        "Download DOR.html — shareable summary report",
        data=dor_html.encode("utf-8"),
        file_name=f"DOR_{report_date or 'report'}.html",
        mime="text/html",
        type="primary",
    )
    st.caption("Self-contained styled report (summary only) — share with any user or client.")
