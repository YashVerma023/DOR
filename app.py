"""Daily Operations Report (DOR) — Streamlit UI.

Run with:  streamlit run app.py

All inputs sit in one block at the top: the compiled orderbook, the compiled
user MTM, the Combined Max Loss file(s) and the outlier deviation. One
"Process" click computes BOTH reports — the trade value analysis and the
Int / Pos+Int segregation pivot — and offers exactly two downloads:

  * DOR_<date>.xlsx — all data sheets in one workbook (tradevalue, summary,
    Strikes, Portfolio QS when the MLOB is given, Segregation,
    Raw_Data_Per_User, Worst 10%ile, Slippage)
  * DOR_<date>.html — the styled, client-shareable summary report
"""

import io
from decimal import Decimal

import pandas as pd
import streamlit as st
from openpyxl import Workbook

import segregate_int_pos_mtm2 as seg
from dor import build_dor_html
from portfolio import (
    DEFAULT_PORTFOLIO_PATTERN,
    PORTFOLIO_SUMMARY_HEADER,
    add_portfolio_sheet,
    portfolio_groups,
    portfolio_report,
    read_mlob,
)
from tradevalue import (
    REPORT_HEADER,
    STRIKE_ALGO_HEADER,
    STRIKE_CHAIN_HEADER,
    SUMMARY_HEADER,
    add_report_sheets,
    add_strikes_sheet,
    aggregate,
    algo_summary,
    dedup_orders,
    format_crore,
    format_indian,
    format_row,
    format_summary_row,
    read_allocations,
    read_orderbook,
    report_totals,
    strike_chain,
    strike_report,
    user_lot_observations,
)
from tradevalue import _user_key as tv_user_key

st.set_page_config(page_title="Daily Operations Report", layout="wide")
st.title("Daily Operations Report")
st.caption(
    "Upload the four inputs (plus the optional MLOB for the Portfolio Analysis), set the "
    "outlier deviations, and click **Process** — the Trade Value analysis, the Portfolio "
    "Analysis and the Int / Pos+Int Segregation are computed in one go, with a single Excel "
    "workbook and a shareable DOR.html as output."
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

in5, _ = st.columns(2)
with in5:
    mlob_file = st.file_uploader(
        "Multileg Orders — MLOB (optional, enables the Portfolio Analysis)",
        type=["xlsx", "xlsm", "xls", "csv"],
    )

oc1, oc2, oc3, oc4 = st.columns(4)
with oc1:
    nifty_open = st.number_input("NIFTY day open", min_value=0.0, value=0.0, step=50.0,
                                 help="0 = not set. With day close, gives the day-mid "
                                      "used as the option chain's ATM.")
with oc2:
    nifty_close = st.number_input("NIFTY day close", min_value=0.0, value=0.0, step=50.0)
with oc3:
    sensex_open = st.number_input("SENSEX day open", min_value=0.0, value=0.0, step=100.0)
with oc4:
    sensex_close = st.number_input("SENSEX day close", min_value=0.0, value=0.0, step=100.0)

@st.cache_data(show_spinner=False)
def _algo_options_from_mtm(mtm_bytes, mtm_name):
    """The algo choices for the segregation / slippage multiselects, read
    straight from the uploaded MTM so they are pickable BEFORE the first
    Process."""
    try:
        if str(mtm_name).lower().endswith(".csv"):
            df = pd.read_csv(io.BytesIO(mtm_bytes))
        else:
            df = pd.read_excel(io.BytesIO(mtm_bytes), sheet_name=0)
        algo_col = next((c for c in df.columns if str(c).strip().upper() == "ALGO"), None)
        if algo_col is None:
            return []
        algos = df[algo_col].dropna().unique()
        try:
            algos = sorted(algos, key=float)
        except (TypeError, ValueError):
            algos = sorted(algos, key=str)
        return [seg._algo_val(a) for a in algos]
    except Exception:
        return []


_known_algos = []
if summary_file is not None:
    _known_algos = _algo_options_from_mtm(summary_file.getvalue(), summary_file.name)
if not _known_algos:
    _known_algos = st.session_state.get("dor", {}).get("slip_algo_options", [])

dev_col, seg_col, slip_col, btn_col = st.columns([1, 1.5, 1.5, 1])
with dev_col:
    deviation = st.number_input(
        "Outlier deviation (× MAD)",
        min_value=0.1, max_value=10.0, value=1.0, step=0.5,
        help="A user is an outlier when their Lots per Cr is more than this "
             "many robust deviations (MAD) away from their algo's median.",
    )
with seg_col:
    seg_algos_input = st.multiselect(
        "Segregation — algos to include (empty = all)",
        options=_known_algos, default=_known_algos, key="seg_algos_input",
        help="Which algos the Int / Pos+Int segregation pivot and its KPI "
             "summary cover — dashboard and DOR.html. Populated as soon as "
             "the User MTM is uploaded; empty = every algo.",
    )
with slip_col:
    slip_algos_input = st.multiselect(
        "Slippage — algos to analyse (empty = all)",
        options=_known_algos, default=_known_algos, key="slip_algos_input",
        help="Populated as soon as the User MTM is uploaded; leaving it empty "
             "analyses every algo. Applies to the dashboard AND the DOR.html report.",
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
                 mlob_bytes, mlob_name, deviation_value, schema):
    # `schema` only serves as part of the cache key: when the report columns
    # change, stale cached rows from an older code version are not reused.
    multiplier = Decimal(str(deviation_value))

    # Segregation first — its Int / Positional classification also splits the
    # trade value users
    four = seg.load_combined(io.BytesIO(four_bytes)) if four_bytes else {}
    one = seg.load_combined(io.BytesIO(one_bytes))
    comp = seg.prepare_comp(io.BytesIO(mtm_bytes), four, one)
    report_date = seg.infer_report_date(comp)
    servers = (comp["SERVER"].map(seg._server_key) if "SERVER" in comp.columns
               else [""] * len(comp))
    type_map = {
        (tv_user_key(uid), srv): ("Int" if t == "Intraday" else "Pos+Int")
        for uid, srv, t in zip(comp["UserID"], servers, comp["Type"])
    }

    # Trade value
    orders = read_orderbook(io.BytesIO(ob_bytes), ob_name)
    allocations = read_allocations(io.BytesIO(mtm_bytes), mtm_name)
    deduped = dedup_orders(orders)
    tv_rows = aggregate(deduped, allocations, multiplier, type_map)
    strikes = strike_report(deduped, allocations)

    # Portfolio analysis (optional MLOB)
    pf_groups = None
    if mlob_bytes:
        mlob = read_mlob(io.BytesIO(mlob_bytes), mlob_name)
        pf_groups = portfolio_groups(mlob, allocations)

    return {
        "tv_rows": tv_rows,
        "order_count": len(orders),
        "allocation_count": len(allocations),
        "strikes": strikes,
        "portfolio_groups": pf_groups,
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
        mlob_file.name if mlob_file else None,
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
                mlob_file.getvalue() if mlob_file else None,
                mlob_file.name if mlob_file else None,
                deviation,
                tuple(REPORT_HEADER + STRIKE_ALGO_HEADER + STRIKE_CHAIN_HEADER
                      + PORTFOLIO_SUMMARY_HEADER + ["chain-breakdown-v1"]),
            )
        st.session_state["dor"] = {
            **result, "deviation": deviation, "signature": _signature(),
            "slip_algo_options": [seg._algo_val(a) for a in
                                  sorted(result["comp"]["ALGO"].dropna().unique(), key=float)],
        }
    except Exception as exc:
        st.error(f"Processing failed: {exc}")


PIVOT_COLS = ["Section", "ALGO", "SERVER", "No. of Users", "MAX LOSS", "ALLOCATION",
              "Realized P&L", "P&L %", "No. of SL Hit Users"]

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
algo_options = dor_state.get("slip_algo_options") or [
    seg._algo_val(a) for a in sorted(dor_state["comp"]["ALGO"].dropna().unique(), key=float)]
picked_seg_algos = seg_algos_input or algo_options
comp_seg = dor_state["comp"][
    dor_state["comp"]["ALGO"].map(seg._algo_val).isin(picked_seg_algos)]
stats = seg.comp_stats(comp_seg)
pivot_rows_view = seg.pivot_rows(comp_seg)
grand_view = next(r for r in pivot_rows_view if r["kind"] == "grandtotal")
p1, p2, p3, p4, p5, p6 = st.columns(6)
p1.metric("Accounts", stats["accounts"])
p2.metric("Positional", stats["positional"])
p3.metric("Intraday", stats["intraday"])
p4.metric("Allocation", format_crore(grand_view["Allocation"] * 100))
p5.metric("Realized P&L", format_crore(grand_view["Realized"]))
p6.metric("P&L %", f"{grand_view['Return']:.2f}")
st.caption(
    (f"Report date: {report_date} · " if report_date else "")
    + "Algos included (from the inputs at the top): "
    + ", ".join(str(a) for a in picked_seg_algos)
)

pivot_df = pd.DataFrame(
    [
        {
            "Section": r["Section"],
            "ALGO": r["ALGO"],
            "SERVER": r["SERVER"],
            "No. of Users": r["Users"],
            "MAX LOSS": r["MaxLoss"],
            # display convention: the stored allocation is in hundreds
            "ALLOCATION": r["Allocation"] * 100,
            "Realized P&L": r["Realized"],
            "P&L %": r["Return"],
            "No. of SL Hit Users": r["SLHit"],
        }
        for r in pivot_rows_view
    ],
    columns=PIVOT_COLS,
)
row_kinds = [r["kind"] for r in pivot_rows_view]


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
    .map(_pnl_color, subset=["Realized P&L"])
    .format({
        "MAX LOSS": format_crore, "ALLOCATION": format_crore,
        "Realized P&L": format_crore, "P&L %": "{:.2f}",
    })
)
st.dataframe(styled, width="stretch", hide_index=True, height=600)

# ---- Slippage (realized loss % beyond max-loss %) ----
st.header("Max SL Slippage Analysis")
st.caption(
    "ML % = MAX LOSS / ALLOCATION; Realized ML % = |Realized P&L| / ALLOCATION — plain ratios "
    "of allocation, same convention as Return %. An account has slippage only when Realized "
    "ML % exceeds ML % by at least 0.1 (1.00 → 1.09 is not slippage; 1.10 is). Avg Slippage = "
    "average Realized ML % of the algo's slippage accounts; Major = slippage accounts above "
    "their algo's average."
)
all_slip_algos = dor_state.get("slip_algo_options") or [
    seg._algo_val(a) for a in sorted(dor_state["comp"]["ALGO"].dropna().unique(), key=float)]
picked_slip_algos = slip_algos_input or all_slip_algos
st.caption(f"Algos analysed (from the inputs at the top): "
           f"{', '.join(str(a) for a in picked_slip_algos)}")
comp_sel = dor_state["comp"][
    dor_state["comp"]["ALGO"].map(seg._algo_val).isin(picked_slip_algos)]
slip_rows = seg.slippage_rows(comp_sel)
slip_algo, slip_overall = seg.slippage_summary(slip_rows, comp_sel)
slip_majors = seg.major_slippages(slip_rows, slip_algo)
slip_empty_note = ("-- No slippage today --"
                   if len(picked_slip_algos) == len(all_slip_algos)
                   else "-- No slippage in the selected algos --")
if not slip_rows:
    st.info(slip_empty_note)
else:
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Accounts", format_indian(slip_overall["accounts"]))
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
                     format_indian(r["Allocation"]), format_crore(r["MaxLoss"]),
                     format_indian(r["Realized"]),
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
m1.metric("Users", format_indian(totals["users"]))
m2.metric("Orders", format_indian(totals["orders"]))
m3.metric("Total Lots", format_indian(totals["lots"]))
m4.metric("Total Trade Value", format_crore(totals["trade_value"]))

st.caption(f"Outliers: beyond {used_deviation:g} robust deviation(s) (MAD) from each algo + type group's median Lots per Cr.")
if matched < totals["users"]:
    st.caption(
        f"Allocation matched for {matched} of {totals['users']} users "
        f"({dor_state['allocation_count']} users in the summary); unmatched users have a blank Allocation."
    )

st.subheader("Algo Summary")
st.caption(
    "All columns count users. Outliers are judged per user on **Lots per Cr** — combined "
    "lots ÷ normalise, where normalise = allocation / 1,00,000 (sub-1-Cr allocations "
    "normalise fractionally: 80,000 → 0.8, … 20,000 → 0.2) — so Below + In + Above = "
    "Total Users (only users with no usable allocation carry no flag). Pick an algo below "
    "to see its outlier user ids."
)
summary_rows = algo_summary(tv_rows, used_multiplier)
user_obs = user_lot_observations(tv_rows, used_multiplier)
df_summary = pd.DataFrame([format_summary_row(s) for s in summary_rows], columns=SUMMARY_HEADER)
st.dataframe(df_summary.style.set_properties(**{"text-align": "center"}),
             width="stretch", hide_index=True)

_summary_labels = [f"Algo {s['algo']}" + (f" · {s['user_type']}" if s["user_type"] else "")
                   for s in summary_rows]
algo_pick = st.selectbox(
    "Outlier users — pick an algo / type",
    ["—"] + _summary_labels,
    key="summary_drill",
)
if algo_pick != "—":
    picked = summary_rows[_summary_labels.index(algo_pick)]
    flagged = [o for o in user_obs
               if o["algo"] == picked["algo"] and o["trade_date"] == picked["trade_date"]
               and o["user_type"] == picked["user_type"]
               and o["outlier"] in ("Below average range", "Above average range")]
    flagged.sort(key=lambda o: ((0, o["lots_per_cr"]) if o["outlier"].startswith("Below")
                                else (1, -o["lots_per_cr"])))
    if flagged:
        df_flagged = pd.DataFrame(
            [[o["user_id"], o["server"], round(float(o["lots_per_cr"]), 2),
              format_indian(o["lots"]), o["outlier"]] for o in flagged],
            columns=["User ID", "Server", "Lots per Cr", "Lots", "Outlier"],
        )
        st.dataframe(df_flagged.style.set_properties(**{"text-align": "center"}),
                     width="stretch", hide_index=True)
    else:
        st.info("No outlier users in this algo.")

# ---- Portfolio analysis (from the MLOB) ----
def _centered(df):
    return (df.style.set_properties(**{"text-align": "center"})
            .map(_pnl_color, subset=["PnL"])
            .format({"PnL": format_indian, "Total Orders": format_indian}))


def _portfolio_tables(report, key_prefix):
    """Three views of one portfolio report: the algo summary (default), a
    per-server view for a chosen algo, and a per-user view for a chosen
    server."""
    if not report["algos"]:
        st.info(f"No portfolio name contains \"{report['pattern']}\".")
        return
    total = report["total"]
    df_algos = pd.DataFrame(
        [[a["algo"], a["n_servers"], a["users"], a["portfolios"], a["orders"], round(a["pnl"])]
         for a in report["algos"]]
        + [["Total", "", total["users"], total["portfolios"], total["orders"], round(total["pnl"])]],
        columns=["Algo", "No. of Server", "Portfolio Executed Users", "No. of Portfolio",
                 "Total Orders", "PnL"],
    )
    st.dataframe(_centered(df_algos), width="stretch", hide_index=True)

    d1, d2 = st.columns(2)
    with d1:
        algo_choice = st.selectbox(
            "Server view — pick an algo",
            ["—"] + [f"Algo {a['algo']}" for a in report["algos"]],
            key=f"{key_prefix}_algo",
        )
    if algo_choice == "—":
        return
    algo_row = report["algos"][[f"Algo {a['algo']}" for a in report["algos"]].index(algo_choice)]
    df_servers = pd.DataFrame(
        [[algo_row["algo"], s["server"], s["users"], s["portfolios"], s["orders"], round(s["pnl"])]
         for s in algo_row["server_rows"]],
        columns=PORTFOLIO_SUMMARY_HEADER,
    )
    st.dataframe(_centered(df_servers), width="stretch", hide_index=True)
    with d2:
        server_choice = st.selectbox(
            "User view — pick a server",
            ["—"] + [s["server"] for s in algo_row["server_rows"]],
            key=f"{key_prefix}_server",
        )
    if server_choice == "—":
        return
    server_row = next(s for s in algo_row["server_rows"] if s["server"] == server_choice)
    df_users = pd.DataFrame(
        [[u["user_id"], u["portfolios"], u["orders"], round(u["pnl"])] for u in server_row["user_rows"]],
        columns=["User ID", "No. of Portfolio", "Total Orders", "PnL"],
    )
    st.dataframe(_centered(df_users), width="stretch", hide_index=True)


pf_groups = dor_state.get("portfolio_groups")
if pf_groups is not None:
    st.header("Portfolio Analysis")
    st.caption(
        "From the Multileg Orders (MLOB), COMPLETED Orders only. "
        "PnL = Σ sell (Avg Price × Filled Qty) − Σ buy (Avg Price × Filled Qty). "
        "The algo comes from the User MTM matching; users with no MTM entry inherit "
        "their server's algo. Total Orders = the COMPLETE entries of the portfolios."
    )
    st.subheader(f"{DEFAULT_PORTFOLIO_PATTERN} Portfolio Analysis")
    _portfolio_tables(portfolio_report(pf_groups, DEFAULT_PORTFOLIO_PATTERN), "pf_qs")
    pc1, pc2 = st.columns(2)
    with pc1:
        picked_name = st.selectbox(
            "Analyze another portfolio — type to filter the names",
            ["—"] + sorted({g["portfolio"] for g in pf_groups}),
            key="pf_pick",
        )
    with pc2:
        custom_pattern = st.text_input(
            "…or analyse a name fragment (aggregates every match)",
            key="pf_custom", placeholder="e.g. WT3% or PPP — leave empty to skip",
        ).strip()
    if picked_name != "—":
        st.subheader(f"{picked_name} Portfolio Analysis")
        _portfolio_tables(portfolio_report(pf_groups, picked_name), "pf_pick_tbl")
    if custom_pattern:
        st.subheader(f"{custom_pattern} Portfolio Analysis")
        _portfolio_tables(portfolio_report(pf_groups, custom_pattern), "pf_custom_tbl")

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

# day-mid per index from the open/close inputs -> the chain's ATM
mids = {}
if nifty_open > 0 and nifty_close > 0:
    mids["NIFTY"] = (nifty_open + nifty_close) / 2
if sensex_open > 0 and sensex_close > 0:
    mids["SENSEX"] = (sensex_open + sensex_close) / 2

s1, s2 = st.columns(2)
with s1:
    st.subheader("Strikes per Algo / Server")
    f1, f2 = st.columns(2)
    with f1:
        as_expiry = st.selectbox("Expiry", ["All"] + expiry_labels, key="as_expiry")
    with f2:
        as_index = st.selectbox("Index", ["All"] + segments, key="as_index")
    by_algo, as_distinct = {}, set()
    for r in strikes["by_algo_server"]:
        selected = [
            c for c in r["contracts"]
            if (as_expiry == "All" or c[1].strftime("%d%b%y").upper() == as_expiry)
            and (as_index == "All" or c[0] == as_index)
        ]
        if selected:
            bucket = by_algo.setdefault(r["algo"], {"servers": [], "contracts": set()})
            bucket["servers"].append([r["server"], len(selected)])
            bucket["contracts"].update(selected)
            as_distinct.update(selected)
    df_algo_strikes = pd.DataFrame(
        [[algo, len(b["servers"]), len(b["contracts"])] for algo, b in by_algo.items()],
        columns=["Algo", "No. of Server", "Strikes Traded"],
    )
    st.dataframe(df_algo_strikes.style.set_properties(**{"text-align": "center"}),
                 width="stretch", hide_index=True)
    st.caption(f"Total (distinct) strikes in this view: {format_indian(len(as_distinct))} — "
               "an algo's count is distinct across its servers, not the column sum.")
    strike_algo_pick = st.selectbox(
        "Server view — pick an algo",
        ["—"] + [f"Algo {algo}" for algo in by_algo],
        key="as_algo_drill",
    )
    if strike_algo_pick != "—":
        picked_algo = list(by_algo)[[f"Algo {a}" for a in by_algo].index(strike_algo_pick)]
        df_servers = pd.DataFrame(
            [[picked_algo, server, n] for server, n in sorted(by_algo[picked_algo]["servers"])],
            columns=STRIKE_ALGO_HEADER,
        )
        st.dataframe(df_servers.style.set_properties(**{"text-align": "center"}),
                     width="stretch", hide_index=True)
with s2:
    st.subheader("Lots per Strike — Option Chain")
    c_exp, c_idx, c_n = st.columns(3)
    with c_exp:
        expiry_label = st.selectbox("Expiry", expiry_labels, key="chain_expiry")
    expiry = expiries[expiry_labels.index(expiry_label)]
    with c_idx:
        index = st.selectbox("Index", sorted(seg for seg, e in chain if e == expiry),
                             key="chain_index")
    with c_n:
        n_around = st.number_input("No. of strikes", min_value=1, value=10, key="chain_n",
                                   help="Strikes shown above and below the ATM — applies "
                                        "when the index has a day open/close entered.")
    f_algo, f_h, f_v, f_s = st.columns([1.4, 1, 1, 1])
    breakdown_algos = sorted(
        {str(a) or "—" for r in strikes["per_strike"] for (a, _c) in r.get("breakdown", {})},
        key=lambda a: (0, int(a)) if a.isdigit() else (1, a))
    with f_algo:
        chain_algo = st.selectbox("Algo", ["All"] + breakdown_algos, key="chain_algo")
    with f_h:
        inc_hedge = st.checkbox("Hedge", value=True, key="chain_hedge")
    with f_v:
        inc_var = st.checkbox("VAR", value=True, key="chain_var")
    with f_s:
        inc_sqoff = st.checkbox("Sq-off", value=True, key="chain_sqoff")
    include = {"normal"}
    if inc_hedge:
        include.add("hedge")
    if inc_var:
        include.add("var")
    if inc_sqoff:
        include.add("sqoff")
    by_strike = {}
    for r in strikes["per_strike"]:
        if r["segment"] != index or r["expiry"] != expiry:
            continue
        total = sum(float(lots) for (a, cat), lots in r.get("breakdown", {}).items()
                    if cat in include
                    and (chain_algo == "All" or (str(a) or "—") == chain_algo))
        if total:
            cell = by_strike.setdefault(r["strike"], [0.0, 0.0])
            cell[0 if r["opt_type"] == "CE" else 1] += total
    chain_rows = [(ce, strike, pe) for strike, (ce, pe) in sorted(by_strike.items())]
    mid = mids.get(index)
    if mid and chain_rows:
        atm_i = min(range(len(chain_rows)), key=lambda i: abs(chain_rows[i][1] - mid))
        atm_strike = chain_rows[atm_i][1]
        chain_rows = chain_rows[max(0, atm_i - n_around):atm_i + n_around + 1]
        st.caption(f"Centred on ATM **{atm_strike}** (day mid {mid:g}) — "
                   f"±{n_around} strikes shown.")
    df_chain = pd.DataFrame(
        [[int(round(ce)), strike, int(round(pe))] for ce, strike, pe in chain_rows],
        columns=STRIKE_CHAIN_HEADER,
    )
    st.dataframe(
        df_chain.style.set_properties(**{"text-align": "center"})
        .format({"CE": format_indian, "PE": format_indian}),
        width="stretch", hide_index=True,
        column_config={"Strike": st.column_config.NumberColumn(format="plain")},
    )
    ce_total = sum(int(round(ce)) for ce, _, _ in chain_rows)
    pe_total = sum(int(round(pe)) for _, _, pe in chain_rows)
    st.caption(f"Total lots — CE: {format_indian(ce_total)} · PE: {format_indian(pe_total)} "
               f"· both: {format_indian(ce_total + pe_total)}")

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
if pf_groups is not None:
    add_portfolio_sheet(workbook, portfolio_report(pf_groups, DEFAULT_PORTFOLIO_PATTERN))
seg.add_segregation_sheets(workbook, dor_state["comp"], report_date)
excel_buffer = io.BytesIO()
workbook.save(excel_buffer)

# the report's slippage section carries the algo selection made in the
# inputs (same as the dashboard); the chain centres on the entered day-mids
dor_html = build_dor_html(
    report_date=report_date or "—",
    deviation=float(used_deviation),
    tv_totals=totals,
    tv_summary=summary_rows,
    pivot_stats=stats,
    pivot_rows=pivot_rows_view,
    user_obs=user_obs,
    strikes=strikes,
    slippage={"summary": slip_algo, "overall": slip_overall, "majors": slip_majors,
              "empty_note": slip_empty_note},
    portfolio=pf_groups,
    mids=mids,
    excel_bytes=excel_buffer.getvalue(),
    excel_filename=f"DOR_{report_date or 'report'}.xlsx",
)

d1, d2 = st.columns(2)
with d1:
    st.download_button(
        "Download Excel — all data",
        data=excel_buffer.getvalue(),
        file_name=f"DOR_{report_date or 'report'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption("tradevalue · summary · Strikes"
               + (" · Portfolio QS" if pf_groups is not None else "")
               + " · Segregation · Raw_Data_Per_User · Worst 10%ile · Slippage")
with d2:
    st.download_button(
        "Download DOR.html — shareable summary report",
        data=dor_html.encode("utf-8"),
        file_name=f"DOR_{report_date or 'report'}.html",
        mime="text/html",
        type="primary",
    )
    st.caption("Self-contained styled report (summary only) — share with any user or client.")
