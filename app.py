"""Daily Operations Report (DOR) — Streamlit UI.

Run with:  streamlit run app.py

All inputs sit in one block at the top: the compiled orderbook, the compiled
user MTM, the Combined Max Loss file(s) and the outlier deviation. One
"Process" click computes BOTH reports — the trade value analysis and the
Int / Pos+Int segregation pivot — and offers exactly two downloads:

  * DOR_<date>.xlsx — all 5 data sheets in one workbook
    (tradevalue, summary, Segregation, Raw_Data_Per_User, Worst 10%ile)
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
    REPORT_HEADER,
    SUMMARY_HEADER,
    add_report_sheets,
    aggregate,
    algo_summary,
    dedup_orders,
    format_row,
    format_summary_row,
    read_allocations,
    read_orderbook,
    report_totals,
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

dev_col, btn_col = st.columns([1, 3])
with dev_col:
    deviation = st.number_input(
        "Outlier deviation (× std dev)",
        min_value=0.1, max_value=10.0, value=1.0, step=0.5,
        help="A user is an outlier when their lot pct is more than this many "
             "standard deviations away from their algo's average lot pct.",
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
    tv_rows = aggregate(dedup_orders(orders), allocations, multiplier)

    # Segregation
    four = seg.load_combined(io.BytesIO(four_bytes)) if four_bytes else {}
    one = seg.load_combined(io.BytesIO(one_bytes))
    comp = seg.prepare_comp(io.BytesIO(mtm_bytes), four, one)
    report_date = seg.infer_report_date(comp)

    return {
        "tv_rows": tv_rows,
        "order_count": len(orders),
        "allocation_count": len(allocations),
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
                tuple(REPORT_HEADER),
            )
        st.session_state["dor"] = {**result, "deviation": deviation, "signature": _signature()}
    except Exception as exc:
        st.error(f"Processing failed: {exc}")


def _boxplot_chart(rows):
    """Box plot of lot pct per algo. The boxes show the distribution; the
    std-dev-flagged outliers are overlaid as points (blue below, orange
    above — a colorblind-safe pair) with user details on hover."""
    records = [
        {
            "Algo": str(r["algo"]),
            "Lot Pct": round(float(r["lot_pct"]), 4),
            "User ID": r["user_id"],
            "Server": r["server"],
            "Segment": r["segment"],
            "Lots": round(float(r["lots"]), 2),
            "Outlier": r["outlier"],
        }
        for r in rows
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
            alt.Tooltip("Segment:N"),
            alt.Tooltip("Algo:N"),
            alt.Tooltip("Lot Pct:Q", format=".2f"),
            alt.Tooltip("Lots:Q", format=",.2f"),
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


styled = (
    pivot_df.style
    .apply(_pivot_row_style, axis=1)
    .format({
        "MAX LOSS": "{:,.0f}", "ALLOCATION": "{:,.0f}", "Realized P&L": "{:,.0f}",
        "Unrealized P&L": "{:,.0f}", "MTM": "{:,.0f}",
        "Return %": "{:.2f}", "95%": "{:.2f}", "5%": "{:.2f}",
    })
)
st.dataframe(styled, width="stretch", hide_index=True, height=600)

# ---- Trade Value (after the pivot) ----
st.header("Trade Value")
totals = report_totals(tv_rows)
matched = len({r["user_id"] for r in tv_rows if r["allocation"] is not None})

m1, m2, m3, m4 = st.columns(4)
m1.metric("Users", totals["users"])
m2.metric("Orders", f"{totals['orders']:,}")
m3.metric("Total Lots", f"{totals['lots']:,.2f}")
m4.metric("Total Trade Value", f"{totals['trade_value']:,.2f}")

st.caption(f"Outliers: beyond {used_deviation:g} standard deviation(s) from each algo's average lot pct.")
if matched < totals["users"]:
    st.caption(
        f"Allocation matched for {matched} of {totals['users']} users "
        f"({dor_state['allocation_count']} users in the summary); unmatched users have a blank Allocation."
    )

st.subheader("Algo Summary")
summary_rows = algo_summary(tv_rows, used_multiplier)
df_summary = pd.DataFrame([format_summary_row(s) for s in summary_rows], columns=SUMMARY_HEADER)
st.dataframe(df_summary, width="stretch", hide_index=True)

st.subheader("Lot Pct by Algo (Box Plot)")
chart = _boxplot_chart(tv_rows)
if chart is not None:
    st.altair_chart(chart, use_container_width=True)

st.subheader("Trade Value Rows")
df = pd.DataFrame([format_row(r) for r in tv_rows], columns=REPORT_HEADER)
st.dataframe(df, width="stretch", hide_index=True)

# ---------------------------------------------------------------------------
# Downloads — one Excel (all 5 sheets) + the shareable DOR.html
# ---------------------------------------------------------------------------
st.header("Downloads")

workbook = Workbook()
workbook.remove(workbook.active)
add_report_sheets(workbook, tv_rows, used_multiplier)
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
    tv_rows=tv_rows,
)

d1, d2 = st.columns(2)
with d1:
    st.download_button(
        "Download Excel — all data (5 sheets)",
        data=excel_buffer.getvalue(),
        file_name=f"DOR_{report_date or 'report'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption("tradevalue · summary · Segregation · Raw_Data_Per_User · Worst 10%ile")
with d2:
    st.download_button(
        "Download DOR.html — shareable summary report",
        data=dor_html.encode("utf-8"),
        file_name=f"DOR_{report_date or 'report'}.html",
        mime="text/html",
        type="primary",
    )
    st.caption("Self-contained styled report (summary only) — share with any user or client.")
