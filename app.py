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
    split_report_rows,
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
# Mode — 0DTE is the full pipeline; Other drops the max-loss inputs
# ---------------------------------------------------------------------------
mode = st.sidebar.radio(
    "Report mode", ["0DTE", "Other"],
    help="0DTE — the full report: upload the Combined Max Loss file(s) so accounts "
         "are split Int / Pos+Int and positional Realized P&L carries its addon.\n\n"
         "Other — same inputs without the max-loss files: no Int / Pos+Int split "
         "and no addon.",
)
st.sidebar.caption(
    "**0DTE** uses the Combined Max Loss files for the Int / Pos+Int split and the "
    "Realized P&L addon.\n\n**Other** runs without them — one unsplit block, "
    "Realized P&L exactly as the MTM sheet reports it."
    if mode == "0DTE" else
    "**Other** mode: no max-loss files, so there is no Int / Pos+Int split and no "
    "addon — Realized P&L is exactly what the MTM sheet reports."
)

# ---------------------------------------------------------------------------
# Inputs — everything at the top
# ---------------------------------------------------------------------------
in1, in2 = st.columns(2)
with in1:
    orderbook_file = st.file_uploader("Orderbook (CSV / Excel)", type=["csv", "xlsx", "xlsm", "xls"])
with in2:
    summary_file = st.file_uploader("Compiled User MTM (CSV / Excel)", type=["csv", "xlsx", "xlsm", "xls"])

if mode == "0DTE":
    in3, in4 = st.columns(2)
    with in3:
        maxloss_1dte = st.file_uploader("Combined Max Loss — 1DTE (required)",
                                        type=["xlsx", "xlsm", "xls"])
    with in4:
        maxloss_4dte = st.file_uploader("Combined Max Loss — 4DTE (optional)",
                                        type=["xlsx", "xlsm", "xls"])
else:
    # "Other" mode runs without the max-loss files: no Int / Pos+Int
    # classification and no realized-P&L addon.
    maxloss_1dte = maxloss_4dte = None

in5, in6 = st.columns(2)
with in5:
    mlob_file = st.file_uploader(
        "Multileg Orders — MLOB (optional, enables the Portfolio Analysis)",
        type=["xlsx", "xlsm", "xls", "csv"],
    )
with in6:
    summary2_file = st.file_uploader(
        "Secondary index User MTM (optional)",
        type=["csv", "xlsx", "xlsm", "xls"],
        help="When two indexes ran on different servers (e.g. 8 NIFTY servers "
             "+ 2 BANKNIFTY servers), the orderbook is combined but the User "
             "MTM comes as two files — upload the second one here. The summary "
             "cards / pivot and the outlier (std deviation) tables are then "
             "shown separately per MTM file.",
    )

oc1, oc2, oc3, oc4, oc5, oc6 = st.columns(6)
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
with oc5:
    banknifty_open = st.number_input("BANKNIFTY day open", min_value=0.0, value=0.0,
                                     step=100.0)
with oc6:
    banknifty_close = st.number_input("BANKNIFTY day close", min_value=0.0, value=0.0,
                                      step=100.0)

# the expiry is NOT parsed from the strike symbols (their formats are too
# ambiguous to trust) — every symbol of an index shares the session's single
# expiry, entered here and applied as a label in the strikes section
ex1, ex2, ex3 = st.columns(3)
with ex1:
    nifty_expiry = st.text_input("NIFTY expiry", placeholder="e.g. 28JUL26",
                                 help="The expiry of every NIFTY strike in the orderbook "
                                      "(one expiry per index per day). Shown as the "
                                      "chain's expiry label — strikes themselves are "
                                      "read from the symbols.")
with ex2:
    sensex_expiry = st.text_input("SENSEX expiry", placeholder="e.g. 23JUL26")
with ex3:
    banknifty_expiry = st.text_input("BANKNIFTY expiry", placeholder="e.g. 30JUL26")

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
if summary2_file is not None:
    _known_algos = _known_algos + [
        a for a in _algo_options_from_mtm(summary2_file.getvalue(), summary2_file.name)
        if a not in _known_algos]
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
        disabled=(orderbook_file is None or summary_file is None
                  or (mode == "0DTE" and maxloss_1dte is None)),
    )


# ---------------------------------------------------------------------------
# Processing — one pass computes both reports
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _process_all(ob_bytes, ob_name, mtm_bytes, mtm_name, mtm2_bytes, mtm2_name,
                 one_bytes, four_bytes, mlob_bytes, mlob_name, deviation_value, schema):
    # `schema` only serves as part of the cache key: when the report columns
    # change, stale cached rows from an older code version are not reused.
    multiplier = Decimal(str(deviation_value))

    # Segregation first — its Int / Positional classification also splits the
    # trade value users. Without max-loss files nothing is classified, so the
    # trade value rows stay in one group per algo (type_map = None). A
    # secondary User MTM (its indexes ran on their own servers) is processed
    # as its own comp so the summary section can render per MTM file; the
    # max-loss books are ROUTED between the two files first — handing the
    # full books to both comps would apply the same entry twice whenever a
    # user id appears in both.
    four = seg.load_combined(io.BytesIO(four_bytes)) if four_bytes else {}
    one = seg.load_combined(io.BytesIO(one_bytes)) if one_bytes else {}
    classified = bool(four or one)
    comp_df = seg.read_compiled(io.BytesIO(mtm_bytes), mtm_name)
    comp2_df = seg.read_compiled(io.BytesIO(mtm2_bytes), mtm2_name) if mtm2_bytes else None
    four2 = one2 = {}
    if comp2_df is not None:
        def _ids_servers(df):
            uids = set(df["UserID"].astype(str).str.strip())
            servers = (set(df["SERVER"].map(seg._server_key))
                       if "SERVER" in df.columns else set())
            return uids, servers
        uids1, servers1 = _ids_servers(comp_df)
        uids2, servers2_keys = _ids_servers(comp2_df)
        four, four2 = seg.split_max_loss_books(four, uids1, servers1, uids2, servers2_keys)
        one, one2 = seg.split_max_loss_books(one, uids1, servers1, uids2, servers2_keys)
    comp = seg.prepare_comp(comp_df, four, one, classified=classified)
    comp2 = (seg.prepare_comp(comp2_df, four2, one2, classified=classified)
             if comp2_df is not None else None)
    report_date = seg.infer_report_date(comp)

    type_map = {}
    for c in (comp, comp2):
        if c is None or not seg.is_classified(c):
            continue
        servers = (c["SERVER"].map(seg._server_key) if "SERVER" in c.columns
                   else [""] * len(c))
        type_map.update({
            (tv_user_key(uid), srv): ("Int" if t == "Intraday" else "Pos+Int")
            for uid, srv, t in zip(c["UserID"], servers, c["Type"])
        })
    type_map = type_map or None

    # Trade value — allocations merged across both MTMs (the orderbook is
    # combined); on a (user, server) present in both, the primary wins
    orders = read_orderbook(io.BytesIO(ob_bytes), ob_name)
    allocations = read_allocations(io.BytesIO(mtm_bytes), mtm_name)
    servers2 = set()
    if mtm2_bytes:
        for key, entries in read_allocations(io.BytesIO(mtm2_bytes), mtm2_name).items():
            existing = allocations.setdefault(key, [])
            seen = {e["server"] for e in existing}
            existing.extend(e for e in entries if e["server"] not in seen)
        if comp2 is not None and "SERVER" in comp2.columns:
            servers2 = set(comp2["SERVER"].map(seg._server_key))
    deduped = dedup_orders(orders)
    tv_rows = aggregate(deduped, allocations, multiplier, type_map)
    tv_groups = None
    if servers2:
        # each MTM's servers get their own outlier bands — a secondary index's
        # Lots per Cr sits on a different scale and would corrupt both medians
        tv_primary, tv_secondary = split_report_rows(tv_rows, servers2, multiplier)
        tv_groups = {"primary": tv_primary, "secondary": tv_secondary}
    strikes = strike_report(deduped, allocations)

    # Portfolio analysis (optional MLOB)
    pf_groups = None
    if mlob_bytes:
        mlob = read_mlob(io.BytesIO(mlob_bytes), mlob_name)
        pf_groups = portfolio_groups(mlob, allocations)

    return {
        "tv_rows": tv_rows,
        "tv_groups": tv_groups,
        "order_count": len(orders),
        "allocation_count": len(allocations),
        "strikes": strikes,
        "portfolio_groups": pf_groups,
        "pivot_rows": seg.pivot_rows(comp),
        "pivot_stats": seg.comp_stats(comp),
        "comp": comp,
        "comp2": comp2,
        "date": report_date,
    }


def _signature():
    return (
        orderbook_file.name if orderbook_file else None,
        summary_file.name if summary_file else None,
        summary2_file.name if summary2_file else None,
        maxloss_1dte.name if maxloss_1dte else None,
        maxloss_4dte.name if maxloss_4dte else None,
        mlob_file.name if mlob_file else None,
        deviation,
        mode,
    )


if process_clicked:
    try:
        with st.spinner("Processing…"):
            result = _process_all(
                orderbook_file.getvalue(), orderbook_file.name,
                summary_file.getvalue(), summary_file.name,
                summary2_file.getvalue() if summary2_file else None,
                summary2_file.name if summary2_file else None,
                maxloss_1dte.getvalue() if maxloss_1dte else None,
                maxloss_4dte.getvalue() if maxloss_4dte else None,
                mlob_file.getvalue() if mlob_file else None,
                mlob_file.name if mlob_file else None,
                deviation,
                tuple(REPORT_HEADER + STRIKE_ALGO_HEADER + STRIKE_CHAIN_HEADER
                      + PORTFOLIO_SUMMARY_HEADER
                      # bump when the engine changes shape or parsing rules,
                      # so cached results from an older build are not reused
                      + ["chain-breakdown-v1", "strike-only-v3"]),
            )
        _all_algos = pd.concat(
            [result["comp"]["ALGO"]]
            + ([result["comp2"]["ALGO"]] if result["comp2"] is not None else [])
        ).dropna().unique()
        st.session_state["dor"] = {
            **result, "deviation": deviation, "signature": _signature(),
            "slip_algo_options": [seg._algo_val(a) for a in sorted(_all_algos, key=float)],
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

# ---- MTM groups: a secondary User MTM splits the summary + outlier views ----
tv_groups = dor_state.get("tv_groups")
comp2 = dor_state.get("comp2")
mtm_label1 = mtm_label2 = None
if tv_groups:
    # label each MTM file by the indexes its servers actually traded
    mtm_label1 = ("/".join(sorted({r["segment"] for r in tv_groups["primary"]}))
                  or "Primary") + " (primary MTM)"
    mtm_label2 = ("/".join(sorted({r["segment"] for r in tv_groups["secondary"]}))
                  or "Secondary") + " (secondary MTM)"


def _pnl_color(value):
    """Profit green / loss red for the Realized, Unrealized and MTM columns."""
    if pd.isna(value) or value == 0:
        return ""
    return ("color:#15803D;font-weight:600" if value > 0
            else "color:#DC2626;font-weight:600")


def _styled_pivot(pivot_rows_view):
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

    return (
        pivot_df.style
        .apply(_pivot_row_style, axis=1)
        .set_properties(**{"text-align": "center"})
        .map(_pnl_color, subset=["Realized P&L"])
        .format({
            "MAX LOSS": format_crore, "ALLOCATION": format_crore,
            "Realized P&L": format_crore, "P&L %": "{:.2f}",
        })
    )


# ---- Segregation (first) — one block per MTM file ----
st.header("Segregation Pivot (Int / Pos+Int)")
algo_options = dor_state.get("slip_algo_options") or [
    seg._algo_val(a) for a in sorted(dor_state["comp"]["ALGO"].dropna().unique(), key=float)]
picked_seg_algos = seg_algos_input or algo_options
st.caption(
    (f"Report date: {report_date} · " if report_date else "")
    + "Algos included (from the inputs at the top): "
    + ", ".join(str(a) for a in picked_seg_algos)
)

seg_views = [(None, dor_state["comp"])]
if comp2 is not None:
    seg_views = [(mtm_label1, dor_state["comp"]), (mtm_label2, comp2)]
seg_dor_sections = []
for sec_label, comp_src in seg_views:
    comp_seg = comp_src[comp_src["ALGO"].map(seg._algo_val).isin(picked_seg_algos)]
    sec_stats = seg.comp_stats(comp_seg)
    sec_rows = seg.pivot_rows(comp_seg)
    seg_dor_sections.append({"label": sec_label, "stats": sec_stats, "rows": sec_rows})
    grand_view = next(r for r in sec_rows if r["kind"] == "grandtotal")
    if sec_label:
        st.subheader(sec_label)
    # without max-loss files there is no Positional / Intraday count to show
    tiles = [("Accounts", str(sec_stats["accounts"]))]
    if sec_stats["positional"] is not None:
        tiles += [("Positional", str(sec_stats["positional"])),
                  ("Intraday", str(sec_stats["intraday"]))]
    tiles += [("Allocation", format_crore(grand_view["Allocation"] * 100)),
              ("Realized P&L", format_crore(grand_view["Realized"])),
              ("P&L %", f"{grand_view['Return']:.2f}")]
    for col, (label, value) in zip(st.columns(len(tiles)), tiles):
        col.metric(label, value)
    st.dataframe(_styled_pivot(sec_rows), width="stretch", hide_index=True,
                 height=600 if len(seg_views) == 1 else 420)

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
# slippage covers every account — both MTM files together when two are given
comp_all = (pd.concat([dor_state["comp"], comp2], ignore_index=True)
            if comp2 is not None else dor_state["comp"])
comp_sel = comp_all[comp_all["ALGO"].map(seg._algo_val).isin(picked_slip_algos)]
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

st.caption(
    "All columns count users. Outliers are judged per user on **Lots per Cr** — combined "
    "lots ÷ normalise, where normalise = allocation / 1,00,000 (sub-1-Cr allocations "
    "normalise fractionally: 80,000 → 0.8, … 20,000 → 0.2) — so Below + In + Above = "
    "Total Users (only users with no usable allocation carry no flag). Pick an algo below "
    "to see its outlier user ids."
)
# one Algo Summary (and outlier bands) per MTM file — a secondary index's
# Lots per Cr scale must not share a median with the primary's
tv_view_groups = [(None, tv_rows)]
if tv_groups:
    tv_view_groups = [(mtm_label1, tv_groups["primary"]),
                      (mtm_label2, tv_groups["secondary"])]
tv_dor_sections = []
for gi, (group_label, group_rows) in enumerate(tv_view_groups):
    st.subheader("Algo Summary" + (f" — {group_label}" if group_label else ""))
    summary_rows = algo_summary(group_rows, used_multiplier)
    user_obs = user_lot_observations(group_rows, used_multiplier)
    tv_dor_sections.append({"label": group_label, "summary": summary_rows,
                            "obs": user_obs})
    df_summary = pd.DataFrame([format_summary_row(s) for s in summary_rows],
                              columns=SUMMARY_HEADER)
    st.dataframe(df_summary.style.set_properties(**{"text-align": "center"}),
                 width="stretch", hide_index=True)

    _summary_labels = [f"Algo {s['algo']}" + (f" · {s['user_type']}" if s["user_type"] else "")
                       for s in summary_rows]
    algo_pick = st.selectbox(
        "Outlier users — pick an algo / type",
        ["—"] + _summary_labels,
        key=f"summary_drill_{gi}",
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
    + " — a strike is one contract: index + strike + CE/PE; every symbol format "
    "(e.g. NIFTY21JUL26…, NIFTY26721… and the spaced forms) counts as one. The "
    "expiry per index comes from the expiry inputs at the top."
)
chain = strike_chain(strikes["per_strike"])
segments = sorted(chain)

# expiry per index from the inputs at the top — a pure display label (one
# orderbook holds a single expiry per index); "NA" when not entered
expiry_map = {}
for idx_name, entered in (("NIFTY", nifty_expiry), ("SENSEX", sensex_expiry),
                          ("BANKNIFTY", banknifty_expiry)):
    entered = entered.strip().upper()
    if entered:
        expiry_map[idx_name] = entered
label_of = {s: expiry_map.get(s, "NA") for s in segments}
expiry_labels = sorted(set(label_of.values()))
missing_expiry = sorted(s for s in segments if s not in expiry_map)
if missing_expiry:
    st.info("No expiry entered for " + ", ".join(missing_expiry)
            + " — enter it in the inputs at the top to label the chain "
              "(shown as NA until then).")

# day-mid per index from the open/close inputs -> the chain's ATM
mids = {}
if nifty_open > 0 and nifty_close > 0:
    mids["NIFTY"] = (nifty_open + nifty_close) / 2
if sensex_open > 0 and sensex_close > 0:
    mids["SENSEX"] = (sensex_open + sensex_close) / 2
if banknifty_open > 0 and banknifty_close > 0:
    mids["BANKNIFTY"] = (banknifty_open + banknifty_close) / 2

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
            if (as_expiry == "All" or label_of.get(c[0], "NA") == as_expiry)
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
    with c_idx:
        index = st.selectbox("Index",
                             sorted(s for s in segments if label_of[s] == expiry_label),
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
        if r["segment"] != index:
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
add_report_sheets(workbook, tv_rows, used_multiplier,
                  summary_groups=([("summary", tv_groups["primary"]),
                                   ("summary 2", tv_groups["secondary"])]
                                  if tv_groups else None))
add_strikes_sheet(workbook, strikes, expiry_map)
if pf_groups is not None:
    add_portfolio_sheet(workbook, portfolio_report(pf_groups, DEFAULT_PORTFOLIO_PATTERN))
seg.add_segregation_sheets(workbook, dor_state["comp"], report_date)
if comp2 is not None:
    seg.add_segregation_sheets(workbook, comp2, report_date, suffix=" 2")
excel_buffer = io.BytesIO()
workbook.save(excel_buffer)

# the report's slippage section carries the algo selection made in the
# inputs (same as the dashboard); the chain centres on the entered day-mids;
# each MTM file gets its own segregation + Algo Summary section
dor_html = build_dor_html(
    report_date=report_date or "—",
    deviation=float(used_deviation),
    tv_totals=totals,
    seg_sections=seg_dor_sections,
    tv_sections=tv_dor_sections,
    strikes=strikes,
    slippage={"summary": slip_algo, "overall": slip_overall, "majors": slip_majors,
              "empty_note": slip_empty_note},
    portfolio=pf_groups,
    mids=mids,
    expiry_map=expiry_map,
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
