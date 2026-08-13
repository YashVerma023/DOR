"""Daily Operations Report (DOR) — Streamlit UI.

Run with:  streamlit run app.py

All inputs sit in one block at the top: the compiled orderbook, the compiled
user MTM, the All User Details sheet and the outlier deviation; the DTE the
report covers is picked in the sidebar. One "Process" click computes BOTH
reports — the trade value analysis and the Int / Pos+Int segregation pivot —
and offers exactly two downloads:

  * DOR_<date>.xlsx — all data sheets in one workbook (tradevalue, summary,
    Strikes, Portfolio QS when the MLOB is given, Summary, MTM Data,
    Slippage)
  * DOR_<date>.html — the styled, client-shareable summary report
"""

import io
import logging
from datetime import date
from decimal import Decimal

import pandas as pd
import streamlit as st
from openpyxl import Workbook

import marketdata
import summary as seg
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
    ORDER_SERVER_HEADER,
    ORDER_SUMMARY_HEADER,
    REPORT_HEADER,
    STRIKE_ALGO_HEADER,
    STRIKE_CHAIN_HEADER,
    SUMMARY_HEADER,
    add_orders_sheet,
    add_report_sheets,
    add_strikes_sheet,
    add_user_aliases,
    aggregate,
    algo_summary,
    dedup_orders,
    dedup_orders_with_report,
    format_crore,
    lots_timeline,
    format_indian,
    format_order_row,
    format_row,
    fill_missing_servers,
    format_summary_row,
    indexes_in_orderbook,
    reference_bands_from_symbols,
    strike_bands,
    order_summary,
    order_summary_totals,
    read_allocations,
    read_orderbook,
    report_totals,
    multi_index_users,
    split_rows_by_segment,
    strike_chain,
    strike_report,
    user_lot_observations,
)
from tradevalue import _server_key as tv_server_key
from tradevalue import _user_key as tv_user_key

# Shape marker for the cached/processed result. BUMP THIS whenever the dict
# returned by _process_all gains, loses or renames a field: it both busts the
# st.cache_data entry and invalidates any session_state result an older build
# left behind, so a shape change asks for a re-Process instead of crashing.
STATE_VERSION = "orderbook-v2-format-v21"

# One-time logging setup. Streamlit re-runs this module on every interaction,
# so guard against stacking a handler per rerun.
if not logging.getLogger().handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(name)s | %(message)s",
    )
logger = logging.getLogger("dor.app")

st.set_page_config(page_title="Daily Operations Report", layout="wide")
st.title("Daily Operations Report")
st.caption(
    "Pick the DTE in the sidebar, upload the three inputs (plus the optional MLOB for the "
    "Portfolio Analysis and the optional secondary MTM), set the outlier deviation, and "
    "click **Process** — the Trade Value analysis, the Portfolio Analysis and the "
    "Int / Pos+Int Segregation are computed in one go, with a single Excel workbook and a "
    "shareable DOR.html as output."
)

# ---------------------------------------------------------------------------
# DTE — decides which All User accounts are in scope for the report
# ---------------------------------------------------------------------------
dte = st.sidebar.radio(
    "DTE", seg.DTE_OPTIONS,
    help="Which accounts the report covers, read from the All User sheet's "
         "Running Days column. Running Days states the days an account RUNS "
         "ON, so the scopes are cumulative — a DAILY account trades on every "
         "one of them.\n\n"
         "0DTE — every running account (0DTE + 1DTE/0DTE + DAILY).\n\n"
         "1DTE — 1DTE/0DTE + DAILY.\n\n"
         "4DTE — DAILY only.\n\n"
         "Dealer and stopped accounts (DLR ACC / NOT RUNNING) are dropped for "
         "every DTE. An in-scope account is typed by its Running Type: "
         "INT → Int, POS → Pos+Int.",
)
_DTE_CAPTION = {
    "0DTE": "**0DTE** — every running account: `0DTE` + `1DTE/0DTE` + `DAILY`.",
    "1DTE": "**1DTE** — `1DTE/0DTE` **+ `DAILY`**, since a DAILY account trades "
            "on a 1DTE day too.",
    "4DTE": "**4DTE** — `DAILY` only; those are the accounts running that far out.",
}
st.sidebar.caption(
    _DTE_CAPTION[dte]
    + "\n\nRows reading `DLR ACC` or `NOT RUNNING` in **server**, **Running Type** "
      "or **Running Days** are dropped on upload, whatever the DTE. A compiled "
      "account outside the scope is reported as **Unclassified** at the bottom "
      "of the pivot."
)

# ---------------------------------------------------------------------------
# Inputs — everything at the top
# ---------------------------------------------------------------------------
in1, in2 = st.columns(2)
with in1:
    orderbook_file = st.file_uploader("Orderbook (CSV / Excel)", type=["csv", "xlsx", "xlsm", "xls"])
with in2:
    summary_file = st.file_uploader("Compiled User MTM (CSV / Excel)", type=["csv", "xlsx", "xlsm", "xls"])

au1, au2 = st.columns([2, 1])
with au1:
    all_user_file = st.file_uploader(
        f"All User Details — tab \"{seg.ALL_USER_SHEET}\" (required)",
        type=["xlsx", "xlsm", "xls", "csv"],
        help="The account master. Columns read: userId, server, algo, max_loss, "
             "Running Type, Running Days. It supplies the Int / Pos+Int "
             "classification; MAX LOSS still comes from the User MTM, and the "
             "sheet's max_loss is carried through as a reference column only.",
    )
with au2:
    alias_file = st.file_uploader(
        "User alias map (optional override)",
        type=["json", "csv", "xlsx", "xlsm", "xls"],
        help="Accounts the two files name differently — the MTM's XLDH142 is "
             "the All User sheet's CC04. Leave empty to use user_aliases.json "
             "from the app folder; upload a JSON "
             "{\"CC04\": \"XLDH142\"} or a two-column table "
             "(All User id, MTM id) to replace it for this run only.",
    )

in5, in6 = st.columns(2)
with in5:
    mlob_file = st.file_uploader(
        "Multileg Orders — MLOB (optional, enables the Portfolio Analysis)",
        type=["xlsx", "xlsm", "xls", "csv"],
    )
# One ATM premium file per index, each optional — only the charted index's
# file is read, so uploading all three costs nothing.
st.caption("**ATM premium data** (optional) — adds the premium line to the "
           "intraday chart. Only the charted index's file is used.")
_prem_cols = st.columns(3)
premium_files = {}
for _col, _idx in zip(_prem_cols, marketdata.INDEX_SYMBOLS):
    with _col:
        premium_files[_idx] = st.file_uploader(
            f"{_idx} premium ({marketdata.INDEX_ABBR[_idx]})",
            type=["csv", "xlsx", "xlsm", "xls"], key=f"prem_file_{_idx}",
            help="Any CSV/Excel with a time column and a premium value — the "
                 "columns are auto-detected and correctable below.",
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

# ---------------------------------------------------------------------------
# Market data — day High / Low per index. There is NO date input: the date is
# read from the data itself (the compiled MTM's Date column, and the premium
# upload when one is given), so the report can never be built against a date
# nobody typed correctly.
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Fetching index levels…")
def _index_levels(day, indexes=None):
    """Day High / Low for the given indexes on `day` — only the ones the
    orderbook traded, so BANKNIFTY is not fetched on the days it is absent.
    Cached on the arguments, so changing any other input does not refetch."""
    return marketdata.fetch_index_levels(day, indexes=list(indexes) if indexes else None)


@st.cache_data(show_spinner="Fetching intraday series…")
def _index_intraday(day, index_name):
    """One-minute closes for the charted index. Fetched at 1m and bucketed to
    the chosen timeframe IN THE BROWSER, so every timeframe comes from a single
    embedded payload and switching one is instant."""
    levels, problems = marketdata.fetch_index_levels(
        day, indexes=[index_name], interval="1m")
    entry = levels.get(index_name) or {}
    return ([[bar["t"], round(bar["close"], 2)] for bar in entry.get("series", [])],
            problems.get(index_name))



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

@st.cache_data(show_spinner=False)
def _scope_preview(all_user_bytes, all_user_name, dte_value):
    """(running accounts after the drop, in scope for this DTE, POS, INT) —
    shown next to the uploader so the scope is visible before Processing."""
    all_users = seg.read_all_users(io.BytesIO(all_user_bytes), all_user_name)
    return (len(all_users), *seg.scope_stats(all_users, dte_value))


if all_user_file is not None:
    try:
        _running, _in_scope, _n_pos, _n_int = _scope_preview(
            all_user_file.getvalue(), all_user_file.name, dte)
        _alias_src = ("the uploaded map" if alias_file is not None
                      else "`user_aliases.json`")
        try:
            _n_alias = len(seg.load_user_aliases(
                io.BytesIO(alias_file.getvalue()), alias_file.name)
                if alias_file is not None else None)
            _alias_note = f" · **{_n_alias}** user alias(es) from {_alias_src}"
        except Exception as exc:
            _alias_note = f" · ⚠️ alias map unreadable ({exc})"
        st.caption(
            f"All User sheet: **{_running}** running account(s) after the "
            f"`DLR ACC` / `NOT RUNNING` drop · **{_in_scope}** in scope for "
            f"**{dte}** (POS {_n_pos} · INT {_n_int})" + _alias_note + "."
        )
    except Exception as exc:
        st.error(f"Could not read the All User sheet: {exc}")

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
                  or all_user_file is None),
    )


# ---------------------------------------------------------------------------
# Processing — one pass computes both reports
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _process_all(ob_bytes, ob_name, mtm_bytes, mtm_name, mtm2_bytes, mtm2_name,
                 all_user_bytes, all_user_name, dte_value, alias_bytes, alias_name,
                 mlob_bytes, mlob_name, deviation_value, schema):
    # `schema` only serves as part of the cache key: when the report columns
    # change, stale cached rows from an older code version are not reused.
    multiplier = Decimal(str(deviation_value))

    # Segregation first — its Int / Positional classification also splits the
    # trade value users. The classification comes from the All User sheet,
    # narrowed to the accounts the selected DTE covers; both MTM files read
    # the SAME scope map, since it is keyed on the user id alone and each
    # compiled row simply looks itself up (nothing is consumed or shared, so
    # no routing between the two files is needed).
    all_users = seg.read_all_users(io.BytesIO(all_user_bytes), all_user_name)
    scope = seg.dte_scope(all_users, dte_value)
    # accounts the two files name differently (MTM XLDH142 = sheet CC04) —
    # the uploaded map wins over the on-disk user_aliases.json for this run
    aliases = seg.load_user_aliases(
        io.BytesIO(alias_bytes) if alias_bytes else None, alias_name)

    comp_df = seg.read_compiled(io.BytesIO(mtm_bytes), mtm_name)
    comp2_df = seg.read_compiled(io.BytesIO(mtm2_bytes), mtm2_name) if mtm2_bytes else None
    comp = seg.prepare_comp(comp_df, scope, aliases=aliases)
    comp2 = (seg.prepare_comp(comp2_df, scope, aliases=aliases)
             if comp2_df is not None else None)
    report_date = seg.infer_report_date(comp)

    # Unclassified accounts carry no Int / Pos+Int type into the trade value
    # and orders tables either — a blank type, which _resolve_type already
    # handles, rather than a fabricated Intraday row.
    type_map = {}
    for c in (comp, comp2):
        if c is None or not seg.is_classified(c):
            continue
        servers = (c["SERVER"].map(seg._server_key) if "SERVER" in c.columns
                   else [""] * len(c))
        type_map.update({
            (tv_user_key(uid), srv): ("Int" if t == "Intraday" else "Pos+Int")
            for uid, srv, t in zip(c["UserID"], servers, c["Type"])
            if t != seg.UNCLASSIFIED
        })
    type_map = type_map or None

    # Trade value — allocations merged across both MTMs (the orderbook is
    # combined); on a (user, server) present in both, the primary wins.
    # The orderbook is read ONCE with every status: the trade value and
    # strikes use the COMPLETE subset, the orders summary needs them all.
    all_orders = read_orderbook(io.BytesIO(ob_bytes), ob_name, all_statuses=True)
    orders = [o for o in all_orders if o.status == "COMPLETE"]
    allocations = read_allocations(io.BytesIO(mtm_bytes), mtm_name)
    if mtm2_bytes:
        for key, entries in read_allocations(io.BytesIO(mtm2_bytes), mtm2_name).items():
            existing = allocations.setdefault(key, [])
            seen = {e["server"] for e in existing}
            existing.extend(e for e in entries if e["server"] not in seen)
    # a newer export carries no SERVER column — recover it per user from the
    # User MTM before anything matches on it
    all_orders, server_fill = fill_missing_servers(all_orders, allocations)
    orders = [o for o in all_orders if o.status == "COMPLETE"]

    # the orderbook writes the BASE user id where the MTM writes the account
    # (JSR129 vs JSR129A31) — resolve those by prefix within the server before
    # anything reads an algo, or the user drops out of every algo table
    aliases = add_user_aliases(allocations, all_orders)

    # every uploaded orderbook is duplicate-checked on
    # user + date + order id + symbol before anything is aggregated
    # Strike validation band, derived from the day's own High / Low rather
    # than a hardcoded rupee range. Only the indexes the orderbook actually
    # traded are fetched — BANKNIFTY is absent most days. If the fetch fails
    # the band comes from the orderbook's own unambiguous spaced symbols.
    ob_indexes = indexes_in_orderbook(all_orders)
    band_levels, _band_problems = ({}, {})
    if report_date and ob_indexes:
        band_levels, _band_problems = marketdata.fetch_index_levels(
            report_date, indexes=ob_indexes)
    bands = strike_bands(band_levels)
    if len(bands) < len(ob_indexes):
        fallback = reference_bands_from_symbols(o.symbol for o in all_orders)
        for name in ob_indexes:
            if name not in bands and name in fallback:
                bands[name] = fallback[name]

    deduped, ob_dup = dedup_orders_with_report(orders, "orderbook")
    tv_rows = aggregate(deduped, allocations, multiplier, type_map)
    # Trade Value and Orders are reported PER INDEX — lot sizes (and so lots
    # per Cr) differ per index, and one pooled median would belong to none of
    # them. Each index's outlier bands are computed within that index.
    tv_by_segment = split_rows_by_segment(tv_rows, multiplier)
    multi_count, multi_map = multi_index_users(tv_rows)
    strikes = strike_report(deduped, allocations, bands)
    # the orders summary splits per MTM exactly like the trade value tables —
    # one table over ALL servers would compare a per-MTM Algo Summary against
    # an everything Orders Summary and the user counts could never agree
    # lots fired per minute / algo / colour category — the intraday chart's
    # dots. Built from ALL statuses so failed orders are visible, and from the
    # raw list (not the deduped one) is deliberately avoided: duplicates would
    # inflate the lots exactly as they would inflate the trade value.
    lots_chart = lots_timeline(dedup_orders(all_orders), allocations)

    order_rows = order_summary(all_orders, allocations, type_map, bands)
    orders_by_segment = {}
    for segment in sorted({o.segment for o in all_orders}):
        orders_by_segment[segment] = order_summary(
            [o for o in all_orders if o.segment == segment], allocations,
            type_map, bands)

    # Portfolio analysis (optional MLOB)
    pf_groups = None
    mlob_dup = None
    if mlob_bytes:
        mlob, mlob_dup = read_mlob(io.BytesIO(mlob_bytes), mlob_name)
        pf_groups = portfolio_groups(mlob, allocations)

    return {
        "tv_rows": tv_rows,
        "tv_by_segment": tv_by_segment,
        "multi_index": (multi_count, multi_map),
        "aliases": aliases,
        "order_count": len(orders),
        "allocation_count": len(allocations),
        "strikes": strikes,
        "order_rows": order_rows,
        "lots_chart": lots_chart,
        "ob_indexes": ob_indexes,
        "orders_by_segment": orders_by_segment,
        "portfolio_groups": pf_groups,
        "pivot_rows": seg.pivot_rows(comp),
        "pivot_stats": seg.comp_stats(comp),
        # data-quality probe: the compiled sheet ships an MTM column that is
        # only filled where Unrealized != 0, so it is recomputed — surfaced so
        # the correction is visible rather than silent
        "mtm_check": seg.mtm_column_mismatch(comp_df),
        # duplicate checks on the two files that can carry repeats; the User
        # MTM and All User sheets are pre-checked upstream and not tested here
        "dup_checks": [r for r in (ob_dup, mlob_dup) if r],
        "server_fill": server_fill,
        # diagnostic lookup for the "unclassified" sheet — covers every All
        # User row, dropped ones included, so it can say WHY an account fell out
        "all_user_ref": seg.all_user_reference(all_users),
        "comp": comp,
        "comp2": comp2,
        "date": report_date,
        "dte": dte_value,
        "scope_size": len(scope),
        "alias_count": len(aliases),
        # which compiled accounts actually matched under a different id
        "alias_hits": sorted({
            (str(u), str(a)) for c in (comp, comp2) if c is not None
            for u, a in zip(c["UserID"], c["All User ID"]) if str(a)
        }),
    }


def _signature():
    return (
        orderbook_file.name if orderbook_file else None,
        summary_file.name if summary_file else None,
        summary2_file.name if summary2_file else None,
        all_user_file.name if all_user_file else None,
        alias_file.name if alias_file else None,
        mlob_file.name if mlob_file else None,
        # one premium upload per index, so the signature carries them all;
        # `premium_file` (the charted index's one) does not exist yet here
        tuple(sorted((idx, f.name) for idx, f in premium_files.items() if f)),
        deviation,
        dte,
    )


if process_clicked:
    try:
        with st.spinner("Processing…"):
            result = _process_all(
                orderbook_file.getvalue(), orderbook_file.name,
                summary_file.getvalue(), summary_file.name,
                summary2_file.getvalue() if summary2_file else None,
                summary2_file.name if summary2_file else None,
                all_user_file.getvalue(), all_user_file.name,
                dte,
                alias_file.getvalue() if alias_file else None,
                alias_file.name if alias_file else None,
                mlob_file.getvalue() if mlob_file else None,
                mlob_file.name if mlob_file else None,
                deviation,
                tuple(REPORT_HEADER + STRIKE_ALGO_HEADER + STRIKE_CHAIN_HEADER
                      + PORTFOLIO_SUMMARY_HEADER
                      # bump when the engine changes shape or parsing rules,
                      # so cached results from an older build are not reused
                      + ORDER_SUMMARY_HEADER
                      + ["chain-breakdown-v1", "strike-only-v3",
                         "user-alias-v1", STATE_VERSION]),
            )
        _all_algos = pd.concat(
            [result["comp"]["ALGO"]]
            + ([result["comp2"]["ALGO"]] if result["comp2"] is not None else [])
        ).dropna().unique()
        st.session_state["dor"] = {
            **result, "deviation": deviation, "signature": _signature(),
            "state_version": STATE_VERSION,
            "slip_algo_options": [seg._algo_val(a) for a in sorted(_all_algos, key=float)],
        }
    except Exception as exc:
        st.error(f"Processing failed: {exc}")


# MTM = Realized P&L + Unrealized P&L, computed per account from the uploaded
# User MTM (the sheet's own MTM column is not used — it is only populated where
# Unrealized is non-zero). All three are shown so the sum is verifiable.
PIVOT_COLS = ["Section", "ALGO", "SERVER", "No. of Users", "MAX LOSS", "ALLOCATION",
              "Realized P&L", "Unrealized P&L", "MTM", "MTM %",
              "No. of SL Hit Users"]

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
dor_state = st.session_state.get("dor")
if dor_state is None:
    st.info("Upload the orderbook, the Compiled User MTM and the All User Details "
            "sheet, pick the DTE in the sidebar, then click **Process**.")
    st.stop()

# A result computed by an older build can be missing fields this page now
# reads — the dashboard re-renders from session_state on every interaction,
# so without this guard a shape change surfaces as a raw KeyError traceback.
if dor_state.get("state_version") != STATE_VERSION:
    st.session_state.pop("dor", None)
    st.info("The app was updated since these results were computed — "
            "click **Process** to recompute.")
    st.stop()

if dor_state["signature"] != _signature():
    st.warning("Inputs changed since the last run — click **Process** to refresh the results.")

# Downloads sit here, right under the inputs — but the workbook and the HTML
# can only be built once every section below has been computed. A container
# reserves the slot now and is filled at the very end of the script, so the
# buttons render at the top without moving the work that produces them.
downloads_slot = st.container()

tv_rows = dor_state["tv_rows"]
used_deviation = dor_state["deviation"]
used_multiplier = Decimal(str(used_deviation))
report_date = dor_state["date"]

# ---- Index day High / Low ----
# The date is PRE-FILLED from the data (the compiled MTM's Date column) but
# stays editable: the index series is fetched per date, so a missing or wrong
# Date column must not strand the report, and a past day can be re-charted
# without touching the uploads.
st.subheader("Market data")
_dc1, _dc2 = st.columns([1, 3])
with _dc1:
    try:
        _default_date = marketdata._as_date(report_date) if report_date else date.today()
    except ValueError:
        _default_date = date.today()
    market_date = st.date_input(
        "Market date", value=_default_date, format="DD-MM-YYYY",
        help="Which trading day's index data to fetch. Pre-filled from the "
             "User MTM's Date column — change it only to override.",
    )
_market_date_text = f"{market_date:%d-%m-%Y}"
if report_date and _market_date_text != report_date:
    st.warning(
        f"⚠️ Market date **{_market_date_text}** differs from the report date "
        f"**{report_date}** read from the User MTM. The index levels and the "
        "chart will describe a different day than the book below."
    )

_levels, _level_problems = _index_levels(
    market_date, tuple(dor_state.get("ob_indexes") or ()))
if _levels:
    with _dc2:
        lvl_cols = st.columns(len(_levels))
        for col, (idx_name, lv) in zip(lvl_cols, _levels.items()):
            col.metric(
                f"{idx_name} mid", f"{lv['mid']:,.2f}",
                help=f"High {lv['high']:,.2f} · Low {lv['low']:,.2f} · "
                     f"Open {lv['open']:,.2f} · Close {lv['close']:,.2f} ({lv['symbol']})",
            )
    st.caption(f"Index day High / Low for **{_market_date_text}** · "
               "ATM anchor = (High + Low) / 2.")
if _level_problems:
    st.warning(
        "Could not fetch " + ", ".join(sorted(_level_problems))
        + f" for {_market_date_text} — " + list(_level_problems.values())[0]
        + ". Enter the day mid manually below to still centre the chain."
    )

# manual override — only for the indexes the fetch could not resolve
_manual_mids = {}
if _level_problems:
    man_cols = st.columns(len(_level_problems))
    for col, idx_name in zip(man_cols, sorted(_level_problems)):
        with col:
            _manual_mids[idx_name] = st.number_input(
                f"{idx_name} day mid (manual)", min_value=0.0, value=0.0,
                step=50.0, key=f"manual_mid_{idx_name}",
                help="0 = not set; the chain then simply is not centred for "
                     "this index.",
            )

if not tv_rows:
    st.warning("No completed NIFTY / BANKNIFTY / SENSEX orders found in the orderbook.")
    st.stop()

# ---- Per-index groups (Trade Value + Orders) and per-MTM groups (Segregation).
# The segregation pivot has no index dimension — it reports account P&L, which
# a User MTM states per account, not per index — so it stays split per MTM file.
tv_by_segment = dor_state.get("tv_by_segment") or {}
orders_by_segment = dor_state.get("orders_by_segment") or {}
comp2 = dor_state.get("comp2")


def _mtm_label(comp_src, suffix, fallback):
    """Name a User MTM by the indexes its own servers actually traded."""
    servers = (set(comp_src["SERVER"].map(seg._server_key))
               if "SERVER" in comp_src.columns else set())
    idx = sorted({r["segment"] for r in tv_rows
                  if tv_server_key(r["server"]) in servers})
    return ("/".join(idx) or fallback) + suffix


mtm_label1 = mtm_label2 = None
if comp2 is not None:
    mtm_label1 = _mtm_label(dor_state["comp"], " (primary MTM)", "Primary")
    mtm_label2 = _mtm_label(comp2, " (secondary MTM)", "Secondary")


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
                "Unrealized P&L": r.get("Unrealized", 0),
                "MTM": r["MTM"],
                "MTM %": r["Return"],
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
        .map(_pnl_color, subset=["Realized P&L", "Unrealized P&L", "MTM"])
        .format({
            "MAX LOSS": format_crore, "ALLOCATION": format_crore,
            "Realized P&L": format_crore, "Unrealized P&L": format_crore,
            "MTM": format_crore, "MTM %": "{:.2f}",
        })
    )


# ---- Segregation (first) — one block per MTM file ----
st.header("Segregation Pivot (Int / Pos+Int)")
algo_options = dor_state.get("slip_algo_options") or [
    seg._algo_val(a) for a in sorted(dor_state["comp"]["ALGO"].dropna().unique(), key=float)]
picked_seg_algos = seg_algos_input or algo_options
_used_dte = dor_state.get("dte", dte)
st.caption(
    (f"Report date: {report_date} · " if report_date else "")
    + f"DTE: **{_used_dte}** ({dor_state.get('scope_size', 0)} account(s) in scope) · "
    + "Algos included (from the inputs at the top): "
    + ", ".join(str(a) for a in picked_seg_algos)
)
st.caption(
    "**MTM = Realized P&L + Unrealized P&L**, computed per account from the uploaded "
    "User MTM; **MTM %** = MTM / Allocation. Every summary figure is aggregated from the "
    "**MTM Data** sheet in the Excel download."
)
# isinstance, not a bare .get: a result left by a build whose mtm_check had a
# different shape must degrade to "no note" rather than a raw traceback. The
# STATE_VERSION guard above is the real defence — this is the seatbelt for the
# next time that bump is forgotten.
# ---- Duplicate checks on the uploaded files ----
for _dup in dor_state.get("dup_checks") or []:
    if _dup.get("skipped"):
        st.caption(
            f"⚠️ **{_dup['label']}** — duplicate check skipped, the file is "
            f"missing the key column(s): {', '.join(_dup['skipped'])}."
        )
    elif _dup["colliding"]:
        _samples = ", ".join(f"`{' | '.join(str(p) for p in k)}` ×{n}"
                             for k, n in _dup["samples"][:3])
        st.caption(
            f"🔁 **{_dup['label']}** — {format_indian(_dup['surplus'])} duplicate "
            f"row(s) across {format_indian(_dup['colliding'])} key(s), removed "
            f"before any aggregation. {format_indian(_dup['rows'])} rows in → "
            f"{format_indian(_dup['distinct'])} unique. Sample: {_samples}"
            + (f" · ⚠️ {_dup['unkeyable']} row(s) had no usable order id and "
               "bypassed the check" if _dup.get("unkeyable") else "")
        )
    else:
        st.caption(
            f"✅ **{_dup['label']}** — no duplicates "
            f"({format_indian(_dup['rows'])} rows, all keys distinct)."
        )

_mtm = dor_state.get("mtm_check")
if isinstance(_mtm, dict) and _mtm.get("mismatch"):
    # report what was measured — the gap in rupees, not two crore-rounded
    # totals that can render identically when they differ by < 50,000
    _kinds = []
    if _mtm["blank"]:
        _kinds.append(f"{_mtm['blank']} left blank (0)")
    if _mtm["differs"]:
        _kinds.append(f"{_mtm['differs']} carrying a different value")
    st.caption(
        f"ℹ️ The uploaded sheet's own **MTM** column disagrees with "
        f"Realized + Unrealized on **{_mtm['mismatch']} of {_mtm['rows']}** rows"
        + (" — " + ", ".join(_kinds) if _kinds else "")
        + f". Book total: **₹{format_indian(_mtm['gap'])}** "
        + ("higher" if _mtm["gap"] >= 0 else "lower")
        + " than the column reports"
        + (f" (affected: {', '.join(f'`{u}`' for u in _mtm['users'][:5])}"
           + (" …" if len(_mtm["users"]) > 5 else "") + ")" if _mtm["users"] else "")
        + ". That column is **ignored** — MTM is recomputed from "
          "Realized + Unrealized on every row, so the figures above are unaffected."
    )
_alias_hits = dor_state.get("alias_hits") or []
if _alias_hits:
    st.caption(
        f"ℹ️ {len(_alias_hits)} account(s) matched the All User sheet under a "
        "different id (the User MTM and the sheet name them differently): "
        + ", ".join(f"`{mtm}` → `{au}`" for mtm, au in _alias_hits[:8])
        + (" …" if len(_alias_hits) > 8 else "")
        + " — from the alias map, so they are classified instead of falling "
          "into Unclassified."
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
    # nothing matched the All User sheet -> no Positional / Intraday counts
    tiles = [("Accounts", str(sec_stats["accounts"]))]
    if sec_stats["positional"] is not None:
        tiles += [("Positional", str(sec_stats["positional"])),
                  ("Intraday", str(sec_stats["intraday"]))]
    if sec_stats.get("unclassified"):
        tiles += [("Unclassified", str(sec_stats["unclassified"]))]
    tiles += [("Allocation", format_crore(grand_view["Allocation"] * 100)),
              ("Realized P&L", format_crore(grand_view["Realized"])),
              ("Unrealized P&L", format_crore(grand_view.get("Unrealized", 0))),
              ("MTM", format_crore(grand_view["MTM"])),
              ("MTM %", f"{grand_view['Return']:.2f}")]
    for col, (label, value) in zip(st.columns(len(tiles)), tiles):
        col.metric(label, value)
    st.dataframe(_styled_pivot(sec_rows), width="stretch", hide_index=True,
                 height=600 if len(seg_views) == 1 else 420)
    if sec_stats.get("unclassified"):
        st.caption(
            f"⚠️ {sec_stats['unclassified']} account(s) in this MTM are outside the "
            f"**{_used_dte}** scope — absent from the All User sheet, dropped as "
            "`DLR ACC` / `NOT RUNNING`, or running on other days. They are grouped in "
            "the **Unclassified** section at the bottom of the pivot and carry no "
            "Int / Pos+Int type in the Trade Value and Orders tables. Algo totals "
            "cover classified accounts only; the Grand Total covers everything."
        )

# ---- Slippage (realized loss % beyond max-loss %) ----
st.header("Max SL Slippage Analysis")
st.caption(
    "ML % = MAX LOSS / ALLOCATION; Realized ML % = |Realized P&L| / ALLOCATION — measured "
    "on the realized loss, not MTM, because a max-loss stop is about what was actually "
    "booked. Plain ratios of allocation, same convention as MTM %. "
    "An account has slippage only when Realized "
    "ML % exceeds ML % by at least 0.1 (1.00 → 1.09 is not slippage; 1.10 is). Avg Slippage = "
    "average Realized ML % of the algo's slippage accounts; **Major** = slippage accounts "
    "above their algo's average, plus the lone slippage account of an algo — with one "
    "account the average *is* that account, so it is that algo's worst by definition."
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
# accounts with an allocation but no configured max loss are not judged at all
slip_no_sl = seg.no_sl_accounts(comp_sel)
if slip_no_sl:
    st.caption(
        f"ℹ️ {len(slip_no_sl)} account(s) carry an allocation but **MAX LOSS = 0**, so no "
        "stop-loss is configured — they are **excluded** from this analysis (ML % would be "
        "0, making any loss past the threshold read as slippage against a limit that never "
        "existed) and listed in the Excel **no_sl_Acc** sheet: "
        + ", ".join(f"`{n['UserID']}`" for n in slip_no_sl[:8])
        + (" …" if len(slip_no_sl) > 8 else "")
    )
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
# one Algo Summary (and its own outlier bands) PER INDEX — lot sizes differ
# per index, so a pooled median would belong to none of them
tv_view_groups = ([(idx, rows) for idx, rows in tv_by_segment.items()]
                  or [(None, tv_rows)])
_multi_count, _multi_map = dor_state.get("multi_index") or (0, {})
if _multi_count and len(tv_view_groups) > 1:
    st.caption(
        f"⚠️ {_multi_count} user(s) traded more than one index, so they appear in "
        "more than one table below and each index judges them **independently**. "
        "In every table their **Lots per Cr** is that index's lots over their "
        "**whole** allocation — a partial-exposure figure — so a user who splits "
        "capital across indexes reads lower than a single-index peer deploying "
        "the same capital, and can show as *Below average range* without actually "
        "under-trading. Such users are marked **⚠ also trades …** in the outlier "
        "lists."
    )
tv_dor_sections = []
for gi, (group_label, group_rows) in enumerate(tv_view_groups):
    st.subheader("Algo Summary" + (f" — {group_label}" if group_label else ""))
    summary_rows = algo_summary(group_rows, used_multiplier)
    user_obs = user_lot_observations(group_rows, used_multiplier)
    # mark the users whose ratio here is a partial exposure, and of what
    for _o in user_obs:
        _others = [i for i in _multi_map.get(_o["user_id"], []) if i != group_label]
        _o["also_trades"] = _others
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
                [[o["user_id"], o["server"], round(float(o["lots_per_cr"])),
                  format_indian(o["lots"]), o["outlier"],
                  ("⚠ also trades " + ", ".join(o["also_trades"]))
                  if o.get("also_trades") else ""] for o in flagged],
                columns=["User ID", "Server", "Lots per Cr", "Lots", "Outlier",
                         "Partial exposure"],
            )
            st.dataframe(df_flagged.style.set_properties(**{"text-align": "center"}),
                         width="stretch", hide_index=True)
        else:
            st.info("No outlier users in this algo.")

# ---- Orders Summary (below Trade Value) ----
st.header("Orders Summary")
order_rows = dor_state.get("order_rows") or []
st.caption(
    "Every **option** order fired, whatever its outcome, per algo and Int / Pos+Int "
    "type — futures rows are excluded (same scope as the strikes section), so a rejected "
    "FUT order can't invent an algo row in an index that algo never traded. "
    "**Executed** = COMPLETE orders; "
    "**Failed/Cancelled/Rejected** = cancelled + rejected; **Pending** = still live at "
    "end of day (OPEN / OPEN_PENDING) — live is not failed, so they are kept apart. "
    "**Hedge** and **VAR** count the **executed** "
    "orders tagged `h_…` / `v_…` — they are a slice of **Executed**, never of Total "
    "Orders, so a cancelled hedge is counted in Failed like any other cancellation. "
    "Algo and "
    "type are resolved exactly as in the Trade Value rows, so the two tables agree; the "
    "one legitimate difference is a user whose *every* order in an index failed — shown "
    "here, absent from that index's Trade Value table. Users with no MTM entry have "
    "neither algo nor type and collect in the **—** row. Pick an algo below for its "
    "per-server split."
)
# split per index, exactly like the Trade Value Algo Summary above — so the
# two tables cover the same accounts and their user counts reconcile
order_view_groups = ([(idx, rows) for idx, rows in orders_by_segment.items()]
                     or [(None, order_rows)])
order_dor_sections = []
if not order_rows:
    st.info("No orders found in the orderbook.")
else:
    for oi, (order_label, group_order_rows) in enumerate(order_view_groups):
        order_dor_sections.append({"label": order_label, "rows": group_order_rows})
        if order_label:
            st.subheader(order_label)
        if not group_order_rows:
            st.info("No orders on these servers.")
            continue
        ord_totals = order_summary_totals(group_order_rows)
        df_orders = pd.DataFrame(
            [format_order_row(r) for r in group_order_rows]
            + [["Total", "", ord_totals["users"], ord_totals["orders"], ord_totals["executed"],
                ord_totals["failed"], ord_totals["pending"], ord_totals["hedge"],
            ord_totals["var"]]],
            columns=ORDER_SUMMARY_HEADER,
        )
        n_order_rows = len(group_order_rows)

        def _order_total_style(row, _n=n_order_rows):
            style = ("background-color:#D1C9E1; color:#1F2937; font-weight:700"
                     if row.name == _n else "")
            return [style] * len(row)

        st.dataframe(
            df_orders.style
            .apply(_order_total_style, axis=1)
            .set_properties(**{"text-align": "center"})
            .format({c: format_indian for c in ORDER_SUMMARY_HEADER[2:]}),
            width="stretch", hide_index=True,
        )
        _aliases = dor_state.get("aliases") or {}
        if _aliases and oi == 0:
            st.caption(
                f"ℹ️ {len(_aliases)} orderbook account(s) carried the base user id where "
                "the User MTM carries the full account id (e.g. "
                + ", ".join(f"`{k[0]}` on `{k[1]}` → `{v}`"
                            for k, v in list(_aliases.items())[:4])
                + (" …" if len(_aliases) > 4 else "")
                + "); they were matched by prefix within the same server, so their "
                  "orders and lots now sit under the right algo instead of the — row."
            )
        unattributed = next((r for r in group_order_rows if not r["algo"]), None)
        if unattributed:
            ids = unattributed.get("user_ids") or []
            st.caption(
                f"**—** = {unattributed['users']} user(s) that fired orders but appear "
                "in no User MTM, so they have no algo and no Int / Pos+Int type "
                "(the Trade Value Algo Summary leaves them out entirely)"
                + (": " + ", ".join(ids) if ids else "")
            )
        _order_labels = [f"Algo {r['algo'] or '—'}"
                         + (f" · {r['user_type']}" if r["user_type"] else "")
                         for r in group_order_rows]
        order_pick = st.selectbox("Server view — pick an algo / type",
                                  ["—"] + _order_labels, key=f"orders_drill_{oi}")
        if order_pick != "—":
            picked_order = group_order_rows[_order_labels.index(order_pick)]
            df_order_servers = pd.DataFrame(
                [[s["server"], s["users"], s["orders"], s["executed"], s["failed"],
                  s["pending"], s["hedge"], s["var"]]
                 for s in picked_order["servers"]],
                columns=ORDER_SERVER_HEADER,
            )
            st.dataframe(
                df_order_servers.style.set_properties(**{"text-align": "center"})
                .format({c: format_indian for c in ORDER_SERVER_HEADER[1:]}),
                width="stretch", hide_index=True,
            )


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

# day-mid per index -> the chain's ATM. Fetched (High + Low) / 2 first, with
# any manual entry filling in the indexes the fetch could not resolve.
mids = marketdata.mids_from(_levels, _manual_mids)

# ---- Intraday charts — ONE PER INDEX that traded ----
# Normally NIFTY and SENSEX; a third appears on days BANKNIFTY is in the book.
st.header("Intraday Charts")
_chart_choices = [s for s in marketdata.INDEX_SYMBOLS if s in segments]
chart_payloads = []
if not _chart_choices:
    st.info("No index traded in this orderbook, so there is nothing to chart.")
for chart_index in _chart_choices:
    st.markdown(f"**{chart_index}**")
    chart_payload = None
    _points, _chart_problem = _index_intraday(market_date, chart_index)
    if _chart_problem:
        st.warning(f"No intraday series for {chart_index} on "
                   f"{_market_date_text} — {_chart_problem}. The chart is "
                   "omitted from the report.")
        chart_payload = None
    elif not _points:
        st.warning(
            f"No 1-minute bars for {chart_index} on {_market_date_text}. "
            "Yahoo keeps 1-minute history for roughly 30 days, so an older "
            "report date returns nothing at this granularity."
        )
        chart_payload = None
    else:
        # series is a LIST so the premium line (and the algo-driven series
        # after it) are appended here without touching the renderer
        chart_payload = {
            "index": chart_index,
            "date": _market_date_text,
            # tooltip reads "Index (SX)" / "Premium" — the short forms are far
            # easier to scan at a hover point than the full index names
            "series": [{"name": f"Index ({marketdata.INDEX_ABBR.get(chart_index, chart_index)})",
                        "axis": "left", "points": _points}],
            "shade": ["15:15", "15:40"],
            "shade_label": "auction / extended window",
            # axis ticks snap to these: index to the nearest 50, premium to 1
            "axis_step": {"left": 50, "right": 1},
            # dots: lots fired per minute, split by algo and colour category
            "lots": (dor_state.get("lots_chart") or {}).get(chart_index),
        }
        st.caption(
            f"📈 {chart_index} — {format_indian(len(_points))} one-minute bars "
            f"({_points[0][0]} → {_points[-1][0]}), bucketed to the timeframe "
            "chosen in the report."
        )

    # ---- Premium line — the charted index's own upload ----
    premium_file = premium_files.get(chart_index)
    if premium_file is None and chart_payload is not None:
        st.caption(f"No premium file uploaded for **{chart_index}** — the chart "
                   "shows the index and the lots dots only.")
    if premium_file is not None and chart_payload is not None:
        try:
            _pdf, _guess = marketdata.read_premium(
                io.BytesIO(premium_file.getvalue()), premium_file.name)
        except Exception as exc:
            st.error(f"Could not read the premium file: {exc}")
            _pdf = None
        if _pdf is not None and not _pdf.empty:
            st.markdown("**Premium data — column mapping**")
            _cols = list(_pdf.columns)
            _none = "— none —"

            def _pick(label, key, guessed, optional=False):
                opts = ([_none] + _cols) if optional else _cols
                default = guessed if guessed in opts else opts[0]
                return st.selectbox(label, opts, index=opts.index(default),
                                    key=f"prem_{key}_{chart_index}")

            pc1, pc2, pc3, pc4 = st.columns(4)
            with pc1:
                _t_col = _pick("Time column", "time", _guess["time"])
            with pc2:
                _v_col = _pick("Premium column", "value", _guess["value"])
            with pc3:
                _d_col = _pick("Date column", "date", _guess["date"], optional=True)
            with pc4:
                _i_col = _pick("Index column", "index", _guess["index"], optional=True)

            # the file was uploaded into a named slot, so the index is already
            # known — detection is used only to catch a file in the wrong slot
            _prem_index = chart_index
            _detected, _how = marketdata.detect_index(
                _pdf, _guess, premium_file.name)
            if _detected and _detected != chart_index:
                st.warning(
                    f"⚠️ This file was uploaded under **{chart_index}** but "
                    f"looks like **{_detected}** ({_how}). Check you have not "
                    "swapped the premium files."
                )
            elif _detected:
                st.caption(f"🔎 Confirmed as **{_detected}** from {_how}.")

            _prem_points, _prem_meta = marketdata.premium_series(
                _pdf, _t_col, _v_col,
                date_col=None if _d_col == _none else _d_col,
                index_col=None if _i_col == _none else _i_col,
                index_name=None if _i_col == _none else _prem_index,
            )

            if _prem_meta["dates"] and _market_date_text not in _prem_meta["dates"]:
                st.warning(
                    f"⚠️ The premium file's date(s) "
                    f"({', '.join(_prem_meta['dates'])}) do not include the "
                    f"market date **{_market_date_text}** — the two lines would "
                    "describe different days."
                )
            if not _prem_points:
                st.error(
                    "No usable rows in the premium file with that mapping — "
                    f"{_prem_meta['dropped_time']} row(s) had no readable time "
                    f"and {_prem_meta['dropped_value']} no numeric value. "
                    "Check the Time and Premium columns above."
                )
            else:
                chart_payload["series"].append({
                    "name": "Premium", "axis": "right", "points": _prem_points,
                })
                st.caption(
                    f"💹 Premium: {format_indian(_prem_meta['used'])} minute(s) "
                    f"from {format_indian(_prem_meta['rows'])} row(s) "
                    f"({_prem_points[0][0]} → {_prem_points[-1][0]}, "
                    f"{_prem_points[0][1]:,.2f} → {_prem_points[-1][1]:,.2f})"
                    + (f" · {format_indian(_prem_meta['filtered_out'])} row(s) "
                       f"for other indexes ignored"
                       if _prem_meta["filtered_out"] else "")
                )
                with st.expander("Preview the parsed premium series"):
                    st.dataframe(
                        pd.DataFrame(_prem_points, columns=["Time", "Premium"]),
                        width="stretch", hide_index=True, height=240)

    if chart_payload is not None:
        chart_payloads.append(chart_payload)

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
                                        "when the index has a fetched day High / Low.")
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
# Build the Excel workbook and the shareable DOR.html. Both are rendered into
# `downloads_slot` at the top of the page once assembled.
# ---------------------------------------------------------------------------
workbook = Workbook()
workbook.remove(workbook.active)
add_report_sheets(workbook, tv_rows, used_multiplier,
                  summary_groups=([(f"summary {idx}"[:31], rows)
                                   for idx, rows in tv_by_segment.items()]
                                  or None))
add_strikes_sheet(workbook, strikes, expiry_map)
if orders_by_segment:
    for idx, rows in orders_by_segment.items():
        add_orders_sheet(workbook, rows, suffix=f" {idx}"[:24])
elif order_rows:
    add_orders_sheet(workbook, order_rows)
if pf_groups is not None:
    add_portfolio_sheet(workbook, portfolio_report(pf_groups, DEFAULT_PORTFOLIO_PATTERN))
_all_user_ref = dor_state.get("all_user_ref") or {}
seg.add_summary_sheets(workbook, dor_state["comp"], report_date,
                       all_user_ref=_all_user_ref)
if comp2 is not None:
    seg.add_summary_sheets(workbook, comp2, report_date, suffix=" 2",
                           all_user_ref=_all_user_ref)
excel_buffer = io.BytesIO()
workbook.save(excel_buffer)

# the DTE is part of the file name: the same date can be reported for 0DTE,
# 1DTE and 4DTE, and those are three different account scopes
_file_stem = f"DOR_{_used_dte}_{report_date or 'report'}"

# the report's slippage section carries the algo selection made in the
# inputs (same as the dashboard); the chain centres on the fetched day-mids;
# each MTM file gets its own segregation + Algo Summary section
dor_html = build_dor_html(
    report_date=report_date or "—",
    deviation=float(used_deviation),
    tv_totals=totals,
    seg_sections=seg_dor_sections,
    tv_sections=tv_dor_sections,
    strikes=strikes,
    slippage={"summary": slip_algo, "overall": slip_overall, "majors": slip_majors,
              "empty_note": slip_empty_note, "no_sl": slip_no_sl},
    portfolio=pf_groups,
    mids=mids,
    expiry_map=expiry_map,
    order_sections=order_dor_sections,
    excel_bytes=excel_buffer.getvalue(),
    excel_filename=f"{_file_stem}.xlsx",
    dte=_used_dte,
    charts=chart_payloads,
)

# rendered into the slot reserved just under the inputs, so both buttons are
# reachable without scrolling past the whole dashboard
with downloads_slot:
    st.subheader("Downloads")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "Download Excel — all data",
            data=excel_buffer.getvalue(),
            file_name=f"{_file_stem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
        )
        st.caption("tradevalue · summary · Strikes"
                   + (" · Portfolio QS" if pf_groups is not None else "")
                   + " · unclassified · Summary · MTM Data · Slippage · no_sl_Acc")
    with d2:
        st.download_button(
            "Download DOR.html — shareable summary report",
            data=dor_html.encode("utf-8"),
            file_name=f"{_file_stem}.html",
            mime="text/html",
            type="primary",
            width="stretch",
        )
        st.caption("Self-contained styled report (summary only) — "
                   "share with any user or client.")
    st.divider()
