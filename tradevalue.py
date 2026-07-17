"""Trade value report.

Reads a compiled orderbook (CSV/Excel) plus the compiled user MTM summary
(CSV/Excel) and produces tradevalue.csv: trade value / lots per
(date, server, user, segment), with each user's ALLOCATION pulled from the
summary sheet (matched on user id).

Rules (same as the original report):
  - only F&O rows count: Exchange must be BFO or NFO (drops NSE/BSE/MCX cash
    rows like NIFTYBEES-EQ, whose quantities aren't lot multiples);
  - only COMPLETE orders count;
  - duplicate orders are dropped on (user_id, date, order_id, exchg_order_id),
    keeping the first occurrence (lowest SNO);
  - symbols are classified into NIFTY / BANKNIFTY / SENSEX segments
    (anything else is ignored);
  - lots = abs(quantity) / lot size, using the date-based lot-size history;
  - trade value = sum(avg price * quantity).

Usage:
    python tradevalue.py                                 # auto-pick files in this folder
    python tradevalue.py orderbook.csv -s user_mtm.xlsx  # explicit inputs
    python tradevalue.py -o report.csv                   # custom output name
    streamlit run app.py                                 # browser UI with file upload
"""

import argparse
import csv
import io
import re
import sys
from collections import Counter, namedtuple
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
ORDERBOOK_PATTERN = "Compiled_Orderbook_*"
SUMMARY_PATTERN = "Compiled_User_MTM_*"
TABLE_SUFFIXES = (".csv", ".xlsx", ".xlsm", ".xls")
DEFAULT_OUTPUT = SCRIPT_DIR / "tradevalue.xlsx"

REPORT_HEADER = ["Date", "Server", "User ID", "Allocation", "Algo", "Segment",
                 "Order Count", "Quantity", "Lots", "Lot Pct", "Trade Value", "Outlier"]

SUMMARY_HEADER = ["Algo", "Total Users", "Avg Lot Pct", "Std Dev", "Band", "Below", "In", "Above"]

STRIKE_ALGO_HEADER = ["Algo", "Server", "Strikes Traded"]
STRIKE_CHAIN_HEADER = ["CE", "Strike", "PE"]

OUTLIER_CLIENTS_HEADER = ["Algo", "Server", "User ID", "Average Lots",
                          "Outlier", "Lots Fired", "Diff of Lots"]

# Outliers are judged per (date, algo) group by standard deviation: a row is
# "in range" when its lot pct lies within mean +/- k * std dev of the group's
# lot pct. Functions that take a `std_multiplier` accept k as a Decimal
# (e.g. 1, 1.5, 2); the default is 1 standard deviation.
DEFAULT_STD_MULTIPLIER = Decimal("1")


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def _norm_col(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").strip().lower())


def _pick_column(columns, *candidates):
    by_norm = {_norm_col(col): col for col in columns}
    for candidate in candidates:
        found = by_norm.get(_norm_col(candidate))
        if found:
            return found
    return None


def _decimal(value, default="0"):
    try:
        if value is None or str(value).strip().lower() in {"", "none", "nan", "na"}:
            return Decimal(default)
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return Decimal(default)


def _fmt_decimal(value, places=2):
    quant = Decimal("1") if places == 0 else Decimal("1").scaleb(-places)
    try:
        return f"{_decimal(value).quantize(quant):f}"
    except Exception:
        return str(value or "")


def _classify_symbol(symbol):
    upper = str(symbol or "").upper()
    if "BANKNIFTY" in upper or "BANK NIFTY" in upper:
        return "BANKNIFTY"
    if "SENSEX" in upper:
        return "SENSEX"
    if "NIFTY" in upper:
        return "NIFTY"
    return ""


# The SAME option contract circulates under two symbol formats (different
# brokers/servers), so both must normalise to one key or a contract would be
# counted as two distinct strikes:
#   NIFTY21JUL2624100PE — DDMMMYY expiry (day, month name, 2-digit year)
#   NIFTY2672124100PE   — YYMDD expiry (2-digit year, month digit, day;
#                         Oct/Nov/Dec are written as the single letter O/N/D)
_MONTH_NAMES = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN",
                "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
_MONTH_BY_NAME = {name: i + 1 for i, name in enumerate(_MONTH_NAMES)}
_MONTH_BY_CODE = {**{str(i): i for i in range(1, 10)}, "O": 10, "N": 11, "D": 12}
_CONTRACT_DDMMMYY = re.compile(
    r"^(BANKNIFTY|NIFTY|SENSEX)(\d{2})(" + "|".join(_MONTH_NAMES) + r")(\d{2})(\d+)(CE|PE)$")
_CONTRACT_YYMDD = re.compile(
    r"^(BANKNIFTY|NIFTY|SENSEX)(\d{2})([1-9OND])(\d{2})(\d+)(CE|PE)$")


def parse_contract(symbol):
    """Split an option symbol into its canonical contract key
    (segment, expiry date, strike, option type) — both expiry formats
    normalise to the same key. Returns None for anything that isn't a
    NIFTY / BANKNIFTY / SENSEX option (futures, equity, other formats)."""
    text = str(symbol or "").upper().replace(" ", "")
    match = _CONTRACT_DDMMMYY.match(text)
    if match:
        segment, dd, mon, yy, strike, opt = match.groups()
        month = _MONTH_BY_NAME[mon]
    else:
        match = _CONTRACT_YYMDD.match(text)
        if not match:
            return None
        segment, yy, mcode, dd, strike, opt = match.groups()
        month = _MONTH_BY_CODE[mcode]
    try:
        expiry = date(2000 + int(yy), month, int(dd))
    except ValueError:
        return None
    return (segment, expiry, int(strike), opt)


def contract_label(contract):
    """Display form of a contract key, e.g. "NIFTY 21JUL26 24100 PE"."""
    segment, expiry, strike, opt = contract
    return f"{segment} {expiry.strftime('%d%b%y').upper()} {strike} {opt}"


def _user_key(value):
    """Canonical user-id key for matching across files. Excel/CSV strips
    leading zeros from numeric ids ("04101961" becomes 4101961), so all-digit
    ids are keyed without leading zeros; matching is also case-insensitive."""
    key = str(value or "").strip().upper()
    if key.isdigit():
        return key.lstrip("0") or "0"
    return key


def _server_key(value):
    """Canonical server key for matching the orderbook against the MTM
    summary; a missing/NaN server reads as blank."""
    key = str(value or "").strip().upper()
    return "" if key in {"NAN", "NONE", "NA"} else key


_DATE_FORMATS = ("%d-%m-%Y", "%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y")
_DATE_PARSE_CACHE = {}


def _parse_date(value):
    raw = str(value or "").strip().split(" ")[0]
    if not raw:
        return None
    if raw in _DATE_PARSE_CACHE:
        return _DATE_PARSE_CACHE[raw]
    parsed = None
    for fmt in _DATE_FORMATS:
        try:
            parsed = datetime.strptime(raw, fmt).date()
            break
        except ValueError:
            continue
    _DATE_PARSE_CACHE[raw] = parsed
    return parsed


# ---------------------------------------------------------------------------
# Lot-size history (divisor applied to the traded quantity)
# ---------------------------------------------------------------------------

def _nifty_lot_size(d):
    # up to 26-Dec-2024 -> 25; 27-Dec-2024..30-Dec-2025 -> 75
    # (29/30-Jan-2025 -> 25, exception inside the 75 window); from 31-Dec-2025 -> 65
    if d is None:
        return Decimal("65")
    if d <= date(2024, 12, 26) or d in (date(2025, 1, 29), date(2025, 1, 30)):
        return Decimal("25")
    if d <= date(2025, 12, 30):
        return Decimal("75")
    return Decimal("65")


def _sensex_lot_size(d):
    # up to 03-Jan-2025 -> 10 (also 30/31-Jan-2025); otherwise -> 20
    if d is None:
        return Decimal("20")
    if d <= date(2025, 1, 3) or d in (date(2025, 1, 30), date(2025, 1, 31)):
        return Decimal("10")
    return Decimal("20")


def _banknifty_lot_size(d):
    # up to 26-Feb-2025 -> 15; 27-Feb-2025..26-Jun-2025 -> 30;
    # 27-Jun-2025..30-Dec-2025 -> 35; from 31-Dec-2025 -> 30
    if d is None:
        return Decimal("30")
    if d <= date(2025, 2, 26):
        return Decimal("15")
    if d <= date(2025, 6, 26):
        return Decimal("30")
    if d <= date(2025, 12, 30):
        return Decimal("35")
    return Decimal("30")


_LOT_SIZE_BY_SEGMENT = {
    "NIFTY": _nifty_lot_size,
    "SENSEX": _sensex_lot_size,
    "BANKNIFTY": _banknifty_lot_size,
}


# ---------------------------------------------------------------------------
# Generic table reading: CSV or Excel, from a path or an uploaded file object
# ---------------------------------------------------------------------------

def _cell_str(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _excel_rows(source):
    from openpyxl import load_workbook
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        # First sheet holds the per-row data (later sheets are summary views).
        for row in workbook.worksheets[0].iter_rows(values_only=True):
            yield [_cell_str(v) for v in row]
    finally:
        workbook.close()


def _csv_rows(source):
    if hasattr(source, "read"):
        raw = source.read()
        text = raw.decode("utf-8-sig", errors="replace") if isinstance(raw, bytes) else raw
        fh = io.StringIO(text)
    else:
        fh = open(source, newline="", encoding="utf-8-sig", errors="replace")
    try:
        yield from csv.reader(fh)
    finally:
        fh.close()


def _read_table(source, name=None):
    """Return (headers, row_iterator) for a CSV or Excel table. `source` may be
    a filesystem path or a file-like object (e.g. a Streamlit upload); `name`
    supplies the filename when `source` has none."""
    filename = str(name or getattr(source, "name", source)).lower()
    rows = _excel_rows(source) if filename.endswith((".xlsx", ".xlsm", ".xls")) else _csv_rows(source)
    try:
        headers = next(rows)
    except StopIteration:
        return [], iter(())
    return headers, rows


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return (row[idx] or "").strip()


# ---------------------------------------------------------------------------
# Reading the two inputs
# ---------------------------------------------------------------------------

Order = namedtuple(
    "Order",
    "rowid trade_date server user_id segment qty avg_price order_id exch_order_id symbol",
)


FO_EXCHANGES = {"BFO", "NFO"}


def read_orderbook(source, name=None):
    """Parse an orderbook CSV/Excel into Order tuples. Rows that never traded
    (status != COMPLETE), sit on a non-F&O exchange (anything but BFO / NFO —
    e.g. NSE cash rows like NIFTYBEES-EQ) or belong to none of the three
    segments are dropped."""
    headers, rows = _read_table(source, name)
    if not headers:
        return []

    def col(*candidates):
        found = _pick_column(headers, *candidates)
        return headers.index(found) if found else None

    i_date = col("date", "trade_date")
    i_server = col("server")
    i_user = col("user_id", "userid", "user id", "UserID")
    i_exchange = col("exchange", "exch", "exchg")
    i_symbol = col("symbol", "tradingsymbol", "trading_symbol")
    i_avg = col("avg_price", "avg price", "OrderAverageTradedPrice", "AveragePrice", "price")
    i_qty = col("quantity", "qty", "order_quantity", "OrderQuantity")
    i_status = col("status", "order_status", "OrderStatus")
    i_order_id = col("order_id", "orderid", "order id")
    i_exch_order_id = col("exchg_order_id", "exch_order_id", "exchange_order_id", "exchgorderid")
    i_rowid = col("row_id", "id", "sno")
    if i_symbol is None or i_avg is None or i_qty is None:
        raise ValueError("orderbook is missing the symbol / avg price / quantity columns")

    orders = []
    for line_no, raw in enumerate(rows, start=2):
        if i_exchange is not None and _cell(raw, i_exchange).upper() not in FO_EXCHANGES:
            continue
        if i_status is not None and _cell(raw, i_status).upper() != "COMPLETE":
            continue
        symbol = _cell(raw, i_symbol)
        segment = _classify_symbol(symbol)
        if not segment:
            continue
        try:
            rowid = int(_cell(raw, i_rowid))
        except (ValueError, TypeError):
            rowid = line_no
        orders.append(Order(
            rowid=rowid,
            trade_date=_parse_date(_cell(raw, i_date)),
            server=sys.intern(_cell(raw, i_server)),
            user_id=sys.intern(_cell(raw, i_user)),
            segment=segment,
            qty=_decimal(_cell(raw, i_qty)),
            avg_price=_decimal(_cell(raw, i_avg)),
            order_id=_cell(raw, i_order_id),
            exch_order_id=_cell(raw, i_exch_order_id),
            symbol=sys.intern(symbol),
        ))
    return orders


def read_allocations(source, name=None):
    """Read the compiled user MTM summary and return a lookup keyed by the
    canonical user id (see _user_key): {key: [{"allocation", "algo",
    "user_id", "server"}, ...]}. The same user id can legitimately run on
    more than one server — a distinct account with its own algo and
    allocation — so one entry is kept per (user, server) row and aggregate()
    resolves duplicates by the orderbook row's server. The original user id
    string is kept so the report can show the proper zero-padded form when
    the orderbook lost the leading zeros."""
    headers, rows = _read_table(source, name)
    if not headers:
        return {}
    user_col = _pick_column(headers, "user_id", "userid", "user id", "UserID")
    alloc_col = _pick_column(headers, "allocation")
    algo_col = _pick_column(headers, "algo")
    server_col = _pick_column(headers, "server")
    if user_col is None or alloc_col is None:
        raise ValueError("summary is missing the UserID / ALLOCATION columns")
    i_user = headers.index(user_col)
    i_alloc = headers.index(alloc_col)
    i_algo = headers.index(algo_col) if algo_col else None
    i_server = headers.index(server_col) if server_col else None

    allocations = {}
    for raw in rows:
        user = _cell(raw, i_user)
        key = _user_key(user)
        if not key:
            continue
        server = _server_key(_cell(raw, i_server)) if i_server is not None else ""
        entries = allocations.setdefault(key, [])
        if any(e["server"] == server for e in entries):
            continue  # true duplicate (same user AND server) — keep the first
        entries.append({
            "allocation": _decimal(_cell(raw, i_alloc)),
            "algo": _cell(raw, i_algo),
            "user_id": user,
            "server": server,
        })
    return allocations


def _pick_allocation(allocations, user_key, server):
    """Resolve a user's MTM entry: unambiguous when the user has one row;
    a user id that appears on several servers (distinct accounts) is matched
    on the orderbook row's server, and stays unmatched (blank allocation)
    when no server agrees rather than guessing the wrong account."""
    entries = allocations.get(user_key)
    if not entries:
        return None
    if len(entries) == 1:
        return entries[0]
    server = _server_key(server)
    return next((e for e in entries if e["server"] == server), None)


# ---------------------------------------------------------------------------
# Dedup, aggregation and output
# ---------------------------------------------------------------------------

# Order ids that went through Excel can come back in scientific notation
# ("2.60623E+13"), which collapses many DISTINCT orders onto one value.
# Such ids cannot be used as a dedup key.
_SCI_NOTATION = re.compile(r"^-?\d+(\.\d+)?E[+-]?\d+$", re.IGNORECASE)


def dedup_orders(rows):
    """Drop duplicate orders on (user_id, date, order_id, exchg_order_id),
    keeping the first occurrence (lowest row id). Rows without a usable order
    id (missing, or mangled into scientific notation by Excel) can't be keyed
    and are kept as-is."""
    best = {}
    keyless = []
    for row in rows:
        if not row.order_id or _SCI_NOTATION.match(row.order_id):
            keyless.append(row)
            continue
        key = (_user_key(row.user_id), row.trade_date, row.order_id, row.exch_order_id)
        current = best.get(key)
        if current is None or row.rowid < current.rowid:
            best[key] = row
    return list(best.values()) + keyless


def aggregate(rows, allocations=None, std_multiplier=None):
    """Group orders by (date, server, user, segment) and attach each user's
    allocation from the summary (None when the user isn't in the summary).
    Users are grouped on their canonical id, so "4101961" and "04101961"
    are the same user; the report shows the zero-padded form when known."""
    allocations = allocations or {}
    groups = {}
    display = {}
    for row in rows:
        user_key = _user_key(row.user_id)
        shown = display.get(user_key)
        if shown is None or (row.user_id.isdigit() and len(row.user_id) > len(shown)):
            display[user_key] = row.user_id
        key = (row.trade_date, row.server, user_key, row.segment)
        group = groups.get(key)
        if group is None:
            group = groups[key] = {
                "order_count": 0,
                "quantity": Decimal("0"),
                "lots": Decimal("0"),
                "trade_value": Decimal("0"),
            }
        lot_size = _LOT_SIZE_BY_SEGMENT[row.segment](row.trade_date)
        group["order_count"] += 1
        group["quantity"] += row.qty
        group["lots"] += abs(row.qty) / lot_size
        group["trade_value"] += row.avg_price * row.qty

    out = []
    for (trade_date, server, user_key, segment), group in groups.items():
        entry = _pick_allocation(allocations, user_key, server)
        user_id = display[user_key]
        if entry:
            # Prefer the summary's zero-padded id over a leading-zero-stripped one.
            mtm_id = entry["user_id"]
            if mtm_id.isdigit() and user_id.isdigit() and len(mtm_id) > len(user_id):
                user_id = mtm_id
        out.append({
            "trade_date": trade_date,
            "server": server,
            "user_id": user_id,
            "allocation": entry["allocation"] if entry else None,
            "algo": entry["algo"] if entry else "",
            "segment": segment,
            "order_count": group["order_count"],
            "quantity": group["quantity"],
            "lots": group["lots"],
            "trade_value": group["trade_value"],
        })
    # date DESC, then server, user_id, segment ASC
    out.sort(key=lambda r: (r["server"], r["user_id"], r["segment"]))
    out.sort(key=lambda r: r["trade_date"] or date.min, reverse=True)
    return _attach_lot_metrics(out, std_multiplier)


def _lot_pct_stats(rows, std_multiplier=None):
    """Per (date, algo) statistics of lot_pct: mean, population std dev, and
    the outlier band mean +/- k * std dev. Returns
    {(date, algo): (mean, std, low, high)}."""
    k = std_multiplier if std_multiplier is not None else DEFAULT_STD_MULTIPLIER
    groups = {}
    for row in rows:
        if row.get("lot_pct") is None or not row["algo"]:
            continue
        groups.setdefault((row["trade_date"], row["algo"]), []).append(row["lot_pct"])

    stats = {}
    for key, pcts in groups.items():
        n = len(pcts)
        mean = sum(pcts) / n
        variance = sum((p - mean) ** 2 for p in pcts) / n
        std = variance.sqrt()
        stats[key] = (mean, std, mean - k * std, mean + k * std)
    return stats


def user_lot_observations(rows, std_multiplier=None):
    """One observation per (date, algo, user): the OUTLIER UNIT IS THE USER.

    A report row is per (user, segment), so a user trading NIFTY and SENSEX
    would otherwise be judged twice on partial exposures. Here the user's
    lots are summed across all their segments (and servers) inside the algo
    and divided by their TOTAL allocation (one account per server):

        lot_pct = total lots / total allocation * 100

    Each observation carries the (date, algo) band (mean ± k·σ over the
    per-user lot pcts) and its outlier flag; users with no usable allocation
    get a blank flag. Rows without an algo produce no observation."""
    obs_map = {}
    for row in rows:
        if not row["algo"]:
            continue
        key = (row["trade_date"], row["algo"], _user_key(row["user_id"]))
        o = obs_map.get(key)
        if o is None:
            o = obs_map[key] = {
                "trade_date": row["trade_date"],
                "algo": row["algo"],
                "user_key": key[2],
                "user_id": row["user_id"],
                "lots": Decimal("0"),
                "_alloc_by_server": {},
                "_servers": [],
            }
        o["lots"] += row["lots"]
        server = _server_key(row["server"])
        if server not in o["_servers"]:
            o["_servers"].append(server)
        if row["allocation"] is not None:
            o["_alloc_by_server"][server] = row["allocation"]

    out = list(obs_map.values())
    for o in out:
        alloc_by_server = o.pop("_alloc_by_server")
        allocation = sum(alloc_by_server.values(), Decimal("0")) if alloc_by_server else None
        o["allocation"] = allocation
        o["lot_pct"] = (o["lots"] / allocation * 100) if allocation else None
        o["server"] = " / ".join(s for s in o.pop("_servers") if s)

    stats = _lot_pct_stats(out, std_multiplier)
    for o in out:
        stat = stats.get((o["trade_date"], o["algo"]))
        if o["lot_pct"] is None or stat is None:
            o["outlier"], o["band_low"], o["band_high"] = "", None, None
            continue
        _, _, low, high = stat
        o["band_low"], o["band_high"] = low, high
        if o["lot_pct"] < low:
            o["outlier"] = "Below average range"
        elif o["lot_pct"] > high:
            o["outlier"] = "Above average range"
        else:
            o["outlier"] = "In range"
    return out


def _attach_lot_metrics(rows, std_multiplier=None):
    """Add lot_pct to every row (that ROW's lots vs the account allocation —
    it documents the row), then judge outliers PER USER (combined lots across
    the user's segments/servers ÷ combined allocation, see
    user_lot_observations) and stamp the user's flag on each of their rows.
    Rows with no allocation or no algo can't be judged and stay blank."""
    for row in rows:
        allocation = row["allocation"]
        row["lot_pct"] = (row["lots"] / allocation * 100) if allocation else None

    flags = {(o["trade_date"], o["algo"], o["user_key"]): o["outlier"]
             for o in user_lot_observations(rows, std_multiplier)}
    for row in rows:
        row["outlier"] = flags.get(
            (row["trade_date"], row["algo"], _user_key(row["user_id"])), "")
    return rows


def format_row(row):
    return [
        row["trade_date"].strftime("%d-%m-%Y") if row["trade_date"] else "",
        row["server"],
        row["user_id"],
        _fmt_decimal(row["allocation"], places=0) if row["allocation"] is not None else "",
        row.get("algo", ""),
        row["segment"],
        row["order_count"],
        _fmt_decimal(row["quantity"], places=0),
        _fmt_decimal(row["lots"], places=0),
        _fmt_decimal(row["lot_pct"]) if row.get("lot_pct") is not None else "",
        _fmt_decimal(row["trade_value"]),
        row.get("outlier", ""),
    ]


def algo_summary(rows, std_multiplier=None):
    """Per (date, algo) outlier summary, ALL columns in users: "Total Users"
    (distinct users of the algo), the per-USER lot pct statistics (mean, std
    dev, mean +/- k*std band — see user_lot_observations), and how many USERS
    fall below / in / above the band, so Below + In + Above = Total Users
    (users with no usable allocation carry no lot pct and are the only ones
    that can fall outside the three flag columns). Pass the SAME
    std_multiplier that was given to aggregate() so the band shown matches
    the flags on the rows."""
    observations = user_lot_observations(rows, std_multiplier)
    stats = _lot_pct_stats(observations, std_multiplier)
    counts = {}
    for o in observations:
        key = (o["trade_date"], o["algo"])
        group = counts.setdefault(key, {"users": 0, "below": 0, "in_range": 0, "above": 0})
        group["users"] += 1
        if o["outlier"] == "Below average range":
            group["below"] += 1
        elif o["outlier"] == "Above average range":
            group["above"] += 1
        elif o["outlier"] == "In range":
            group["in_range"] += 1

    out = []
    for (trade_date, algo), group in counts.items():
        # a group whose users all lack an allocation has no lot-pct statistics
        mean, std, low, high = stats.get((trade_date, algo), (None, None, None, None))
        out.append({
            "trade_date": trade_date,
            "algo": algo,
            "users": group["users"],
            "avg_lot_pct": mean,
            "std_dev": std,
            "band_low": low,
            "band_high": high,
            "below": group["below"],
            "in_range": group["in_range"],
            "above": group["above"],
        })
    out.sort(key=lambda r: (0, int(r["algo"])) if str(r["algo"]).isdigit() else (1, str(r["algo"])))
    out.sort(key=lambda r: r["trade_date"] or date.min, reverse=True)
    return out


def format_summary_row(row):
    has_stats = row["avg_lot_pct"] is not None
    return [
        row["algo"],
        row["users"],
        _fmt_decimal(row["avg_lot_pct"]) if has_stats else "",
        _fmt_decimal(row["std_dev"]) if has_stats else "",
        f"{_fmt_decimal(row['band_low'])}–{_fmt_decimal(row['band_high'])}" if has_stats else "",
        row["below"],
        row["in_range"],
        row["above"],
    ]


def outlier_clients(rows, std_multiplier):
    """The stricter "outlier clients" view: USERS whose per-user lot pct
    (combined lots ÷ combined allocation, see user_lot_observations) falls
    outside mean ± k·σ of their (date, algo) group, where k is the CLIENT
    deviation — asked separately from (and typically wider than) the flagging
    deviation, e.g. 2 vs 1. One row per user — a user trading two segments is
    judged once on their combined exposure.

    The group band is a lot-pct band, so it is translated into LOTS through
    the user's own allocation: "Average Lots" is the in-range lots window for
    THAT user (lower edge clamped at 0 — nobody can fire negative lots) and
    "Diff of Lots" is how many lots beyond the nearest edge the user actually
    fired. Sorted worst-first (largest lots difference) within date/algo."""
    out = []
    for o in user_lot_observations(rows, std_multiplier):
        if o["outlier"] not in ("Below average range", "Above average range"):
            continue
        allocation = o["allocation"]
        low_lots = max(o["band_low"], Decimal("0")) * allocation / 100
        high_lots = o["band_high"] * allocation / 100
        above = o["outlier"] == "Above average range"
        out.append({
            "trade_date": o["trade_date"],
            "algo": o["algo"],
            "server": o["server"],
            "user_id": o["user_id"],
            "band_low_lots": low_lots,
            "band_high_lots": high_lots,
            "outlier": o["outlier"],
            "lots": o["lots"],
            "diff_lots": (o["lots"] - high_lots) if above else (low_lots - o["lots"]),
        })
    out.sort(key=lambda r: (
        (0, int(r["algo"])) if str(r["algo"]).isdigit() else (1, str(r["algo"])),
        -r["diff_lots"],
    ))
    out.sort(key=lambda r: r["trade_date"] or date.min, reverse=True)
    return out


def format_lots_band(row):
    """Display form of an outlier row's expected-lots window, e.g. "400 – 500"."""
    return f"{int(round(row['band_low_lots'])):,} – {int(round(row['band_high_lots'])):,}"


def add_outlier_clients_sheet(workbook, outlier_rows):
    """Append the "Outlier Clients" sheet: one row per report row outside the
    client deviation band, band and difference in whole lots."""
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    sheet = workbook.create_sheet("Outlier Clients")
    sheet.append(OUTLIER_CLIENTS_HEADER)
    for row in outlier_rows:
        sheet.append([
            row["algo"],
            row["server"],
            row["user_id"],
            format_lots_band(row),
            row["outlier"],
            int(round(row["lots"])),
            int(round(row["diff_lots"])),
        ])
    bold = Font(bold=True)
    center = Alignment(horizontal="center")
    for col_idx, title in enumerate(OUTLIER_CLIENTS_HEADER, start=1):
        sheet.cell(row=1, column=col_idx).font = bold
        sheet.column_dimensions[get_column_letter(col_idx)].width = max(len(title) + 6, 14)
    lots_col = OUTLIER_CLIENTS_HEADER.index("Lots Fired") + 1
    diff_col = OUTLIER_CLIENTS_HEADER.index("Diff of Lots") + 1
    for r in range(1, sheet.max_row + 1):
        for c in range(1, len(OUTLIER_CLIENTS_HEADER) + 1):
            sheet.cell(r, c).alignment = center
        if r > 1:
            sheet.cell(r, lots_col).number_format = "#,##0"
            sheet.cell(r, diff_col).number_format = "#,##0"
    sheet.freeze_panes = "A2"


def strike_report(orders, allocations=None):
    """The two strike summaries, computed over (deduped) orders:

      * by_algo_server — per (algo, server): how many DISTINCT contracts
        (segment, expiry, strike, CE/PE) were traded. The algo comes from the
        same MTM allocation matching as the trade value rows; orders whose
        user has no MTM entry land in a blank-algo bucket.
      * per_strike — per contract: total lots (sum |qty| / lot size, the same
        date-based lot math as the trade value rows) and order count.

    Orders whose symbol doesn't parse as an option contract are skipped."""
    allocations = allocations or {}
    # Server -> algo fallback for users with no MTM entry: a server's algo is
    # the algo of its MTM accounts (the most common one, should they ever
    # mix), so an unmatched user doesn't spawn a duplicate blank-algo row
    # for a server that already sits under an algo.
    server_algo_votes = {}
    for entries in allocations.values():
        for entry in entries:
            if entry["algo"]:
                server_algo_votes.setdefault(entry["server"], Counter())[entry["algo"]] += 1
    server_algo = {server: votes.most_common(1)[0][0]
                   for server, votes in server_algo_votes.items()}

    parse_cache = {}
    by_algo_server = {}
    per_strike = {}
    for row in orders:
        if row.symbol in parse_cache:
            contract = parse_cache[row.symbol]
        else:
            contract = parse_cache[row.symbol] = parse_contract(row.symbol)
        if contract is None:
            continue
        entry = _pick_allocation(allocations, _user_key(row.user_id), row.server)
        server_key = _server_key(row.server)
        algo = entry["algo"] if entry else server_algo.get(server_key, "")
        by_algo_server.setdefault((algo, server_key), set()).add(contract)
        lot_size = _LOT_SIZE_BY_SEGMENT[row.segment](row.trade_date)
        group = per_strike.setdefault(contract, {"lots": Decimal("0"), "order_count": 0})
        group["lots"] += abs(row.qty) / lot_size
        group["order_count"] += 1

    # each row keeps its contract set so the count can be re-filtered by
    # expiry / index in the dashboard and DOR.html
    algo_rows = [
        {"algo": algo, "server": server, "strike_count": len(contracts),
         "contracts": sorted(contracts)}
        for (algo, server), contracts in by_algo_server.items()
    ]
    # numeric algos first in numeric order, then text algos, blank (unmatched) last
    algo_rows.sort(key=lambda r: (
        (2, "") if r["algo"] == "" else
        (0, int(r["algo"])) if str(r["algo"]).isdigit() else (1, str(r["algo"])),
        r["server"],
    ))
    strike_rows = [
        {"segment": c[0], "expiry": c[1], "strike": c[2], "opt_type": c[3],
         "label": contract_label(c), "lots": g["lots"], "order_count": g["order_count"]}
        for c, g in per_strike.items()
    ]
    strike_rows.sort(key=lambda r: (r["segment"], r["expiry"], r["strike"], r["opt_type"]))
    return {"by_algo_server": algo_rows, "per_strike": strike_rows}


def strike_chain(per_strike):
    """Pivot the per-strike rows into an option-chain view:
    {(segment, expiry): [(ce_lots, strike, pe_lots), ...]} — one row per
    strike price, sorted by strike, CE and PE lots side by side (0 when only
    one side traded). Keys come out sorted by (segment, expiry)."""
    by_key = {}
    for row in per_strike:
        sides = by_key.setdefault((row["segment"], row["expiry"]), {}) \
                      .setdefault(row["strike"], {"CE": Decimal("0"), "PE": Decimal("0")})
        sides[row["opt_type"]] += row["lots"]
    return {
        key: [(sides[s]["CE"], s, sides[s]["PE"]) for s in sorted(sides)]
        for key, sides in sorted(by_key.items())
    }


def add_strikes_sheet(workbook, strikes):
    """Append the "Strikes" sheet: the per-(algo, server) distinct strike
    counts in columns A-C, and one option-chain block (CE | Strike | PE) per
    (segment, expiry) stacked from column E. Lots are whole numbers."""
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    sheet = workbook.create_sheet("Strikes")
    bold = Font(bold=True)

    for col_idx, title in enumerate(STRIKE_ALGO_HEADER, start=1):
        cell = sheet.cell(row=1, column=col_idx, value=title)
        cell.font = bold
        sheet.column_dimensions[get_column_letter(col_idx)].width = max(len(title) + 4, 12)
    for r, row in enumerate(strikes["by_algo_server"], start=2):
        sheet.cell(row=r, column=1, value=row["algo"])
        sheet.cell(row=r, column=2, value=row["server"])
        sheet.cell(row=r, column=3, value=row["strike_count"])
    # a per-(algo, server) count can share strikes with another server, so the
    # overall figure is the distinct count, not the column sum
    total_row = len(strikes["by_algo_server"]) + 2
    sheet.cell(row=total_row, column=1, value="Total (distinct)").font = bold
    sheet.cell(row=total_row, column=3, value=len(strikes["per_strike"])).font = bold

    offset = len(STRIKE_ALGO_HEADER) + 2  # leave column D blank between the tables
    for j in range(3):
        sheet.column_dimensions[get_column_letter(offset + j)].width = 12
    r = 1
    for (segment, expiry), rows in strike_chain(strikes["per_strike"]).items():
        sheet.merge_cells(start_row=r, start_column=offset, end_row=r, end_column=offset + 2)
        title = sheet.cell(row=r, column=offset,
                           value=f"{segment} {expiry.strftime('%d%b%y').upper()}")
        title.font = bold
        r += 1
        for j, header in enumerate(STRIKE_CHAIN_HEADER):
            sheet.cell(row=r, column=offset + j, value=header).font = bold
        r += 1
        ce_total = pe_total = Decimal("0")
        for ce, strike, pe in rows:
            sheet.cell(row=r, column=offset, value=int(round(ce))).number_format = "#,##0"
            sheet.cell(row=r, column=offset + 1, value=strike)
            sheet.cell(row=r, column=offset + 2, value=int(round(pe))).number_format = "#,##0"
            ce_total += ce
            pe_total += pe
            r += 1
        for j, value in enumerate((int(round(ce_total)), "Total", int(round(pe_total)))):
            cell = sheet.cell(row=r, column=offset + j, value=value)
            cell.font = bold
            if j != 1:
                cell.number_format = "#,##0"
        r += 2  # blank spacer row between chains
    sheet.freeze_panes = "A2"


def report_csv_text(rows):
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(REPORT_HEADER)
    for row in rows:
        writer.writerow(format_row(row))
    return buffer.getvalue()


def _excel_number(value, places=2):
    if value is None:
        return None
    return round(float(value), places)


def _excel_report_row(row):
    # User id stays a string cell so Excel keeps leading zeros ("04101961").
    return [
        row["trade_date"].strftime("%d-%m-%Y") if row["trade_date"] else "",
        row["server"],
        row["user_id"],
        _excel_number(row["allocation"], places=0),
        row["algo"],
        row["segment"],
        row["order_count"],
        _excel_number(row["quantity"], places=0),
        _excel_number(row["lots"], places=0),
        _excel_number(row.get("lot_pct")),
        _excel_number(row["trade_value"]),
        row.get("outlier", ""),
    ]


def add_report_sheets(workbook, rows, std_multiplier=None):
    """Append the "tradevalue" and "summary" sheets to an existing openpyxl
    workbook (used to combine this report with other reports in one file)."""
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    sheet = workbook.create_sheet("tradevalue")
    sheet.append(REPORT_HEADER)
    for row in rows:
        sheet.append(_excel_report_row(row))

    summary_sheet = workbook.create_sheet("summary")
    summary_sheet.append(SUMMARY_HEADER)
    for srow in algo_summary(rows, std_multiplier):
        has_stats = srow["avg_lot_pct"] is not None
        summary_sheet.append([
            srow["algo"],
            srow["users"],
            _excel_number(srow["avg_lot_pct"]),
            _excel_number(srow["std_dev"]),
            (f"{_fmt_decimal(srow['band_low'])}–{_fmt_decimal(srow['band_high'])}"
             if has_stats else ""),
            srow["below"],
            srow["in_range"],
            srow["above"],
        ])

    bold = Font(bold=True)
    for ws, header in ((sheet, REPORT_HEADER), (summary_sheet, SUMMARY_HEADER)):
        for col_idx, title in enumerate(header, start=1):
            ws.cell(row=1, column=col_idx).font = bold
            ws.column_dimensions[get_column_letter(col_idx)].width = max(len(title) + 4, 12)
        ws.freeze_panes = "A2"


def write_report_excel(rows, target, std_multiplier=None, strikes=None, outliers=None):
    """Write the report workbook: sheet "tradevalue" holds the full report,
    sheet "summary" the per-algo outlier table, plus the "Strikes" /
    "Outlier Clients" sheets when their data is given. `target` may be a
    filesystem path or a writable buffer."""
    from openpyxl import Workbook

    workbook = Workbook()
    workbook.remove(workbook.active)
    add_report_sheets(workbook, rows, std_multiplier)
    if strikes:
        add_strikes_sheet(workbook, strikes)
    if outliers is not None:
        add_outlier_clients_sheet(workbook, outliers)
    workbook.save(target)


def report_excel_bytes(rows, std_multiplier=None, strikes=None, outliers=None):
    buffer = io.BytesIO()
    write_report_excel(rows, buffer, std_multiplier, strikes, outliers)
    return buffer.getvalue()


def write_report(rows, output_path, std_multiplier=None, strikes=None, outliers=None):
    if str(output_path).lower().endswith(".csv"):
        with open(output_path, "w", newline="", encoding="utf-8") as fh:
            fh.write(report_csv_text(rows))
    else:
        write_report_excel(rows, output_path, std_multiplier, strikes, outliers)


def report_totals(rows):
    return {
        "users": len({r["user_id"] for r in rows if r["user_id"]}),
        "orders": sum(r["order_count"] for r in rows),
        "lots": sum((r["lots"] for r in rows), Decimal("0")),
        "trade_value": sum((r["trade_value"] for r in rows), Decimal("0")),
    }


def build_report(orderbook_source, summary_source, orderbook_name=None, summary_name=None):
    """Full pipeline: read both inputs, dedup, aggregate. Returns report rows."""
    orders = read_orderbook(orderbook_source, orderbook_name)
    allocations = read_allocations(summary_source, summary_name) if summary_source is not None else {}
    return aggregate(dedup_orders(orders), allocations)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _default_files(pattern):
    return [p for p in sorted(SCRIPT_DIR.glob(pattern)) if p.suffix.lower() in TABLE_SUFFIXES]


def main(argv=None):
    parser = argparse.ArgumentParser(description="Generate the trade value report from orderbook + user MTM files.")
    parser.add_argument("inputs", nargs="*", type=Path,
                        help=f"orderbook CSV/Excel file(s); default: {ORDERBOOK_PATTERN} in the script folder")
    parser.add_argument("-s", "--summary", type=Path,
                        help=f"compiled user MTM summary (CSV/Excel) for the Allocation column; "
                             f"default: {SUMMARY_PATTERN} in the script folder")
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT,
                        help=f"output path, .xlsx or .csv (default: {DEFAULT_OUTPUT.name})")
    parser.add_argument("-d", "--deviation", type=float, default=1.0,
                        help="outlier threshold in standard deviations of the algo's lot pct (default 1)")
    parser.add_argument("-c", "--client-deviation", type=float, default=2.0,
                        help="stricter threshold for the Outlier Clients sheet (default 2)")
    args = parser.parse_args(argv)

    if args.deviation <= 0:
        parser.error("--deviation must be greater than 0")
    if args.client_deviation <= 0:
        parser.error("--client-deviation must be greater than 0")
    std_multiplier = Decimal(str(args.deviation))
    client_multiplier = Decimal(str(args.client_deviation))

    inputs = args.inputs or _default_files(ORDERBOOK_PATTERN)
    if not inputs:
        parser.error(f"no orderbook files found ({ORDERBOOK_PATTERN})")
    missing = [p for p in inputs if not p.is_file()]
    if missing:
        parser.error("input file not found: " + ", ".join(str(p) for p in missing))

    summary = args.summary
    if summary is None:
        candidates = _default_files(SUMMARY_PATTERN)
        summary = candidates[0] if candidates else None
        if summary is None:
            print(f"WARNING: no summary file found ({SUMMARY_PATTERN}); Allocation column will be empty")
    elif not summary.is_file():
        parser.error(f"summary file not found: {summary}")

    orders = []
    for path in inputs:
        file_orders = read_orderbook(path)
        print(f"{path.name}: {len(file_orders)} completed NIFTY/BANKNIFTY/SENSEX orders")
        orders.extend(file_orders)

    allocations = {}
    if summary is not None:
        allocations = read_allocations(summary)
        print(f"{summary.name}: allocations for {len(allocations)} users")

    deduped = dedup_orders(orders)
    dropped = len(orders) - len(deduped)
    if dropped:
        print(f"dropped {dropped} duplicate orders")

    report_rows = aggregate(deduped, allocations, std_multiplier)
    strikes = strike_report(deduped, allocations)
    outliers = outlier_clients(report_rows, client_multiplier)
    write_report(report_rows, args.output, std_multiplier, strikes, outliers)

    totals = report_totals(report_rows)
    matched = len({r["user_id"] for r in report_rows if r["allocation"] is not None})
    print(f"wrote {len(report_rows)} rows to {args.output}")
    print(f"allocation matched for {matched}/{totals['users']} users")
    print(f"totals: {totals['users']} users | lots {_fmt_decimal(totals['lots'], places=0)} "
          f"| trade value {_fmt_decimal(totals['trade_value'])}")
    print(f"strikes: {len(strikes['per_strike'])} distinct contracts")
    print(f"outlier clients (±{args.client_deviation:g}σ): {len(outliers)} rows")


if __name__ == "__main__":
    main()
