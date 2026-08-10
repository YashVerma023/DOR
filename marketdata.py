"""Index market data for the report date.

Supplies the day High / Low of NIFTY, BANKNIFTY and SENSEX, which the strikes
section turns into the option chain's ATM anchor:

    day mid = (High + Low) / 2

and returns the full OHLC alongside, plus an optional intraday series, so the
DOR's charts can be drawn from the same fetch rather than a second one.

Source
------
TradingView has no official public API for historical OHLC — its only official
offering is the Charting Library, where the caller supplies the data. The
unofficial websocket scrapers break whenever TradingView changes its internals
and need a TradingView login. yfinance is used instead: same index values, a
documented library, and no credentials.

    NIFTY      ^NSEI
    BANKNIFTY  ^NSEBANK
    SENSEX     ^BSESN

Every call is best-effort. A market holiday, a weekend, a future date or a
dropped network connection yields no data for that index and an explanatory
message — never an exception — so the dashboard can offer manual entry and the
report still builds.
"""

import logging
from datetime import date, datetime, timedelta

logger = logging.getLogger(__name__)

# Report index name -> Yahoo Finance ticker.
INDEX_SYMBOLS = {
    "NIFTY": "^NSEI",
    "BANKNIFTY": "^NSEBANK",
    "SENSEX": "^BSESN",
}

# Short forms used on the chart, where "Index (SX)" is far more readable at a
# hover point than the full name.
INDEX_ABBR = {"NIFTY": "NF", "BANKNIFTY": "BNF", "SENSEX": "SX"}

# Intraday granularity offered for the charts. Yahoo only serves 1m for the
# last ~30 days and 5m/15m for ~60, so a request for an older date returns
# nothing at these intervals while the daily bar still resolves.
INTRADAY_INTERVALS = ("1m", "5m", "15m")

OHLC_FIELDS = ("open", "high", "low", "close")


def _as_date(value):
    """Accept a date, datetime or ISO / dd-mm-YYYY string."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"unrecognised date: {value!r}")


def day_mid(levels):
    """(High + Low) / 2 for one index's level dict, or None when either side
    is missing. Kept as a function so the ATM basis is defined in exactly one
    place."""
    if not levels:
        return None
    high, low = levels.get("high"), levels.get("low")
    if high is None or low is None:
        return None
    return (float(high) + float(low)) / 2


def fetch_index_levels(day, indexes=None, interval=None):
    """({index: levels}, {index: reason}) for the given trading day.

    `levels` carries open / high / low / close / mid / volume / symbol, and
    `series` (a list of {t, open, high, low, close} bars) when `interval` is
    one of INTRADAY_INTERVALS. `reason` explains every index that produced no
    data, so the caller can prompt for manual entry.

    Never raises: an import failure, a network error or an empty response all
    come back as a reason string.
    """
    wanted = list(indexes or INDEX_SYMBOLS)
    try:
        day = _as_date(day)
    except ValueError as exc:
        return {}, {name: str(exc) for name in wanted}

    try:
        import yfinance as yf
    except ImportError:
        reason = ("yfinance is not installed — run `pip install yfinance` "
                  "or enter the levels manually")
        logger.error("market data unavailable: %s", reason)
        return {}, {name: reason for name in wanted}

    if day > date.today():
        reason = f"{day:%d-%m-%Y} is in the future"
        return {}, {name: reason for name in wanted}

    levels, problems = {}, {}
    # yfinance's end is exclusive, so ask for a single day as [day, day+1)
    start, end = day.isoformat(), (day + timedelta(days=1)).isoformat()

    for name in wanted:
        symbol = INDEX_SYMBOLS.get(name)
        if symbol is None:
            problems[name] = f"no ticker mapped for {name}"
            continue
        try:
            frame = yf.Ticker(symbol).history(start=start, end=end)
        except Exception as exc:                      # network, parse, rate-limit
            problems[name] = f"fetch failed ({type(exc).__name__}: {exc})"
            logger.warning("%s (%s): fetch failed — %s", name, symbol, exc)
            continue

        if frame is None or frame.empty:
            problems[name] = (f"no data for {day:%d-%m-%Y} — market holiday, "
                              "weekend, or the date has not settled yet")
            logger.warning("%s (%s): no bar for %s", name, symbol, day)
            continue

        row = frame.iloc[0]
        entry = {"symbol": symbol}
        for field in OHLC_FIELDS:
            value = row.get(field.capitalize())
            entry[field] = None if value is None else float(value)
        volume = row.get("Volume")
        entry["volume"] = None if volume is None else float(volume)
        entry["mid"] = day_mid(entry)
        if interval in INTRADAY_INTERVALS:
            entry["series"] = _intraday(yf, symbol, start, end, interval, name)
        levels[name] = entry
        logger.info("%s (%s) %s: H=%.2f L=%.2f mid=%.2f", name, symbol, day,
                    entry["high"], entry["low"], entry["mid"])

    return levels, problems


def _intraday(yf, symbol, start, end, interval, name):
    """Intraday bars for the charts; [] when the interval is out of Yahoo's
    retention window for that date. Never raises."""
    try:
        frame = yf.Ticker(symbol).history(start=start, end=end, interval=interval)
    except Exception as exc:
        logger.warning("%s: intraday %s unavailable — %s", name, interval, exc)
        return []
    if frame is None or frame.empty:
        logger.info("%s: no %s bars for that date (outside Yahoo's retention)",
                    name, interval)
        return []
    return [
        {
            "t": idx.strftime("%H:%M"),
            "open": float(row.Open), "high": float(row.High),
            "low": float(row.Low), "close": float(row.Close),
        }
        for idx, row in frame.iterrows()
    ]


# ---------------------------------------------------------------------------
# Premium upload — the ATM straddle series, supplied as a file
# ---------------------------------------------------------------------------

# Header fragments used to GUESS which column is which. They are only a first
# offer: the dashboard always lets the mapping be corrected, because a wrong
# auto-detect that silently plots the wrong column is far worse than asking.
_TIME_HINTS = ("time", "timestamp", "datetime", "date time", "bucket", "minute")
_VALUE_HINTS = ("premium", "straddle", "atm", "total", "value", "close", "ltp",
                "price", "sum")
_DATE_HINTS = ("date", "day")
_INDEX_HINTS = ("index", "symbol", "instrument", "underlying", "scrip")


def _norm(text):
    return "".join(ch for ch in str(text).lower() if ch.isalnum() or ch == " ").strip()


def _guess(columns, hints, exclude=()):
    """First column whose header contains any hint, preferring an exact match
    and skipping anything already claimed by another role."""
    cols = [c for c in columns if c not in exclude]
    for hint in hints:
        for col in cols:
            if _norm(col) == hint:
                return col
    for hint in hints:
        for col in cols:
            if hint in _norm(col):
                return col
    return None


def read_premium(source, name=None):
    """Read the premium upload and offer a column mapping.

    Returns (dataframe, guesses) where guesses is {"time", "value", "date",
    "index"} — each either a column name or None. Nothing is parsed yet; the
    caller confirms or corrects the mapping first.
    """
    import pandas as pd

    filename = str(name or getattr(source, "name", "") or "").lower()
    if filename.endswith(".csv"):
        df = pd.read_csv(source)
    else:
        df = pd.read_excel(source, sheet_name=0)
    df = df.dropna(how="all").dropna(axis=1, how="all")

    cols = list(df.columns)
    time_col = _guess(cols, _TIME_HINTS)
    date_col = _guess(cols, _DATE_HINTS, exclude={time_col})
    index_col = _guess(cols, _INDEX_HINTS, exclude={time_col, date_col})
    value_col = _guess(cols, _VALUE_HINTS, exclude={time_col, date_col, index_col})
    if value_col is None:
        # fall back to the first numeric column that is not already claimed
        claimed = {time_col, date_col, index_col}
        for col in cols:
            if col in claimed:
                continue
            if pd.to_numeric(df[col], errors="coerce").notna().any():
                value_col = col
                break
    return df, {"time": time_col, "value": value_col,
                "date": date_col, "index": index_col}


def detect_index(df, guesses=None, filename=None, strike_col=None):
    """Work out which index a premium file describes, without asking.

    Tried in order of how directly each states the index:

      1. an index/symbol column      SENSEX, BANKNIFTY 06AUG26 76000 PE, …
      2. the file name               SENSEX_premium_05-08-2026.csv
      3. any text column at all      a stray header or label cell
      4. strike magnitude            NIFTY ~24k, BANKNIFTY ~57k, SENSEX ~79k

    Returns (index name, how it was found) or (None, None). BANKNIFTY is
    always tested before NIFTY, since "BANKNIFTY" contains "NIFTY".
    """
    import pandas as pd

    from tradevalue import _classify_symbol

    guesses = guesses or {}

    def from_values(series):
        hits = {}
        for value in series.dropna().astype(str).head(500):
            name = _classify_symbol(value)
            if name:
                hits[name] = hits.get(name, 0) + 1
        return max(hits, key=hits.get) if hits else None

    col = guesses.get("index")
    if col and col in df.columns:
        found = from_values(df[col])
        if found:
            return found, f"the '{col}' column"

    if filename:
        found = _classify_symbol(filename)
        if found:
            return found, "the file name"

    for col in df.columns:
        if df[col].dtype == object:
            found = from_values(df[col])
            if found:
                return found, f"text in the '{col}' column"
        found = _classify_symbol(col)
        if found:
            return found, f"the '{col}' header"

    # last resort: a strike-like column, matched against each index's band
    bands = {"NIFTY": (20000, 30000), "BANKNIFTY": (45000, 65000),
             "SENSEX": (70000, 90000)}
    for col in ([strike_col] if strike_col else list(df.columns)):
        if col not in df.columns:
            continue
        values = pd.to_numeric(df[col], errors="coerce").dropna()
        if values.empty:
            continue
        median = float(values.median())
        for name, (lo, hi) in bands.items():
            if lo <= median <= hi:
                return name, f"strike magnitude in '{col}' (~{median:,.0f})"
    return None, None


def _minute_label(value):
    """'09:15' from a time, datetime, Excel fraction or free text. None when
    the cell carries no usable clock time."""
    import pandas as pd

    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel stores a bare time as a fraction of a day
        if 0 <= float(value) < 1:
            total = int(round(float(value) * 24 * 60))
            return f"{total // 60:02d}:{total % 60:02d}"
        return None
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return f"{value.hour:02d}:{value.minute:02d}"
    text = str(value).strip()
    if not text:
        return None
    stamp = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.notna(stamp):
        return f"{stamp.hour:02d}:{stamp.minute:02d}"
    return None


def premium_series(df, time_col, value_col, date_col=None, index_col=None,
                   index_name=None):
    """([[hh:mm, value]], meta) — the premium as one-minute closes.

    Rows are bucketed to the minute and the LAST value in each minute is kept,
    matching how the chart treats every other series (close, not average).
    `index_col` + `index_name` filter a multi-index file down to one index.
    `meta` reports the row counts, the date(s) seen and anything dropped, so
    the dashboard can show what was actually read rather than assert success.
    """
    import pandas as pd

    meta = {"rows": len(df), "used": 0, "dropped_time": 0, "dropped_value": 0,
            "dates": [], "filtered_out": 0}
    work = df

    if index_col and index_name:
        keep = work[index_col].astype(str).str.upper().str.contains(
            str(index_name).upper(), na=False)
        meta["filtered_out"] = int((~keep).sum())
        work = work[keep]

    if date_col and date_col in work.columns:
        stamps = pd.to_datetime(work[date_col], errors="coerce", dayfirst=True)
        meta["dates"] = sorted({d.strftime("%d-%m-%Y")
                                for d in stamps.dropna().dt.normalize().unique()
                                .astype("datetime64[ns]").tolist()
                                for d in [pd.Timestamp(d)]})

    times = work[time_col].map(_minute_label)
    values = pd.to_numeric(work[value_col], errors="coerce")
    meta["dropped_time"] = int(times.isna().sum())
    meta["dropped_value"] = int(values.isna().sum() - (times.isna() & values.isna()).sum())

    ok = times.notna() & values.notna()
    buckets = {}
    for label, value in zip(times[ok], values[ok]):
        buckets[label] = float(value)          # later row wins = the close
    meta["used"] = len(buckets)
    series = sorted(buckets.items(), key=lambda kv: kv[0])
    return [[t, round(v, 2)] for t, v in series], meta


def mids_from(levels, manual=None):
    """{index: day mid} for the chain's ATM, fetched values first and any
    manual entry filling the gaps. A zero or blank manual value counts as
    "not supplied" — the same convention the dashboard inputs use."""
    mids = {}
    for name, entry in (levels or {}).items():
        mid = entry.get("mid")
        if mid:
            mids[name] = mid
    for name, value in (manual or {}).items():
        if name not in mids and value:
            mids[name] = float(value)
    return mids
