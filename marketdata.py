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

import base64
import json
import logging
import os
import time
from datetime import date, datetime, timedelta

logger = logging.getLogger(__name__)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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

    # yfinance is the source. Fyers was tried as primary and agrees to the
    # paisa, but the report must not depend on a token that expires at 06:00
    # daily — a missing token would silently change where the levels came from.
    # _fyers_levels stays in this module for volume_fetcher.py, which is run
    # separately and whose output is uploaded.
    levels, problems = {}, {}

    try:
        import yfinance as yf
    except ImportError:
        reason = ("yfinance is not installed — run `pip install yfinance` "
                  "or enter the levels manually")
        logger.error("market data unavailable: %s", reason)
        problems.update({name: f"{problems.get(name, reason)}" for name in wanted})
        return levels, problems

    if day > date.today():
        reason = f"{day:%d-%m-%Y} is in the future"
        problems.update({name: reason for name in wanted})
        return levels, problems
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


def _candles(resp):
    """Fyers history candles, one per timestamp, in time order.

    Fyers returns the SAME-DAY session TWICE — 750 rows covering 375 distinct
    minutes on the report date, but 375 on any settled earlier date. It looks
    like the live intraday feed appended to the stored one. Left alone this
    doubles every volume SUM and makes the chart draw a return stroke back to
    09:15 (high/low/open/close survive it, which is why it hid for a while).
    Verified 20-08-2026: 750 raw -> 375 unique on both NIFTY and SENSEX.

    The LAST row for a timestamp wins: on a live minute the later copy is the
    more complete one."""
    by_time = {}
    for row in resp.get("candles") or []:
        if len(row) >= 6:
            by_time[row[0]] = row
    return [by_time[k] for k in sorted(by_time)]


def _fyers_levels(day, wanted, interval):
    """({index: levels}, {index: reason}) from Fyers — one call per index.

    The day's OHLC is derived from the 1-minute candles rather than asking for
    the daily bar: the two were verified identical to the paisa, and the daily
    endpoint returns None intermittently, so deriving costs nothing and removes
    a failure mode. Never raises — anything unavailable falls back to Yahoo."""
    creds = load_fyers_credentials()
    if not creds:
        return {}, {name: "no Fyers token" for name in wanted}
    try:
        from fyers_apiv3 import fyersModel
        client = fyersModel.FyersModel(client_id=creds["client_id"],
                                       token=creds["access_token"],
                                       is_async=False)
    except Exception as exc:
        return {}, {name: f"Fyers unavailable ({type(exc).__name__})"
                    for name in wanted}

    # the report speaks Yahoo's "5m"; Fyers wants "5"
    resolution = str(interval or "1").rstrip("m") or "1"
    levels, problems = {}, {}
    for name in wanted:
        symbol = FYERS_INDEX_SYMBOLS.get(name)
        if not symbol:
            problems[name] = f"no Fyers symbol for {name}"
            continue
        resp = _throttled(client.history,
                          {"symbol": symbol, "resolution": resolution,
                           "date_format": "1", "range_from": day.isoformat(),
                           "range_to": day.isoformat(), "cont_flag": "1"})
        candles = _candles(resp)
        if not candles:
            problems[name] = (resp.get("message")
                              or f"Fyers has no {day:%d-%m-%Y} bar for {name}")
            continue
        entry = {
            "symbol": symbol,
            "open": float(candles[0][1]),
            "high": float(max(c[2] for c in candles)),
            "low": float(min(c[3] for c in candles)),
            "close": float(candles[-1][4]),
            "volume": float(sum(c[5] for c in candles)),
        }
        entry["mid"] = day_mid(entry)
        if interval in INTRADAY_INTERVALS:
            entry["series"] = [
                {"t": datetime.fromtimestamp(c[0]).strftime("%H:%M"),
                 "open": float(c[1]), "high": float(c[2]),
                 "low": float(c[3]), "close": float(c[4])}
                for c in candles
            ]
        levels[name] = entry
        logger.info("fyers %s (%s) %s: H=%.2f L=%.2f mid=%.2f", name, symbol,
                    day, entry["high"], entry["low"], entry["mid"])
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


# ---------------------------------------------------------------------------
# Market volume — Fyers History API
# ---------------------------------------------------------------------------

FYERS_CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "fyers_credentials.json")

# Fyers index symbols, for the market-volume fetch.
FYERS_INDEX_SYMBOLS = {
    "NIFTY": "NSE:NIFTY50-INDEX",
    "BANKNIFTY": "NSE:NIFTYBANK-INDEX",
    "SENSEX": "BSE:SENSEX-INDEX",
}


# Credentials supplied by the running app rather than the file. A hosted
# deployment (Streamlit Community Cloud) has an EPHEMERAL filesystem and the
# gitignored fyers_credentials.json is not even in the repo, so the token has
# to live in the app's own state and be pushed in here.
_RUNTIME_CREDENTIALS = {}


def set_fyers_credentials(client_id=None, access_token=None):
    """Use these credentials instead of the JSON file. Call with nothing to
    clear and fall back to the file."""
    global _RUNTIME_CREDENTIALS
    if client_id and access_token:
        _RUNTIME_CREDENTIALS = {"client_id": str(client_id).strip(),
                                "access_token": str(access_token).strip()}
    else:
        _RUNTIME_CREDENTIALS = {}
    return _RUNTIME_CREDENTIALS


def load_fyers_credentials(path=None):
    """{'client_id', 'access_token'}, or {} with the reason logged.

    Order: credentials pushed in by the app first (hosted, ephemeral disk),
    then fyers_credentials.json (local). A token is refused — not passed on to
    fail confusingly at the API — when it is the auth_code rather than the
    access token, or when it has expired."""
    if _RUNTIME_CREDENTIALS:
        creds, source = dict(_RUNTIME_CREDENTIALS), "the supplied token"
    else:
        path = path or FYERS_CREDENTIALS_FILE
        if not os.path.exists(path):
            return {}
        try:
            with open(path, encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, ValueError) as exc:
            logger.error("fyers_credentials.json unreadable (%s)", exc)
            return {}
        creds = {k: str(data.get(k, "")).strip()
                 for k in ("client_id", "access_token")}
        source = os.path.basename(path)
        if not all(creds.values()):
            return {}

    if token_kind(creds["access_token"]) == "auth_code":
        logger.error("%s holds the AUTH CODE, not the access token — the code "
                     "is a single-use voucher that must be exchanged. Run "
                     "`python fyers_auth.py`, which does the exchange for you.",
                     source)
        return {}
    expiry = token_expiry(creds["access_token"])
    if expiry and expiry <= datetime.now():
        logger.error("Fyers access_token expired at %s — run "
                     "`python fyers_auth.py`", expiry.strftime("%d-%m-%Y %H:%M"))
        return {}
    return creds


def token_claims(token):
    """The JWT payload, unverified, or {}. A plain base64 read — NOT a
    signature check, which only Fyers can do."""
    try:
        body = str(token).split(".")[1]
        body += "=" * (-len(body) % 4)
        return json.loads(base64.urlsafe_b64decode(body))
    except (IndexError, ValueError, TypeError):
        return {}


def token_kind(token):
    """"access_token", "auth_code", or None if it cannot be read.

    These are two DIFFERENT JWTs and they look alike. The auth_code is the
    single-use voucher the browser redirect carries; it has to be exchanged
    for the access token using the secret. Pasting the code straight into
    access_token yields a token every API call rejects with "Could not
    authenticate the user", which is a confusing way to learn this."""
    return token_claims(token).get("sub") or None


def token_expiry(access_token):
    """When the Fyers JWT stops working, or None if it cannot be read.

    Fyers tokens expire at 06:00 the following morning, not 24h after issue,
    so a token minted at 21:00 is good for nine hours. Reading the claim is
    the only way to say that honestly — it is a plain unsigned base64 read,
    NOT a signature check, which only Fyers can do."""
    exp = token_claims(access_token).get("exp")
    return None if exp is None else datetime.fromtimestamp(exp)


def fetch_market_volume(day, symbols, interval="1"):
    """({symbol: [[minute, volume], ...]}, reason) from the Fyers History API.

    `symbols` are Fyers symbols — an index for index volume, or option
    contracts to sum a strike range. Candle timestamps mark the START of the
    interval, matching how the chart buckets everything else.

    Verified against a live Fyers account on 18-08-2026. NOTE this is the
    CASH-MARKET volume of the index constituents, not options volume — it is
    ~1000x an options figure and must never be compared with MS volume.
    Returns ({}, reason) for every failure so the chart degrades to MS volume
    alone rather than breaking the report.
    """
    creds = load_fyers_credentials()
    if not creds:
        return {}, ("no Fyers client_id + access_token in "
                    "fyers_credentials.json — run `python fyers_auth.py` "
                    "to issue today's token")
    try:
        from fyers_apiv3 import fyersModel
    except ImportError:
        return {}, "fyers-apiv3 is not installed (`pip install fyers-apiv3`)"

    try:
        day = _as_date(day)
    except ValueError as exc:
        return {}, str(exc)

    try:
        client = fyersModel.FyersModel(client_id=creds["client_id"],
                                       token=creds["access_token"], is_async=False)
    except Exception as exc:                                  # auth / construction
        return {}, f"Fyers client failed to start ({type(exc).__name__}: {exc})"

    out, failures = {}, []
    for symbol in symbols:
        payload = {"symbol": symbol, "resolution": interval, "date_format": "1",
                   "range_from": day.isoformat(), "range_to": day.isoformat(),
                   "cont_flag": "1"}
        try:
            resp = client.history(payload) or {}
        except Exception as exc:
            failures.append(f"{symbol}: {type(exc).__name__}")
            continue
        candles = _candles(resp)
        if not candles:
            failures.append(f"{symbol}: {resp.get('message') or 'no candles'}")
            continue
        # candle = [epoch, open, high, low, close, volume]
        out[symbol] = [
            [datetime.fromtimestamp(c[0]).strftime("%H:%M"), float(c[5])]
            for c in candles
        ]
        logger.info("fyers %s: %d candles, volume %s", symbol, len(candles),
                    f"{sum(c[5] for c in candles):,.0f}")
    reason = "; ".join(failures) if failures and not out else None
    if failures:
        logger.warning("fyers: %d symbol(s) returned nothing — %s",
                       len(failures), "; ".join(failures[:3]))
    return out, reason


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


# ---------------------------------------------------------------------------
# Index option volume — everything traded in one index's option chain
# ---------------------------------------------------------------------------
#
# This is the market side of the volume panel: the exchange's per-minute volume
# summed over EVERY strike of one index's expiry. MS volume is the desk's slice
# of exactly this number, in the same unit (contracts), so MS / Index is a real
# share of that index's options flow.
#
# Symbols are ENUMERATED from Fyers' own option chain rather than constructed.
# Fyers spells weeklies and monthlies differently (NSE:NIFTY2681824100CE vs
# NSE:NIFTY26AUG24100PE) and getting that wrong does not fail loudly — a wrong
# guess can resolve to a real but DIFFERENT contract and report its volume as
# yours. Asking the chain removes the guess.
#
# Two hard limits, both reported rather than papered over:
#   * Fyers delists an expiry once it has passed, so a report built after its
#     own expiry gets no market volume at all.
#   * The data API is rate limited (HTTP 429 "request limit reached") at
#     roughly 10/second and 200/minute, so a 200-strike chain is paced and
#     takes about a minute. Un-paced it silently returns short.

def parse_expiry(text, day=None):
    """The dashboard's free-text expiry ("28JUL26", "13-08-2026") as a date.

    A 2-digit year is read in the report's own century rather than blindly
    20xx, so the map cannot silently land a hundred years out."""
    text = str(text or "").strip().upper().replace("-", "").replace("/", "")
    if not text:
        return None
    for fmt in ("%d%b%y", "%d%b%Y", "%d%m%Y", "%d%m%y", "%Y%m%d"):
        try:
            parsed = datetime.strptime(text, fmt).date()
        except ValueError:
            continue
        if fmt.endswith("%y") and day is not None:
            century = _as_date(day).year // 100 * 100
            parsed = parsed.replace(year=century + parsed.year % 100)
        return parsed
    return None


_FYERS_MIN_INTERVAL = 0.35        # ~170 calls/min, inside the 200/min ceiling
_FYERS_RETRIES = 3


def _throttled(call, *args):
    """One Fyers call, paced and retried past 429. Returns {} if it never lands."""
    for attempt in range(_FYERS_RETRIES):
        resp = call(*args) or {}
        if resp.get("code") != 429:
            return resp
        wait = 2 ** attempt
        logger.warning("fyers rate limit — waiting %ss", wait)
        time.sleep(wait)
    logger.error("fyers still rate limited after %d retries", _FYERS_RETRIES)
    return {}


def _chain_symbols(client, index, day, expiry):
    """(option symbols, expiry date, reason) for one index's expiry.

    The chain is a LIVE endpoint: it lists only expiries that have not passed.
    `expiry` (a date, optional) picks one; without it the nearest is used."""
    underlying = FYERS_INDEX_SYMBOLS.get(index)
    if not underlying:
        return [], None, f"{index} has no Fyers underlying symbol"

    chain = _throttled(client.optionchain,
                       {"symbol": underlying, "strikecount": 50, "timestamp": ""})
    expiries = (chain.get("data") or {}).get("expiryData") or []
    if not expiries:
        return [], None, (f"Fyers returned no expiries for {index} "
                          f"({chain.get('message') or 'empty option chain'})")

    stamp, chosen = "", None
    if expiry:
        for entry in expiries:
            if _as_date(entry["date"].replace("-", "/")) == expiry:
                stamp, chosen = entry["expiry"], expiry
                break
        if chosen is None:
            listed = ", ".join(e["date"] for e in expiries[:4])
            return [], None, (f"{index} {expiry:%d-%b-%Y} is not listed by Fyers "
                              f"— expired contracts are delisted (listed: {listed})")
    else:
        chosen = _as_date(expiries[0]["date"].replace("-", "/"))
        stamp = expiries[0]["expiry"]

    if stamp:
        chain = _throttled(client.optionchain,
                           {"symbol": underlying, "strikecount": 50,
                            "timestamp": stamp})
    symbols = [row["symbol"]
               for row in (chain.get("data") or {}).get("optionsChain") or []
               if row.get("strike_price", -1) != -1]
    if not symbols:
        return [], chosen, f"Fyers returned no strikes for {index} {chosen}"
    return symbols, chosen, None


def fetch_index_option_volume(day, index, expiry=None, interval="1"):
    """([[minute, volume], ...], reason) — the whole option chain of one index.

    Volume is in CONTRACTS, the same unit as tradevalue.volume_timeline, so the
    two can share a scale. One paced API call per strike, so a full chain runs
    about a minute — the caller is expected to cache it."""
    creds = load_fyers_credentials()
    if not creds:
        return [], ("no Fyers client_id + access_token in "
                    "fyers_credentials.json — run `python fyers_auth.py` "
                    "to issue today's token")
    try:
        from fyers_apiv3 import fyersModel
        day = _as_date(day)
        client = fyersModel.FyersModel(client_id=creds["client_id"],
                                       token=creds["access_token"],
                                       is_async=False)
    except ImportError:
        return [], "fyers-apiv3 is not installed (`pip install fyers-apiv3`)"
    except Exception as exc:
        return [], f"Fyers client failed to start ({type(exc).__name__}: {exc})"

    symbols, resolved, reason = _chain_symbols(client, index, day, expiry)
    if reason:
        return [], reason

    per_minute, blank, failed = {}, 0, []
    for symbol in symbols:
        time.sleep(_FYERS_MIN_INTERVAL)
        resp = _throttled(client.history,
                          {"symbol": symbol, "resolution": interval,
                           "date_format": "1", "range_from": day.isoformat(),
                           "range_to": day.isoformat(), "cont_flag": "1"})
        candles = _candles(resp)
        if not candles:
            # a strike that simply did not trade is normal; a REFUSED call is
            # not, and the two must not be conflated or a throttled run looks
            # like a quiet day
            if resp.get("s") == "error" and "Invalid symbol" not in str(resp.get("message")):
                failed.append(f"{symbol}: {resp.get('message')}")
            else:
                blank += 1
            continue
        for row in candles:
            minute = datetime.fromtimestamp(row[0]).strftime("%H:%M")
            per_minute[minute] = per_minute.get(minute, 0.0) + float(row[5])

    total = sum(per_minute.values())
    logger.info("fyers %s %s option chain: %d strikes, %d quiet, %d failed, "
                "volume %s", index, resolved, len(symbols), blank, len(failed),
                f"{total:,.0f}")
    if failed:
        # a partial sum is a WRONG number, not a small one — say so
        return ([[m, v] for m, v in sorted(per_minute.items())],
                f"{len(failed)} of {len(symbols)} {index} strikes failed "
                f"({failed[0]}) — the volume line is incomplete")
    if not per_minute:
        return [], f"no {index} option volume on {day:%d-%b-%Y}"
    return [[m, v] for m, v in sorted(per_minute.items())], None


def chain_expiries(index):
    """([expiry dates], reason) currently LISTED for one index, soonest first.

    Fyers lists only expiries that have not passed. An expiry day is therefore
    a trap: ask on 20-Aug and the 20-Aug weekly is there, ask on 21-Aug and the
    nearest becomes 27-Aug — a different contract carrying a fraction of the
    day's volume. Callers should name the expiry rather than take the nearest
    on trust."""
    creds = load_fyers_credentials()
    if not creds:
        return [], "no Fyers token"
    underlying = FYERS_INDEX_SYMBOLS.get(index)
    if not underlying:
        return [], f"{index} has no Fyers underlying symbol"
    try:
        from fyers_apiv3 import fyersModel
        client = fyersModel.FyersModel(client_id=creds["client_id"],
                                       token=creds["access_token"],
                                       is_async=False)
        chain = _throttled(client.optionchain,
                           {"symbol": underlying, "strikecount": 1,
                            "timestamp": ""})
    except Exception as exc:
        return [], f"{type(exc).__name__}: {exc}"
    rows = (chain.get("data") or {}).get("expiryData") or []
    dates = []
    for row in rows:
        try:
            dates.append(_as_date(str(row["date"]).replace("-", "/")))
        except (KeyError, ValueError):
            continue
    return sorted(dates), None if dates else "Fyers listed no expiries"


# ---------------------------------------------------------------------------
# Volume workbook — the file volume_fetcher.py writes, uploaded to the report
# ---------------------------------------------------------------------------

VOLUME_SHEETS = {"nifty": "NIFTY", "sensex": "SENSEX", "banknifty": "BANKNIFTY"}


def read_volume_sheets(source):
    """({index: [[minute, volume]]}, {index: reason}, {index: expiry note}).

    Written by volume_fetcher.py: one sheet per index, named nifty / sensex /
    banknifty, columns `time` and `volume`. Read defensively — this file is
    produced on a different day by a different run, so a stale or hand-edited
    one must be reported rather than silently plotted:

      * sheet names are matched case-insensitively, and a sheet whose header
        is not time/volume falls back to the first two columns;
      * rows whose time or volume will not parse are counted and dropped, not
        coerced to zero, which would draw a dip that never happened;
      * duplicate minutes are SUMMED, since a chain fetched in parts can
        legitimately report the same minute twice.

    The expiry volume_fetcher.py wrote in D1 is returned so the caller can show
    it. It matters: Fyers delists a passed expiry, so a workbook generated the
    morning AFTER an expiry day silently holds the NEXT contract. Comparing
    that against a book traded on the expired one is apples to oranges — it
    read as a 56% market share in testing.
    """
    import pandas as pd

    frames = pd.read_excel(source, sheet_name=None)
    by_name = {str(k).strip().lower(): v for k, v in frames.items()}
    # D1 holds "expiry 25-Aug-2026"; it lands in the header row, so pandas
    # reads it as a column name rather than a cell
    notes = {}
    for sheet, index in VOLUME_SHEETS.items():
        frame = by_name.get(sheet)
        if frame is None:
            continue
        for col in frame.columns:
            text = str(col).strip()
            if text.lower().startswith("expiry"):
                notes[index] = text
    out, problems = {}, {}
    for sheet, index in VOLUME_SHEETS.items():
        frame = by_name.get(sheet)
        if frame is None:
            continue                       # index simply not in this workbook
        if frame.empty:
            problems[index] = f"sheet '{sheet}' is empty"
            continue
        cols = {str(c).strip().lower(): c for c in frame.columns}
        time_col = cols.get("time", frame.columns[0])
        vol_col = cols.get("volume",
                           frame.columns[1] if len(frame.columns) > 1 else None)
        if vol_col is None:
            problems[index] = f"sheet '{sheet}' has no volume column"
            continue
        per_minute, dropped = {}, 0
        for stamp, value in zip(frame[time_col], frame[vol_col]):
            minute = _minute_label(stamp)
            try:
                volume = float(value)
            except (TypeError, ValueError):
                minute = None
            if not minute or pd.isna(value):
                dropped += 1
                continue
            per_minute[minute] = per_minute.get(minute, 0.0) + volume
        if not per_minute:
            problems[index] = (f"sheet '{sheet}' has {len(frame)} row(s) but "
                               "none carried a readable time and volume")
            continue
        if dropped:
            logger.warning("volume sheet %s: %d unreadable row(s) dropped",
                           sheet, dropped)
        out[index] = [[m, v] for m, v in sorted(per_minute.items())]
        logger.info("volume sheet %s: %d minute(s), total %s (%s)", sheet,
                    len(out[index]), f"{sum(per_minute.values()):,.0f}",
                    notes.get(index, "expiry not stated"))
    return out, problems, notes
