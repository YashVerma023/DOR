"""Download today's ATM premium from Grafana, as the CSVs the report expects.

    python premium_fetcher.py                      # today, every configured index
    python premium_fetcher.py --index SENSEX
    python premium_fetcher.py --date 2026-09-22
    python premium_fetcher.py --out D:\\dor\\inputs

Writes one file per index, 09:15:00 to 15:40:00, in exactly the format the
dashboard's uploader already accepts:

    Synthetic Premium NF 22SEP26.csv
    sep=,
    "time","Premium"
    2026-09-22 09:15:00,210
    2026-09-22 09:15:05,210

Byte-identical in shape to the files uploaded by hand, deliberately — the DOR
needs no change, the upload path stays as the fallback, and an old file can be
diffed against a new one to check this is producing the same thing.

Reads grafana_credentials.json: one entry per index, each possibly on its own
Grafana instance. Panel 11 ('Synthetic Premium') runs `ts.range` on the Redis
key `premium`, which is already call + put combined — panel 5 on the same
dashboard shows `cpremium` and `ppremium` as the separate legs.

Also works for a past date — confirmed against the live instances going back
at least a few weeks (RedisTimeSeries here is not evicting on a short
window). Still, if a given day comes back with 0 rows, it may genuinely be a
non-trading day, or it may be a retention edge somewhere further back — check
before assuming the report has to fall back to a manual upload.
"""

import argparse
import csv
import io
import pathlib
import sys
from datetime import date, datetime, timedelta

from grafana_probe import post_query, resolve

# the abbreviations the existing filenames use
ABBR = {"NIFTY": "NF", "SENSEX": "SX", "BANKNIFTY": "BNF"}

SESSION_START = (9, 15)
SESSION_END = (15, 40)

# the refId the premium target carries on panel 11; the panel also returns
# `ILAST` (yesterday's close) and an empty command frame, which are not wanted
PREMIUM_REF = "premium"


def fetch_series(session, url, targets, datasource, start, end, bucket_ms=5000):
    """[[datetime, value], ...] for the premium target over [start, end]."""
    # ask from one bucket earlier: ts.range excludes the opening sample at the
    # exact boundary, which lost 09:15:00 and left the file one row short of
    # the hand-made ones. Trimmed back to `start` after the fetch.
    body = {"from": str(int((start.timestamp() - bucket_ms / 1000) * 1000)),
            "to": str(int(end.timestamp() * 1000)),
            "queries": []}
    for i, t in enumerate(targets):
        q = dict(t)
        q.setdefault("refId", chr(65 + i))
        q.setdefault("datasource", datasource)
        q["intervalMs"] = bucket_ms
        q.setdefault("maxDataPoints", 20000)
        body["queries"].append(q)

    r = post_query(session, url, body, timeout=180)
    if r.status_code != 200:
        raise SystemExit(f"  /api/ds/query -> HTTP {r.status_code}: {r.text[:300]}")

    best = []
    for ref, res in (r.json().get("results") or {}).items():
        for frame in (res.get("frames") or []):
            fields = (frame.get("schema") or {}).get("fields") or []
            values = (frame.get("data") or {}).get("values") or []
            if len(fields) < 2 or len(values) < 2:
                continue
            # a time field plus a number field is the shape we want; the panel's
            # other targets return hashes and empty frames
            kinds = [f.get("type") for f in fields]
            if "time" not in kinds:
                continue
            ti = kinds.index("time")
            vi = next((j for j, k in enumerate(kinds)
                       if k == "number" and j != ti), None)
            if vi is None:
                continue
            rows = [[datetime.fromtimestamp(ms / 1000), v]
                    for ms, v in zip(values[ti], values[vi])
                    if v is not None
                    and start <= datetime.fromtimestamp(ms / 1000) <= end]
            # prefer the frame whose refId names the premium, else the longest —
            # picking by length alone would happily return the spot series
            if ref.lower().startswith(PREMIUM_REF) and rows:
                return rows
            if len(rows) > len(best):
                best = rows
    return best


def _hhmm(text, fallback):
    """'10:15' / '10:15:00' -> (h, m). Falls back to the session boundary."""
    if not text:
        return fallback
    parts = str(text).split(":")
    try:
        return int(parts[0]), int(parts[1]) if len(parts) > 1 else 0
    except (ValueError, IndexError):
        raise SystemExit(f"  --from/--to want HH:MM, not {text!r}")


def window(day, start_at=None, end_at=None):
    """The [from, to] the query covers — the full session unless narrowed.

    Narrowing is for TESTING against a live day: a short, fixed window gives a
    repeatable comparison, where 09:15–15:40 mid-session returns however much
    has accumulated by the moment the command ran."""
    midnight = datetime.combine(day, datetime.min.time())
    sh, sm = _hhmm(start_at, SESSION_START)
    eh, em = _hhmm(end_at, SESSION_END)
    return (midnight + timedelta(hours=sh, minutes=sm),
            midnight + timedelta(hours=eh, minutes=em))


def fetch_resolved(inst, day, start_at=None, end_at=None):
    """[[datetime, value], …] from an instance `resolve()` has already opened."""
    start, end = window(day, start_at, end_at)
    return fetch_series(inst["session"], inst["url"], inst["targets"],
                        inst["datasource"], start, end)


def fetch(index, day, config=None, start_at=None, end_at=None):
    """[[datetime, value], ...] for one index on one day.

    Finds the instance RUNNING that index today rather than the one a config
    file says owns it — the hosts rotate, and a stale mapping would return the
    other index's premium under this index's filename."""
    index = str(index).upper()
    found, problems = resolve(config, want=[index])
    inst = found.get(index)
    if inst is None:
        detail = "; ".join(f"{k}: {v}" for k, v in problems.items())
        raise SystemExit(
            f"no Grafana instance is running {index} right now"
            + (f" ({detail})" if detail else "")
            + ". Instances rotate between indexes — check the dashboards, or "
              "pin one with \"index\" in grafana_credentials.json.")
    return fetch_resolved(inst, day, start_at, end_at)


def filename(index, day):
    return (f"Synthetic Premium {ABBR.get(index, index)} "
            f"{day.strftime('%d%b%y').upper()}.csv")


def csv_text(rows):
    """The uploader's exact format, down to the `sep=,` line and CRLF endings.

    Text rather than a file, so the report can hand these bytes straight to the
    same reader that handles an upload — one format, one code path, and an old
    hand-made file still diffs clean against a new one."""
    buf = io.StringIO()
    buf.write("sep=,\r\n")
    # the header is quoted in the hand-made exports; matched so that a diff
    # between an old file and a new one shows only real differences
    buf.write('"time","Premium"\r\n')
    w = csv.writer(buf, quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
    for when, value in rows:
        # integers stay integers: the hand-made files carry "210", not
        # "210.0", and a diff against one should be empty
        text = (str(int(value)) if float(value).is_integer()
                else f"{float(value):g}")
        w.writerow([when.strftime("%Y-%m-%d %H:%M:%S"), text])
    return buf.getvalue()


def write_csv(rows, index, day, out_dir):
    path = pathlib.Path(out_dir) / filename(index, day)
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        fh.write(csv_text(rows))
    return path


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--date", default=date.today().isoformat(),
                    help="trading day, YYYY-MM-DD (default: today)")
    ap.add_argument("--index", default=None, help="only this index")
    ap.add_argument("--out", default=".", help="output folder (default: here)")
    ap.add_argument("--from", dest="start_at", default=None, metavar="HH:MM",
                    help="window start (default: 09:15) — for testing against "
                         "a live day, where a short fixed window is repeatable")
    ap.add_argument("--to", dest="end_at", default=None, metavar="HH:MM",
                    help="window end (default: 15:40)")
    args = ap.parse_args(argv)

    day = datetime.strptime(args.date, "%Y-%m-%d").date()

    # one pass over the instances, which also decides which index each is on
    want = [args.index.upper()] if args.index else None
    found, problems = resolve(want=want)
    for label, why in problems.items():
        print(f"  {str(label):<10} {why}")
    for index, inst in sorted(found.items()):
        print(f"  {index:<10} on {inst['url']}"
              + (f"  (index {inst['level']:,.0f})" if inst.get("level") else
                 "  (pinned in config)"))
    if want and not found:
        print(f"  {want[0]:<10} no instance is running it right now")

    failures = len(problems)
    for index, inst in sorted(found.items()):
        try:
            rows = fetch_resolved(inst, day, args.start_at, args.end_at)
        except SystemExit as exc:
            print(f"  {index:<10} {exc}")
            failures += 1
            continue
        except Exception as exc:
            print(f"  {index:<10} {type(exc).__name__}: {exc}")
            failures += 1
            continue

        if not rows:
            print(f"  {index:<10} no premium data for {day:%d-%b-%Y} "
                  f"— check it was a trading day, or that this is the "
                  f"instance that ran {index} back then (they rotate)")
            failures += 1
            continue
        path = write_csv(rows, index, day, args.out)
        print(f"  {index:<10} {len(rows):>6,} rows  "
              f"{rows[0][0]:%H:%M:%S}–{rows[-1][0]:%H:%M:%S}  ->  {path.name}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
