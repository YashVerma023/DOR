"""Discover what a Grafana panel actually queries, so a fetcher can be written.

    python grafana_probe.py                    # every configured index
    python grafana_probe.py --index SENSEX     # just one
    python grafana_probe.py --index NIFTY --panel 3 --date 2026-09-16
    python grafana_probe.py --list             # all dashboards on each host

Nothing here is written back to Grafana and nothing is integrated into the DOR.
This exists only to answer three questions that decide how the real fetcher is
built:

  1. does username/password authenticate against this instance's API at all?
  2. which datasource backs the premium panel, and what query does it run?
  3. does /api/ds/query return that data for an arbitrary date range?

Credentials live in grafana_credentials.json beside this script. It lists the
Grafana INSTANCES, not a mapping of index to host. Paste each dashboard URL
exactly as the browser shows it; the host, uid and slug are all read out of it:

    {
      "defaults": {"user": "admin", "password": "…", "verify_ssl": true,
                   "panel": 11},
      "instances": [
        {"url": "http://host-a:3000/d/abc123/premium-tv?orgId=1"},
        {"url": "http://host-b:3000/d/def456/premium-tv-new?orgId=1"}
      ]
    }

NO INDEX IS WRITTEN DOWN, deliberately. The instances SWAP which index they
run — on 22-09 NIFTY was on .119 and SENSEX on .11; on 23-09 they had traded
places. A hand-written mapping is therefore correct only until the next
rotation, and wrong silently: the fetch succeeds, the file is named NF, and it
holds SENSEX premium.

So each instance is ASKED which index it is running, by reading its own `ILAST`
hash — the same "Yesterday" figure the dashboard shows. NIFTY near 23,000 and
SENSEX near 75,000 are not confusable. An instance can still be pinned with
`"index": "NIFTY"` when that detection needs overriding.

This means detection reflects RIGHT NOW. That suits the fetcher, which is
same-day only anyway because Redis retention will not reach a past date.

`panel` is what this probe is for: run it, read the panel ids it prints, put
the right one in `defaults` (or per instance).

That file must be gitignored. The password is never printed, and only ever
sent to the host named in `url`.
"""

import argparse
import json
import pathlib
import re
import sys
from datetime import date, datetime, timedelta
from urllib.parse import urlparse

import requests

_HERE = pathlib.Path(__file__).resolve().parent
CREDENTIALS = _HERE / "grafana_credentials.json"

# Grafana dashboard URLs look like
#   https://host/d/<uid>/<slug>?orgId=1&viewPanel=3
_UID_IN_URL = re.compile(r"/d/([^/?#]+)")


def load_credentials(path=CREDENTIALS):
    """[{url, uid, panel, index, user, password, verify_ssl}, …] — the instances.

    A LIST, not a map keyed by index: which index an instance runs changes day
    to day, so that mapping cannot live in a file. `index` is None unless it
    was pinned; `resolve()` fills it in by asking the instance.

    The dashboard URL is accepted exactly as copied from the browser — the base
    host, the uid and the slug are all read out of it, because that is the
    thing a person actually has to hand."""
    if not path.exists():
        raise SystemExit(
            f"{path} not found. Create it with an `instances` block — see the\n"
            "docstring at the top of this file. Add it to .gitignore first.")
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise SystemExit(f"{path} is not valid JSON ({exc})")

    defaults = raw.get("defaults") or {}
    entries = raw.get("instances")
    if entries is None:
        # the older shape, keyed by index. Still read, so an old file works,
        # but the index it names is now only a PIN — see the module docstring.
        entries = [dict(cfg or {}, index=name)
                   for name, cfg in (raw.get("indexes") or {}).items()]

    out = []
    for cfg in entries:
        url = str((cfg or {}).get("url") or "").strip()
        if not url:
            continue                        # not configured yet — skip quietly
        parsed = urlparse(url)
        if not (parsed.scheme and parsed.netloc):
            raise SystemExit(f"  {url!r} is not a URL")
        found = _UID_IN_URL.search(parsed.path)
        entry = {
            "url": f"{parsed.scheme}://{parsed.netloc}",
            "uid": found.group(1) if found else None,
            "panel": cfg.get("panel") or defaults.get("panel") or None,
            "index": (str(cfg["index"]).upper() if cfg.get("index") else None),
            "user": cfg.get("user") or defaults.get("user"),
            "password": cfg.get("password") or defaults.get("password"),
            "verify_ssl": cfg.get("verify_ssl",
                                  defaults.get("verify_ssl", True)),
        }
        missing = [k for k in ("user", "password") if not entry[k]]
        if missing:
            raise SystemExit(f"  {entry['url']}: missing {', '.join(missing)}")
        out.append(entry)
    if not out:
        raise SystemExit(
            f"  {path} has no instance with a url filled in.\n"
            "  Paste each dashboard URL into instances[].url and re-run.")
    return out


# The index levels are orders of magnitude apart and have been for decades, so
# the live level names the instrument on its own. Bands are deliberately far
# wider than any plausible year's range, and they do not overlap.
# ponytail: a band check, not a symbol lookup — pin "index" in the config if an
# instance ever carries something these do not cover.
INDEX_BANDS = {
    "NIFTY": (10_000, 40_000),
    "BANKNIFTY": (40_000, 65_000),
    "SENSEX": (65_000, 200_000),
}


def instance_index(session, url, datasource):
    """(index, level) — which index this instance is running RIGHT NOW.

    Read from its own `ILAST` hash, the "Yesterday" figure the dashboard
    already shows, so nothing has to be kept in step by hand. Returns
    (None, level) when the level falls in no band, because guessing wrong here
    mislabels a whole day's premium file."""
    body = {"queries": [{"refId": "A", "datasource": datasource,
                         "command": "hgetall", "keyName": "ILAST",
                         "type": "command"}]}
    r = session.post(f"{url}/api/ds/query", json=body, timeout=60)
    if r.status_code != 200:
        raise SystemExit(f"  {url}: ILAST -> HTTP {r.status_code}")

    level = None
    for res in (r.json().get("results") or {}).values():
        for frame in (res.get("frames") or []):
            names = [f.get("name") for f in
                     ((frame.get("schema") or {}).get("fields") or [])]
            values = (frame.get("data") or {}).get("values") or []
            if "INDEX" in names:
                col = values[names.index("INDEX")]
                if col:
                    level = float(col[0])
    if level is None:
        raise SystemExit(f"  {url}: ILAST carries no INDEX field — cannot tell "
                         "which index this instance is running")
    for name, (low, high) in INDEX_BANDS.items():
        if low <= level < high:
            return name, level
    return None, level


def resolve(config=None, want=None):
    """{index: {**cfg, session, targets, datasource, level}} for what is up.

    One pass: connect, read the panel's query, ask the instance which index it
    is on. Everything the fetch needs is carried back, so nothing is fetched
    twice. An unreachable instance is skipped with its reason recorded rather
    than killing the run — one Grafana being down should not cost the other."""
    resolved, problems = {}, {}
    # `is None`, not truthiness: an empty list means "no instances", and must
    # not quietly reload the real config file and go to the network
    for cfg in (load_credentials() if config is None else config):
        label = cfg["index"] or cfg["url"]
        try:
            session, _ = session_for(cfg)
            if not cfg.get("panel"):
                raise SystemExit(f"  {cfg['url']}: no panel set")
            targets, datasource = _panel_targets(session, cfg["url"],
                                                 cfg["uid"], cfg["panel"])
            index, level = cfg["index"], None
            if index is None:
                index, level = instance_index(session, cfg["url"], datasource)
            if index is None:
                raise SystemExit(
                    f"  {cfg['url']}: index level {level:,.0f} matches no known "
                    "index — pin it with \"index\": \"…\" in the config")
        except SystemExit as exc:
            problems[label] = str(exc).strip()
            continue
        except Exception as exc:
            problems[label] = f"{type(exc).__name__}: {exc}"
            continue
        if want and index not in want:
            continue
        resolved[index] = dict(cfg, session=session, targets=targets,
                               datasource=datasource, level=level, index=index)
    return resolved, problems


def _panel_targets(session, url, uid, panel_id):
    """A panel's own query targets, straight from the dashboard definition."""
    r = session.get(f"{url}/api/dashboards/uid/{uid}", timeout=30)
    r.raise_for_status()
    for p in (r.json()["dashboard"].get("panels") or []):
        for panel in ([p] + (p.get("panels") or [])):
            if panel.get("id") == panel_id:
                return panel.get("targets") or [], panel.get("datasource")
    raise SystemExit(f"  panel {panel_id} not found on dashboard {uid}")


def session_for(creds):
    """A logged-in session.

    Tries BASIC AUTH first — it is stateless and works for Grafana-local users.
    If the instance rejects it (common when logins are handled by SSO, where the
    API still accepts a form login), falls back to POST /login, which sets a
    session cookie. Whichever succeeds is reported, because it decides how the
    real fetcher must authenticate."""
    s = requests.Session()
    s.verify = creds.get("verify_ssl", True)
    s.headers["Accept"] = "application/json"

    s.auth = (creds["user"], creds["password"])
    try:
        who = s.get(f"{creds['url']}/api/user", timeout=30)
    except requests.exceptions.RequestException as exc:
        # unreachable is a different problem from unauthorised, and the fix for
        # one is never the fix for the other
        raise SystemExit(
            f"  cannot reach {creds['url']} ({type(exc).__name__}).\n"
            "  Check the host is right and that this machine can see it — the\n"
            "  Fyers outage on 22-09 turned out to be the office LAN, and a\n"
            "  self-hosted Grafana is even more likely to be network-scoped.\n"
            "  If it uses a self-signed certificate, set \"verify_ssl\": false.")
    if who.status_code == 200 and "json" in who.headers.get("content-type", ""):
        print(f"  auth: basic auth OK — {who.json().get('login')}")
        return s, "basic"

    s.auth = None
    login = s.post(f"{creds['url']}/login",
                   json={"user": creds["user"], "password": creds["password"]},
                   timeout=30)
    if login.status_code == 200:
        who = s.get(f"{creds['url']}/api/user", timeout=30)
        if who.status_code == 200 and "json" in who.headers.get("content-type", ""):
            print(f"  auth: session cookie OK — {who.json().get('login')}")
            return s, "cookie"

    raise SystemExit(
        f"  auth FAILED: basic auth returned {who.status_code}, form login "
        f"returned {login.status_code}.\n"
        "  If this instance uses SSO, a service account token is the way in:\n"
        "  Administration -> Service accounts -> Add service account token.")


def list_dashboards(s, url, limit=40):
    r = s.get(f"{url}/api/search", params={"type": "dash-db", "limit": limit},
              timeout=30)
    r.raise_for_status()
    rows = r.json()
    print(f"\n  {len(rows)} dashboard(s):")
    for d in rows:
        print(f"    {d.get('uid'):<24} {d.get('title')}")
    print("\n  Re-run with --dash <uid> to see a dashboard's panels, then put the")
    print("  uid + panel id into grafana_credentials.json:")
    print('    "dashboards": { "NIFTY": {"uid": "…", "panel": 3}, … }')


_INDEX_HINT = re.compile(r"\b(BANKNIFTY|BANK\s*NIFTY|NIFTY|SENSEX|NSE|BSE|BFO|NFO)\b", re.I)


def scan(s, url, term, limit=40):
    """Every dashboard whose title matches `term`, with the INDEX each panel
    actually queries.

    Titles lie, or at least drift — "Premium", "Premium TV" and "VivekPremium"
    cannot be told apart by name. The query text names the instrument, so that
    is what decides which dashboard belongs to which index."""
    r = s.get(f"{url}/api/search", params={"type": "dash-db", "limit": limit},
              timeout=30)
    r.raise_for_status()
    hits = [d for d in r.json() if term.lower() in str(d.get("title", "")).lower()]
    print(f"\n  {len(hits)} dashboard(s) matching {term!r}:\n")
    for d in hits:
        uid = d["uid"]
        got = s.get(f"{url}/api/dashboards/uid/{uid}", timeout=30)
        if got.status_code != 200:
            print(f"    {uid:<20} {d['title']!r}  -> HTTP {got.status_code}")
            continue
        dash = got.json()["dashboard"]
        print(f"    {uid:<20} {d['title']!r}")
        for p in (dash.get("panels") or []):
            for panel in ([p] + (p.get("panels") or [])):
                if panel.get("type") == "row":
                    continue
                blob = json.dumps(panel.get("targets") or [])
                found = sorted({m.group(1).upper().replace(" ", "")
                                for m in _INDEX_HINT.finditer(blob)})
                ds = panel.get("datasource")
                ds_type = ds.get("type") if isinstance(ds, dict) else ds
                print(f"       panel {str(panel.get('id')):<4} "
                      f"{str(panel.get('title'))[:34]:<34} "
                      f"[{ds_type}]  index hints: {', '.join(found) or '—'}")
    print("\n  Then: --dash <uid> to see the full query, "
          "--dash <uid> --panel <id> --date <day> to run it.")


def show_dashboard(s, url, uid):
    """Panels, their datasources and the raw queries behind them."""
    r = s.get(f"{url}/api/dashboards/uid/{uid}", timeout=30)
    if r.status_code != 200:
        raise SystemExit(f"  dashboard {uid}: HTTP {r.status_code} {r.text[:200]}")
    dash = r.json()["dashboard"]
    print(f"\n  dashboard: {dash.get('title')}  (uid {uid})")
    panels = dash.get("panels") or []
    for p in panels:
        # rows nest their children
        for panel in ([p] + (p.get("panels") or [])):
            if panel.get("type") == "row":
                continue
            ds = panel.get("datasource")
            print(f"\n    panel {panel.get('id')}: {panel.get('title')!r} "
                  f"[{panel.get('type')}]")
            print(f"      datasource: {json.dumps(ds)}")
            for t in (panel.get("targets") or []):
                # the query shape differs per datasource; print it whole rather
                # than guessing which key holds it
                keys = {k: v for k, v in t.items()
                        if k not in ("datasource",) and v not in (None, "", [])}
                print(f"      target: {json.dumps(keys)[:600]}")
    print("\n  Re-run with --panel <id> --date YYYY-MM-DD to run that query.")
    return dash


def run_panel_query(s, url, dash, panel_id, day):
    """Execute one panel's query for a single trading day via /api/ds/query."""
    panel = None
    for p in (dash.get("panels") or []):
        for cand in ([p] + (p.get("panels") or [])):
            if cand.get("id") == panel_id:
                panel = cand
    if panel is None:
        raise SystemExit(f"  panel {panel_id} not found on this dashboard")

    start = datetime.combine(day, datetime.min.time()) + timedelta(hours=9, minutes=15)
    end = datetime.combine(day, datetime.min.time()) + timedelta(hours=15, minutes=40)
    body = {
        "from": str(int(start.timestamp() * 1000)),
        "to": str(int(end.timestamp() * 1000)),
        "queries": [],
    }
    for i, t in enumerate(panel.get("targets") or []):
        q = dict(t)
        q.setdefault("refId", chr(65 + i))
        q.setdefault("datasource", panel.get("datasource"))
        q.setdefault("intervalMs", 60000)          # 1-minute buckets
        q.setdefault("maxDataPoints", 2000)
        body["queries"].append(q)
    if not body["queries"]:
        raise SystemExit("  that panel has no targets to run")

    print(f"\n  POST /api/ds/query  {start:%Y-%m-%d %H:%M} .. {end:%H:%M}")
    r = s.post(f"{url}/api/ds/query", json=body, timeout=120)
    print(f"  HTTP {r.status_code}")
    if r.status_code != 200:
        print(f"  {r.text[:800]}")
        return
    data = r.json()
    for ref, res in (data.get("results") or {}).items():
        frames = res.get("frames") or []
        print(f"\n  refId {ref}: {len(frames)} frame(s)")
        for f in frames[:3]:
            schema = f.get("schema") or {}
            fields = schema.get("fields") or []
            values = (f.get("data") or {}).get("values") or []
            print(f"    name={schema.get('name')!r} "
                  f"fields={[fld.get('name') for fld in fields]}")
            print(f"    rows={len(values[0]) if values else 0}")
            for col, vals in zip(fields, values):
                print(f"      {col.get('name')!r} ({col.get('type')}): "
                      f"{vals[:4]}{' …' if len(vals) > 4 else ''}")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--index", default=None,
                    help="only this index (default: every one configured)")
    ap.add_argument("--panel", type=int, help="panel id to execute")
    ap.add_argument("--date", default=date.today().isoformat(),
                    help="trading day, YYYY-MM-DD (default: today)")
    ap.add_argument("--list", action="store_true",
                    help="list every dashboard on each instance")
    args = ap.parse_args(argv)

    for cfg in load_credentials():
        print(f"\n{'=' * 64}\n  {cfg['url']}  (uid {cfg['uid']})")
        try:
            s, mode = session_for(cfg)
        except SystemExit as exc:
            print(str(exc))
            continue
        if args.list or not cfg["uid"]:
            list_dashboards(s, cfg["url"])
            continue
        dash = show_dashboard(s, cfg["url"], cfg["uid"])
        panel = args.panel or cfg["panel"]

        # name the instance from its own data before running anything — with
        # the hosts rotating, "which index is this?" is the first question
        if panel:
            _, datasource = _panel_targets(s, cfg["url"], cfg["uid"], panel)
            index, level = instance_index(s, cfg["url"], datasource)
            print(f"\n  running: {index or 'UNKNOWN'}  (ILAST index {level:,.0f})")
            if args.index and index != args.index.upper():
                print(f"  skipped — not {args.index.upper()}")
                continue
            run_panel_query(s, cfg["url"], dash, panel,
                            datetime.strptime(args.date, "%Y-%m-%d").date())
    return 0


if __name__ == "__main__":
    sys.exit(main())
