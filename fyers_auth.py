"""Refresh the Fyers access token in fyers_credentials.json.

The access token expires DAILY, so this has to be run each morning before the
DOR is generated. Everything else in fyers_credentials.json is permanent.

    python fyers_auth.py            # refresh today's token
    python fyers_auth.py --check    # is the link healthy right now?
    python fyers_auth.py --check 2026-08-17    # ...and pull that day's volume

It prints a login URL, you log in to Fyers in a browser, and you paste the URL
you land on back in. That redirect URL carries the one-time auth_code, which is
exchanged here for the access token and written back to the file.

There is no way to skip the browser step: Fyers requires an interactive login
with the TOTP/PIN, which is the point of it.
"""

import datetime
import json
import logging
import sys
from urllib.parse import parse_qs, urlparse

from fyers_apiv3 import fyersModel

from marketdata import FYERS_CREDENTIALS_FILE

logger = logging.getLogger(__name__)

REQUIRED = ("client_id", "secret_key", "redirect_uri")


def _session(creds):
    return fyersModel.SessionModel(
        client_id=creds["client_id"],
        secret_key=creds["secret_key"],
        redirect_uri=creds["redirect_uri"],
        response_type="code",
        grant_type="authorization_code",
        state="dor",          # echoed back in the redirect; without it the SDK
    )                         # sends the literal string "None"


def login_url(creds):
    """The Fyers URL the operator has to open and log in to.

    Raises ValueError naming what is missing rather than sending a half-built
    request that fails at Fyers with an opaque message."""
    missing = [k for k in REQUIRED if not str(creds.get(k, "")).strip()]
    if missing:
        raise ValueError(
            f"missing {', '.join(missing)} — client_id and secret_key come "
            "from your Fyers API app; redirect_uri must be the exact URL "
            "registered on that app")
    return _session(creds).generate_authcode()


def auth_code_from(landed):
    """The one-time auth_code out of the URL the browser landed on.

    Accepts the whole redirect URL or a bare code, because both are things a
    person plausibly pastes."""
    landed = str(landed or "").strip()
    codes = parse_qs(urlparse(landed).query).get("auth_code")
    if codes:
        return codes[0]
    if landed and "://" not in landed:
        return landed                      # they pasted just the code
    raise ValueError("no auth_code in that URL — paste the whole address you "
                     "landed on, including ?auth_code=…")


def exchange(creds, landed):
    """Trade the one-time auth_code for the daily access token.

    The auth_code is NOT the access token: it is a single-use voucher, valid
    for minutes, and pasting it straight into access_token gives a JWT whose
    `sub` is "auth_code" and which every API call rejects."""
    session = _session(creds)
    session.set_token(auth_code_from(landed))
    response = session.generate_token() or {}
    token = response.get("access_token")
    if not token:
        raise ValueError(f"Fyers refused the auth_code: "
                         f"{response.get('message') or response}")
    return token


def refresh(path=FYERS_CREDENTIALS_FILE):
    """Walk the login at a terminal, write the new token back, return it."""
    with open(path, encoding="utf-8") as fh:
        creds = json.load(fh)
    try:
        url = login_url(creds)
    except ValueError as exc:
        raise SystemExit(f"{path}: {exc}")

    print("\n1. Open this URL and log in to Fyers:\n")
    print("   " + url)
    print("\n2. You will be redirected to a page that may not load — that is "
          "fine.\n   Copy the FULL address bar URL and paste it here.\n")
    try:
        token = exchange(creds, input("   Redirect URL: "))
    except ValueError as exc:
        raise SystemExit(str(exc))

    creds["access_token"] = token
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(creds, fh, indent=2)
        fh.write("\n")
    print(f"\n   Wrote a {len(token)}-character access_token to {path}.")
    print("   It expires at 06:00 tomorrow — run this again then.\n")
    return token


def check(day=None, path=FYERS_CREDENTIALS_FILE):
    """Is the Fyers link healthy right now? Prints, returns True/False.

    Run this BEFORE generating a report. A dead token otherwise surfaces as an
    empty volume panel with a one-line caption, which is easy to miss."""
    import marketdata as md

    creds = json.load(open(path, encoding="utf-8"))
    token = str(creds.get("access_token", ""))
    expiry = md.token_expiry(token)
    print(f"  client_id    {creds.get('client_id') or 'MISSING'}")
    redirect = creds.get("redirect_uri") or ("not set — needed only to "
                                             "refresh the token here")
    print(f"  redirect_uri {redirect}")
    if not token:
        print("  access_token MISSING — run `python fyers_auth.py`")
        return False
    if expiry is None:
        print(f"  access_token {len(token)} chars, but the expiry claim could "
              "not be read — this does not look like a Fyers JWT")
    elif expiry <= datetime.datetime.now():
        print(f"  access_token EXPIRED at {expiry:%d-%m-%Y %H:%M} — "
              "run `python fyers_auth.py`")
        return False
    else:
        left = (expiry - datetime.datetime.now()).total_seconds() / 3600
        print(f"  access_token valid until {expiry:%d-%m-%Y %H:%M} "
              f"({left:.1f}h left)")

    ok = True
    try:
        from fyers_apiv3 import fyersModel
        who = fyersModel.FyersModel(client_id=creds["client_id"],
                                    token=token, is_async=False).get_profile()
        if who.get("s") == "ok":
            print(f"  profile      {who['data'].get('name')} "
                  f"({who['data'].get('fy_id')})")
        else:
            print(f"  profile      REFUSED — {who.get('message')}")
            ok = False
    except Exception as exc:                      # network, SDK, anything
        print(f"  profile      call failed — {exc}")
        ok = False

    if day:
        data, reason = md.fetch_market_volume(
            day, list(md.FYERS_INDEX_SYMBOLS.values()))
        if reason:
            print(f"  volume {day}  FAILED — {reason}")
            ok = False
        for symbol, points in data.items():
            print(f"  volume {day}  {symbol:<20} {len(points):>4} min, "
                  f"total {sum(v for _, v in points):>15,.0f}")
    return ok


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    args = sys.argv[1:]
    if "--check" in args:
        args.remove("--check")
        raise SystemExit(0 if check(args[0] if args else None) else 1)
    refresh()
    check(args[0] if args else None)
