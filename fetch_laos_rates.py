#!/usr/bin/env python3
"""
fetch_laos_rates.py

Fetches live Laos remittance quotes from each competitor's hidden calculator
API and writes the KRW (1) / Service fee (2) values into the INPUT AREA
(columns M:N) of the "Laos INPUT New" sheet in Exchange rate daily report.xlsx.

Unlike Thailand (one corridor, one currency), the Laos sheet is 5 separate
corridor/service blocks spanning 3 currencies and 2 service types:

    Block 1 - LAK, Bank Deposit   (15,000,000 LAK) - rows 8-13
    Block 2 - USD, Bank Deposit   (1,000 USD)       - rows 16-20
    Block 3 - THB, Bank Deposit   (26,000 THB)      - rows 23-27
    Block 4 - LAK, Cash Pickup/WU (15,000,000 LAK)  - rows 30-32
    Block 5 - USD, Cash Pickup/MoneyGram (1,000 USD) - rows 35-40

Blocks 2 and 3 gained a 4th competitor (GmoneyTrans, which added Bank
Deposit for THB/USD - previously it only had LAK Bank Deposit and USD Cash
Pickup) after both blocks had already been built with exactly 3 competitors
and no spare ranking row (unlike Block 5, which already had a pre-built 5th
rank row before GME needed it). Adding a 4th competitor there required
genuinely inserting a new row into each block via Excel COM - see
.claude/skills/laos-rate-report/scripts/insert_gmoneytrans_rows.ps1 (already
run once against the real workbook; not idempotent, don't rerun it). That
script's own comments cover a real gotcha worth knowing if this ever needs
doing again: Excel's row-insert auto-adjusts genuine cell references
elsewhere on the sheet (so Blocks 4/5 shifting down was handled
automatically for their $L$../$Q$.. references), but NOT a bare numeric
literal like the "28" in "ROW()-28" - that's just arithmetic, not a link to
a cell, so Blocks 4/5's rank formulas needed that anchor number fixed by
hand after the insert or they'd have silently ranked wrong (LARGE() asking
for a rank that doesn't exist, blanked out by IFERROR rather than erroring
loudly).

GME - all 6 rows, via the mobile app's API (not the website). GME's own
website calculator only ever covers Bank Deposit and can't tell RIA and
Moneygram-BCEL apart (see below); it also has no Laos-THB entry and its Cash
Payment calculator errors out regardless of amount for Laos LAK/USD - so it
used to only fetch rows 8 and 16, leaving 9/23/30/35 permanently untouched.

That's fixed by calling GME's own mobile app API directly:
POST https://mobileapi.gmeremit.com:8002/api/v1/mobile/sendmoney/calculateV2,
captured live via Charles Proxy against GME's own iOS app (an authorized,
same-company capture, not a third party). Each of the 6 corridors below uses
a placeholder receiver already saved to the capturing account (the receiver
doesn't need to be real - it just gives the calculator a payout route to
quote against) and a fixed pAgent/receiverId/payOutPartner triplet - see
GME_MOBILE_CORRIDORS for the full table:

    ("GME (RIA-Other banks)", "LAK") -> row 8
    ("GME (Moneygram-BCEL)", "LAK")  -> row 9
    ("GME", "USD")                   -> row 16
    ("GME", "THB")                   -> row 23
    ("GME (WU)", "LAK")              -> row 30
    ("GME (Moneygram)", "USD")       -> row 35

Unlike the website UI, the API's payOutPartner field genuinely distinguishes
RIA from Moneygram-BCEL (different partner IDs, different rates), so both
rows 8 and 9 are now written - previously, writing the same tied value to
both broke the sheet's ranking formula (LARGE/MATCH/INDEX can't tell two
identical numbers apart), which is why row 9 was left blank for so long.

Like GME's other rows, fee is always written as 0 (krw holds the endpoint's
full all-in total, Data.collAmt = Data.sAmt + Data.scCharge) - same
deliberate convention used everywhere else GME appears in these scripts.

Auth: the mobile API needs a bearer token + session cookie that expire in
about a week, and there's no captured login endpoint to refresh them
automatically. These live in a local, untracked file next to this script,
gme_mobile_secrets.json:
    {"authorization": "Bearer <jwt>", "cookie": "WMONID=<value>", "username": "<account username>"}
To refresh: open the GME app on a phone proxied through Charles (with
Charles' root cert installed and trusted ON THE PHONE - each Charles
install has its own CA, so switching machines means re-trusting a new one),
trigger any GME quote screen, find the POST to
mobileapi.gmeremit.com:8002/.../calculateV2 in Charles, and copy its
Authorization/Cookie header values into this file. If the file is missing,
or its token has already expired (checked automatically via the JWT's own
exp claim before any request is sent), the mobile fetch is skipped entirely
for this run.

Playwright fallback: fetch_gme_browser_laos (the old website-scraping
approach) is kept as an automatic fallback, but only for the 2 corridors it
was ever able to cover - "GME (RIA-Other banks)"/LAK (row 8) and GME/USD
(row 16) - and only when the mobile API didn't get them this run (missing/
expired secrets, or that specific REST call failing). Rows 9/23/30/35 have
no fallback; if the mobile API doesn't get them, they're simply flagged
stale like any other failed quote.

Several rows across blocks want the exact same underlying quote rather than
a fresh fetch:
  - Hanpass's API silently ignores the Bank-Deposit vs Cash-Pickup flag
    (byte-identical responses either way), so Blocks 4 and 5's Hanpass rows
    reuse Blocks 1 and 2's Hanpass figures instead of querying again.
  - Cross has no separate cash-pickup parameter for Laos (its corridor
    selector only varies by currency, not service type), so Block 5's Cross
    row reuses Block 2's USD figure. This is an assumption - if Cross's
    actual cash-pickup price for Laos differs, this will need revisiting.
  - E9Pay has no cash-pickup variant at all for Laos (only two nation codes
    exist total: LA01=USD, LA02=LAK, both bank-deposit-style), so Block 5's
    E9Pay row is fetched via the same call as if it were Block 2, just
    written into Block 5 too.

Provider notes (all verified live against the real Laos corridor, not
assumed to carry over from the Thailand script):

    Cross        (API) platform_id is corridor-specific, not the Thailand
                 value of 80: 242=LAK, 259=THB, 251=USD (confirmed via each
                 corridor's rate_key: "LA:LAK" / "LA:THB" / "LA:USD"). The
                 first-remit bonus offset Thailand hardcodes as a flat +200
                 THB varies by corridor here (70,000 LAK / 100 THB / 3 USD
                 observed live) - read topup_amount from a query's own
                 response rather than hardcoding an offset.

    GmoneyTrans  (API) payout_country must be the exact string
                 "Lao People`s Democratic Republic" - note that's a literal
                 backtick, not an apostrophe, copied verbatim from the
                 site's own country-picker markup (id="11706LAKLAO"). Plain
                 spellings like "Laos" return an empty response, not an
                 error, which is why an earlier blind-guess pass concluded
                 (wrongly) that GmoneyTrans didn't serve Laos at all.
                 payment_type is "Bank Account" for every Bank-Deposit block
                 (LAK/THB/USD) and "Cash Pickup" for the USD Cash-Pickup
                 block - confirmed live that the same endpoint already
                 accepted "Bank Account" for THB/USD once GmoneyTrans added
                 those corridors, no new backend needed.

    Hanpass      (API) toCountryCode="LA". Works cleanly for all three
                 currencies (LAK/USD/THB) - just inputCurrencyCode changes.

    E9Pay        (API - binary search, same technique as Thailand's
                 fetch_e9pay) nation codes are LA02=LAK, LA01=USD; no THB
                 code exists (LA03-LA10 all error), matching the sheet,
                 which has no E9Pay row in the THB block. E9Pay's own
                 *website* rejects 15,000,000 LAK typed directly into the
                 calculator, but calling the underlying API and binary
                 searching over the KRW send-amount hits the target LAK
                 receive-amount cleanly - no scaling workaround needed.
                 Fee is a flat 5,000 by convention (same reasoning as
                 Thailand's fetch_e9pay): the API's own REMIT_FEE field
                 reports 0, with E9Pay's real margin embedded in the
                 exchange-rate spread instead.

    GME          (Mobile API, all 6 corridors - see above.) The Playwright/
                 website fallback (Bank Deposit only) selects the country
                 picker's "Laos (LAK)" / "Laos (USD)" entry, which defaults
                 to CASH PAYMENT unlike Thailand's BANK DEPOSIT default -
                 #sendingType must be explicitly set to "2" before reading
                 the rate. Cash Payment is never attempted there; it errors
                 regardless of amount (see fetch_gme_browser_laos).

Note: unlike fetch_thailand_rates.py, this script does not post anything
anywhere - it only fetches and writes to the workbook.

Setup:
    pip install requests openpyxl playwright
    playwright install chromium   # only needed for the GME fallback path

    Create gme_mobile_secrets.json next to this script (see Auth, above)
    before the GME mobile fetch will work.

Usage:
    python fetch_laos_rates.py \
        --input "Exchange rate daily report.xlsx" \
        --output "Exchange rate daily report (updated).xlsx"

    # Preview fetched values without touching the file:
    python fetch_laos_rates.py --input "Exchange rate daily report.xlsx" --dry-run

    # Skip GME entirely (both the mobile API and the Playwright fallback),
    # leaving all 6 of its cells untouched:
    python fetch_laos_rates.py --input "..." --output "..." --skip-gme
"""

import argparse
import base64
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests

try:
    import openpyxl
    from openpyxl.styles import PatternFill
except ImportError:  # pragma: no cover
    sys.exit("Missing dependency: pip install openpyxl")

import known_fees
import price_history


BASE_LAK = 15_000_000
BASE_USD = 1_000
BASE_THB = 26_000

SHEET_NAME = "Laos INPUT New"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

# Row of each provider inside the INPUT AREA (columns L:O) on
# "Laos INPUT New", one map per block. L=Provider, M=KRW(1), N=Service
# fee(2), O=Total(1)+(2).
BLOCK1_ROWS = {
    "GME (RIA-Other banks)": 8,
    "GME (Moneygram-BCEL)": 9,
    "GmoneyTrans": 10,
    "Hanpass": 11,
    "E9Pay": 12,
    "Cross": 13,
}
BLOCK2_ROWS = {"GME": 16, "Cross": 17, "Hanpass": 18, "GmoneyTrans": 19, "KEB Hana": 20}
BLOCK3_ROWS = {"GME": 23, "Hanpass": 24, "Cross": 25, "GmoneyTrans": 26}
BLOCK4_ROWS = {"GME (WU)": 30, "Hanpass": 31}
BLOCK5_ROWS = {"GME (Moneygram)": 35, "Cross": 36, "Hanpass": 37, "GmoneyTrans": 38, "E9Pay": 39}


# --------------------------------------------------------------------------
# GME mobile API (all 6 corridors) - see module docstring for the Auth
# section and how to refresh gme_mobile_secrets.json.
# --------------------------------------------------------------------------

GME_MOBILE_URL = "https://mobileapi.gmeremit.com:8002/api/v1/mobile/sendmoney/calculateV2"

# Overridable via env var - same GME_SECRETS_PATH override dashboard_server.py
# uses, since Cloud Run mounts this file at a different path (a
# Secret-Manager-backed volume, one secret per directory - see that file's
# own AUTH_PATH/SECRET_KEY_PATH/GME_SECRETS_PATH comment). Keeping both in
# sync matters: dashboard_server.py's own live-verify call passes an explicit
# secrets dict and never touches this default, but fetch_all()'s pipeline run
# does load through this path, so a mismatch here silently drops every GME
# mobile-API corridor without a browser fallback (confirmed - happened for
# real on the first Cloud Run deploy after paths diverged).
GME_MOBILE_SECRETS_PATH = Path(
    os.environ.get("GME_SECRETS_PATH", str(Path(__file__).with_name("gme_mobile_secrets.json")))
)

# Static per-app headers - identical across every captured request regardless
# of account/session, so unlike Authorization/Cookie these are plain
# constants here rather than secrets-file content (they're baked into the
# compiled app itself, visible to anyone who inspects it).
GME_MOBILE_STATIC_HEADERS = {
    "Accept": "*/*",
    "clientId": "l7xxfad8d2746d824f4cba3c872b1ac13eed",
    "Accept-Language": "en;q=1.0",
    "VersionName": "7.18.3",
    "Platform": "IOS",
    "Cache-Control": "no-store",
    "User-Agent": "GME Remit/7.18.3 (com.gme.gmeremit; build:7; iOS 26.5.2) Alamofire/5.9.1",
    "lang": "th",
    "GME-TOKEN": "39587YT398@FBQOW8RY3#948R7GB@CNEQW987GF87$TD18$1981..919@@##joghndvberteiru",
    "Content-Type": "application/json",
}

# Each GME corridor this sheet needs: which sheet row it goes to, the
# request's currency/amount/serviceType ("2"=Bank Deposit, "1"=Cash Pickup),
# and the placeholder receiver's pAgent/receiverId/payOutPartner triplet
# (captured live per corridor - see module docstring). Keyed the same way
# ROW_SOURCES keys every other provider's results.
GME_MOBILE_CORRIDORS = {
    ("GME (RIA-Other banks)", "LAK"): {
        "row": BLOCK1_ROWS["GME (RIA-Other banks)"],
        "pCurrency": "LAK", "pAmount": BASE_LAK, "serviceType": "2",
        "pAgent": "2123439", "receiverId": "3105397", "payOutPartner": "393865",
    },
    ("GME (Moneygram-BCEL)", "LAK"): {
        "row": BLOCK1_ROWS["GME (Moneygram-BCEL)"],
        "pCurrency": "LAK", "pAmount": BASE_LAK, "serviceType": "2",
        "pAgent": "1702057", "receiverId": "3005762", "payOutPartner": "798242",
    },
    ("GME", "USD"): {
        "row": BLOCK2_ROWS["GME"],
        "pCurrency": "USD", "pAmount": BASE_USD, "serviceType": "2",
        "pAgent": "1702056", "receiverId": "3107824", "payOutPartner": "798242",
    },
    ("GME", "THB"): {
        "row": BLOCK3_ROWS["GME"],
        "pCurrency": "THB", "pAmount": BASE_THB, "serviceType": "2",
        "pAgent": "1702058", "receiverId": "3107823", "payOutPartner": "798242",
    },
    ("GME (WU)", "LAK"): {
        "row": BLOCK4_ROWS["GME (WU)"],
        "pCurrency": "LAK", "pAmount": BASE_LAK, "serviceType": "1",
        "pAgent": "1766325", "receiverId": "3107828", "payOutPartner": "1765579",
    },
    ("GME (Moneygram)", "USD"): {
        "row": BLOCK5_ROWS["GME (Moneygram)"],
        "pCurrency": "USD", "pAmount": BASE_USD, "serviceType": "1",
        "pAgent": "946385", "receiverId": "3107820", "payOutPartner": "798242",
    },
}

# The 2 corridors the Playwright/website fallback can actually cover (Bank
# Deposit only) - see fetch_gme_browser_laos.
GME_BROWSER_FALLBACK_KEYS = [("GME (RIA-Other banks)", "LAK"), ("GME", "USD")]


def load_gme_mobile_secrets(path=GME_MOBILE_SECRETS_PATH):
    """Load the GME mobile API session secrets (bearer token, cookie,
    account username) from a local untracked file - see module docstring for
    the exact format and how to refresh it. Returns None (rather than
    raising) if the file is missing or its token has already expired, so the
    caller can fall back to Playwright for the corridors it can cover
    instead of crashing the whole run.
    """
    if not path.exists():
        print(f"GME mobile secrets not found at {path} - see module docstring to capture one.")
        return None

    with open(path) as f:
        secrets = json.load(f)

    try:
        jwt = secrets["authorization"].removeprefix("Bearer ").strip()
        payload_b64 = jwt.split(".")[1]
        padded = payload_b64 + "=" * (-len(payload_b64) % 4)
        claims = json.loads(base64.urlsafe_b64decode(padded))
        exp = datetime.fromtimestamp(claims["exp"], tz=timezone.utc)
    except Exception:
        return secrets  # can't decode it - let the live call itself succeed/fail

    if exp <= datetime.now(tz=timezone.utc):
        print(f"GME mobile token expired at {exp.isoformat()} - re-capture via Charles and update {path}.")
        return None

    return secrets


# --------------------------------------------------------------------------
# Per-provider fetchers. Each returns {"krw": int, "fee": int}.
# --------------------------------------------------------------------------

CROSS_PLATFORM_IDS = {"LAK": 242, "THB": 259, "USD": 251}


def fetch_cross(amount, currency, session=None):
    """Cross (crossenf.com) - Laos corridor. platform_id is currency-specific
    (242=LAK, 259=THB, 251=USD), unlike Thailand's single platform_id=80.

    The first-remit bonus offset isn't a fixed +200 like Thailand - it varies
    by corridor (70,000 LAK / 100 THB / 3 USD observed live), so read
    topup_amount from an initial query's response and cancel it out by
    requerying at amount + topup_amount, rather than hardcoding an offset.
    """
    s = session or requests
    url = "https://crossenf.com/v2/outbound/quote/"
    platform_id = CROSS_PLATFORM_IDS[currency]

    def quote(receiving_amount):
        params = {
            "platform_id": platform_id,
            "quote_type": "receive",
            "sending_amount": "null",
            "receiving_amount": receiving_amount,
            "use_max_point": "true",
            "deposit_type": "Manual",
            "apply_user_limit": 0,
            "is_home": 0,
        }
        r = s.get(url, params=params, headers=HEADERS, timeout=15)
        r.raise_for_status()
        return r.json()["data"]

    baseline = quote(amount)
    topup = int(baseline["topup_amount"])
    data = quote(amount + topup) if topup else baseline
    return {"krw": int(data["sending_amount"]), "fee": int(data["fee"])}


def fetch_gmoneytrans_laos(amount, currency, payment_type, session=None):
    """GmoneyTrans - same endpoint as Thailand, but payout_country must be
    the exact string below (a literal backtick, copied from the site's own
    country-picker markup) - plain spellings like "Laos" return an empty
    response rather than an error, which looks like "not supported" unless
    you check the country string carefully.
    """
    s = session or requests
    url = "https://mapi.gmoneytrans.net/exratenew1/ajx_calcRate.asp"
    params = {
        "receive_amount": amount,
        "payout_country": "Lao People`s Democratic Republic",
        "total_collected": "",
        "payment_type": payment_type,
        "currencyType": currency,
    }
    r = s.get(url, params=params, headers=HEADERS, timeout=15)
    r.raise_for_status()
    text = r.text
    krw_match = re.search(r"sendAmount--td_clm--(\d+)--td_end", text)
    fee_match = re.search(r"serviceCharge--td_clm--(\d+)--td_end", text)
    if not krw_match or not fee_match:
        raise ValueError(f"Unexpected GmoneyTrans response format: {text[:200]!r}")
    return {"krw": int(krw_match.group(1)), "fee": int(fee_match.group(1))}


def fetch_hanpass_laos(amount, currency, session=None):
    """Hanpass - toCountryCode="LA". remittanceOption is accepted but
    silently ignored by this endpoint (Bank Transfer vs Cash Pickup return
    byte-identical responses), so callers needing a Cash-Pickup figure
    should just reuse this same Bank-Deposit result rather than requerying.
    """
    s = session or requests
    url = "https://old.hanpass.com/getCost"
    payload = {
        "inputAmount": str(amount),
        "inputCurrencyCode": currency,
        "toCurrencyCode": "KRW",
        "toCountryCode": "LA",
        "remittanceOption": "BANK_TRANSFER",
        "mtoServiceCenterCode": "",
        "mtoProviderCode": "",
        "lang": "en",
    }
    r = s.post(url, json=payload, headers=HEADERS, timeout=15)
    r.raise_for_status()
    data = r.json()
    return {
        "krw": int(round(float(data["depositAmount"]))),
        "fee": int(round(float(data["transferFee"]))),
    }


E9PAY_NATION_CODES = {"LAK": "LA02", "USD": "LA01"}
# Rough KRW brackets to start the binary search from, based on live-observed
# rates (~15 LAK per KRW, ~1,490 KRW per USD). No THB code exists for Laos.
E9PAY_KRW_BRACKETS = {
    "LAK": lambda amount: (int(amount / 20), int(amount / 10)),
    "USD": lambda amount: (int(amount * 1200), int(amount * 1700)),
}


def fetch_e9pay_laos(amount, currency, session=None, fixed_fee=None, max_iter=40):
    """E9pay - nation code LA02=LAK / LA01=USD in place of Thailand's TH03;
    no THB code exists (LA03-LA10 all error).

    Same binary-search technique as Thailand's fetch_e9pay (the calculator
    only accepts a KRW *send* amount and returns the receive amount). The
    *website's* input field rejects 15,000,000 LAK typed in directly, but
    querying the API and searching for the KRW amount that yields the
    target receive-amount works cleanly - no scaling workaround needed.

    `fixed_fee` defaults to whatever check_competitor_fees.py last read off
    E9Pay's own homepage widget (falls back to 5000 if that's never run) -
    see known_fees.py. Same flat, country-independent fee as Thailand's
    fetch_e9pay - confirmed live E9Pay's displayed fee doesn't vary by
    destination.
    """
    if fixed_fee is None:
        fixed_fee = known_fees.get_fee("E9Pay", default=5000)
    s = session or requests
    url = "https://www.e9pay.co.kr/cmm/calcExchangeRate.do"
    natn_cod = E9PAY_NATION_CODES[currency]

    def calc(krw_amount):
        data = {
            "DEFRAY_AMOUNT": krw_amount,
            "SEND_NATN_COD": "KR",
            "CRNCY_COD": "KRW",
            "RCVER_EXPECT_NATN_COD": natn_cod,
            "RCVER_EXPECT_CRNCY_COD": currency,
            "SIMULATION_YN": "Y",
            "OVSE_FEE_PROMOTION_YN": "N",
            "LANG_COD": "",
        }
        r = s.post(url, data=data, headers=HEADERS, timeout=15)
        r.raise_for_status()
        outer = r.json()
        inner = json.loads(outer["data"])
        return float(inner["RCVER_EXPECT_RECPT_AMOUNT"])

    lo, hi = E9PAY_KRW_BRACKETS[currency](amount)
    best = hi
    for _ in range(max_iter):
        mid = (lo + hi) // 2
        result = calc(mid)
        if result >= amount:
            hi = mid
            best = mid
        else:
            lo = mid
        if hi - lo <= 1:
            break
    return {"krw": best, "fee": fixed_fee}


def fetch_gme_mobile_laos(corridor, secrets, session=None, retries=2):
    """GME's own mobile-app API, captured via Charles Proxy against the
    user's own GME account/app (not the public website - see module
    docstring). Covers all 6 GME corridors, including the 3 the website has
    no source for at all (THB, LAK Cash Pickup, USD Cash Pickup), and can
    distinguish RIA vs Moneygram-BCEL Bank Deposit rates via payOutPartner,
    which the website's UI cannot.

    `corridor` is one value from GME_MOBILE_CORRIDORS. `secrets` is the dict
    returned by load_gme_mobile_secrets (Authorization/Cookie/username).

    Fee is always written as 0 (krw holds the endpoint's full all-in total)
    - same deliberate convention as GME's other rows: confirmed live that
    Data.sAmt + Data.scCharge == Data.collAmt, so collAmt alone is the
    all-in figure.
    """
    s = session or requests
    headers = dict(GME_MOBILE_STATIC_HEADERS)
    headers["Authorization"] = secrets["authorization"]
    headers["Cookie"] = secrets["cookie"]
    payload = {
        "calcBy": "p",
        "username": secrets["username"],
        "userId": secrets["username"],
        "pAmount": str(corridor["pAmount"]),
        "pAgent": corridor["pAgent"],
        "sCurrency": "KRW",
        "pCountryName": "Laos",
        "sCountry": "118",
        "pCurrency": corridor["pCurrency"],
        "receiverId": corridor["receiverId"],
        "rateType": "",
        "cAmount": "",
        "pCountry": "121",
        "serviceType": corridor["serviceType"],
        "payOutPartner": corridor["payOutPartner"],
        "paymentType": "autodebit",
    }

    last_err = None
    for attempt in range(retries + 1):
        try:
            r = s.post(GME_MOBILE_URL, json=payload, headers=headers, timeout=15)
            r.raise_for_status()
            res_json = r.json()
            data = res_json.get("Data")
            if not data or not isinstance(data, dict) or "collAmt" not in data:
                msg = res_json.get("Msg") or res_json.get("Message") or "No rate data returned"
                code = res_json.get("ErrorCode", "Unknown")
                raise ValueError(f"GME Mobile API ({code}): {msg}")
            return {"krw": int(round(float(data["collAmt"]))), "fee": 0}
        except Exception as e:
            last_err = e
            if attempt < retries:
                time.sleep(1.5)
                continue
            raise last_err


def fetch_gme_browser_laos(amount, currency, headless=True, browser=None):
    """Fallback only - see GME_BROWSER_FALLBACK_KEYS and module docstring.
    GME's own live consumer calculator at online.gmeremit.com - Bank Deposit
    only, and can't distinguish RIA from Moneygram-BCEL (only used when the
    mobile API in fetch_gme_mobile_laos didn't cover a corridor this run).
    Cash Payment is not attempted: it errors regardless of amount for Laos
    (see module docstring), so there's nothing usable to fetch there.

    Laos defaults to CASH PAYMENT the moment its country entry is selected
    (unlike Thailand, which defaults to BANK DEPOSIT) - #sendingType must be
    explicitly switched to "2" before setting the amount, or the calculator
    just re-errors on Cash Payment instead of pricing Bank Deposit.

    `browser` lets a caller pass in an already-launched Playwright browser
    so this can share one Chromium process with other browser-driven calls
    in the same run, matching fetch_thailand_rates.py's fetch_gme_browser.
    """
    from playwright.sync_api import sync_playwright

    def _run(b):
        page = b.new_page()
        try:
            page.goto("https://online.gmeremit.com/", wait_until="networkidle")

            page.evaluate("document.getElementById('nCountry').click()")
            page.wait_for_timeout(500)
            page.evaluate(
                """
                (currency) => {
                    const li = [...document.querySelectorAll('#toCurrUl li')]
                        .find(el => el.textContent.includes('Laos (' + currency + ')'));
                    li.click();
                }
                """,
                currency,
            )
            page.wait_for_timeout(800)

            # Force Bank Deposit - Laos defaults to Cash Payment on selection.
            page.evaluate(
                """
                () => {
                    const sel = document.getElementById('sendingType');
                    sel.value = '2';
                    sel.dispatchEvent(new Event('change', {bubbles: true}));
                }
                """
            )
            page.wait_for_timeout(500)

            page.evaluate(
                """
                (amount) => {
                    const el = document.getElementById('recAmt');
                    el.focus();
                    el.value = String(amount);
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                    el.dispatchEvent(new Event('blur', {bubbles: true}));
                }
                """,
                amount,
            )
            page.wait_for_timeout(2000)

            # Value is formatted like "1,012,462.00" - strip the decimal part
            # before stripping commas, or ".00" collapses into the digit
            # string and inflates the result 100x.
            krw_text = page.eval_on_selector("#numAmount", "el => el.value").split(".")[0]
            krw = int(re.sub(r"[^\d]", "", krw_text))

            # fee is intentionally always 0, same convention as Thailand's GME row.
            return {"krw": krw, "fee": 0}
        finally:
            page.close()

    if browser is not None:
        return _run(browser)

    with sync_playwright() as p:
        b = p.chromium.launch(headless=headless)
        try:
            return _run(b)
        finally:
            b.close()


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

# Each unique (provider, currency) quote this sheet needs, fetched exactly
# once and then distributed to every row that wants it (see ROW_SOURCES
# below) - several rows across blocks reuse the same fetched quote instead
# of querying twice.
FETCH_SPECS = [
    ("Cross", "LAK", lambda s: fetch_cross(BASE_LAK, "LAK", session=s)),
    ("Cross", "USD", lambda s: fetch_cross(BASE_USD, "USD", session=s)),
    ("Cross", "THB", lambda s: fetch_cross(BASE_THB, "THB", session=s)),
    ("GmoneyTrans", "LAK", lambda s: fetch_gmoneytrans_laos(BASE_LAK, "LAK", "Bank Account", session=s)),
    ("GmoneyTrans", "USD", lambda s: fetch_gmoneytrans_laos(BASE_USD, "USD", "Cash Pickup", session=s)),
    # GmoneyTrans added Bank Deposit for THB/USD later - same endpoint,
    # confirmed live it already accepts these params, no new backend needed.
    # "USD-BD" (not "USD") avoids colliding with the Cash Pickup entry above,
    # which is a genuinely different corridor despite sharing a currency -
    # this key is internal only, unrelated to the "GmoneyTrans" text actually
    # written to column L.
    ("GmoneyTrans", "THB", lambda s: fetch_gmoneytrans_laos(BASE_THB, "THB", "Bank Account", session=s)),
    ("GmoneyTrans", "USD-BD", lambda s: fetch_gmoneytrans_laos(BASE_USD, "USD", "Bank Account", session=s)),
    ("Hanpass", "LAK", lambda s: fetch_hanpass_laos(BASE_LAK, "LAK", session=s)),
    ("Hanpass", "USD", lambda s: fetch_hanpass_laos(BASE_USD, "USD", session=s)),
    ("Hanpass", "THB", lambda s: fetch_hanpass_laos(BASE_THB, "THB", session=s)),
    ("E9Pay", "LAK", lambda s: fetch_e9pay_laos(BASE_LAK, "LAK", session=s)),
    ("E9Pay", "USD", lambda s: fetch_e9pay_laos(BASE_USD, "USD", session=s)),
]

# Which sheet row each fetched (provider, currency) quote should be written
# to. Some quotes are written to more than one row (Hanpass/Cross reused
# across Bank-Deposit and Cash-Pickup blocks - see module docstring). GME's
# 6 corridors are each fetched and written separately (no reuse), unlike the
# old browser-only approach that could only tell currency apart, not payout
# rail.
ROW_SOURCES = [
    (BLOCK1_ROWS["GME (RIA-Other banks)"], ("GME (RIA-Other banks)", "LAK")),
    (BLOCK1_ROWS["GME (Moneygram-BCEL)"], ("GME (Moneygram-BCEL)", "LAK")),
    (BLOCK2_ROWS["GME"], ("GME", "USD")),
    (BLOCK3_ROWS["GME"], ("GME", "THB")),
    (BLOCK4_ROWS["GME (WU)"], ("GME (WU)", "LAK")),
    (BLOCK5_ROWS["GME (Moneygram)"], ("GME (Moneygram)", "USD")),
    (BLOCK1_ROWS["Cross"], ("Cross", "LAK")),
    (BLOCK1_ROWS["GmoneyTrans"], ("GmoneyTrans", "LAK")),
    (BLOCK1_ROWS["Hanpass"], ("Hanpass", "LAK")),
    (BLOCK1_ROWS["E9Pay"], ("E9Pay", "LAK")),
    (BLOCK2_ROWS["Cross"], ("Cross", "USD")),
    (BLOCK2_ROWS["Hanpass"], ("Hanpass", "USD")),
    (BLOCK2_ROWS["GmoneyTrans"], ("GmoneyTrans", "USD-BD")),
    (BLOCK2_ROWS["KEB Hana"], ("KEB Hana", "USD")),
    (BLOCK3_ROWS["Cross"], ("Cross", "THB")),
    (BLOCK3_ROWS["Hanpass"], ("Hanpass", "THB")),
    (BLOCK3_ROWS["GmoneyTrans"], ("GmoneyTrans", "THB")),
    (BLOCK4_ROWS["Hanpass"], ("Hanpass", "LAK")),  # reused from Block 1
    (BLOCK5_ROWS["Cross"], ("Cross", "USD")),  # reused from Block 2
    (BLOCK5_ROWS["Hanpass"], ("Hanpass", "USD")),  # reused from Block 2
    (BLOCK5_ROWS["GmoneyTrans"], ("GmoneyTrans", "USD")),
    (BLOCK5_ROWS["E9Pay"], ("E9Pay", "USD")),
]

# Human-readable corridor label per row, for price_history logging (see
# --log-history in main()). Keyed by row rather than by fetched (provider,
# currency), since a single fetched quote can be written into two different
# corridors (e.g. Hanpass's LAK quote covers both Block 1's Bank Deposit row
# and Block 4's Cash Pickup row) - "days more expensive" is a per-corridor
# question, so each corridor needs its own logged entry even when the
# underlying number is identical.
ROW_TO_CORRIDOR = {}
for _row in BLOCK1_ROWS.values():
    ROW_TO_CORRIDOR[_row] = "LAK Bank Deposit"
for _row in BLOCK2_ROWS.values():
    ROW_TO_CORRIDOR[_row] = "USD Bank Deposit"
for _row in BLOCK3_ROWS.values():
    ROW_TO_CORRIDOR[_row] = "THB Bank Deposit"
for _row in BLOCK4_ROWS.values():
    ROW_TO_CORRIDOR[_row] = "LAK Cash Pickup"
for _row in BLOCK5_ROWS.values():
    ROW_TO_CORRIDOR[_row] = "USD Cash Pickup"
del _row


def fetch_all(skip_gme=False, skip_kebhana=False):
    """Fetch every (provider, currency) quote this sheet needs. Quotes whose
    fetch raises an exception are omitted (and reported), so the caller can
    leave those cells untouched rather than write bad data - same fail-soft
    convention as fetch_thailand_rates.py.
    """
    results = {}
    errors = {}
    session = requests.Session()

    if not skip_gme:
        secrets = load_gme_mobile_secrets()
        if secrets is not None:
            for key, corridor in GME_MOBILE_CORRIDORS.items():
                provider, currency = key
                label = f"{provider} ({currency})"
                try:
                    print(f"Fetching {label} (mobile API) ...", end=" ", flush=True)
                    result = fetch_gme_mobile_laos(corridor, secrets, session=session)
                    # Same guard as Thailand's fetch_gme_browser (see
                    # price_history.is_plausible's docstring for the real
                    # incident that motivated it) - a "successful" API call
                    # can still return a wrong number (e.g. a stale/cached
                    # response, or the API silently defaulting a malformed
                    # field), and that wouldn't raise on its own the way a
                    # timeout does.
                    corridor_name = ROW_TO_CORRIDOR[corridor["row"]]
                    if not price_history.is_plausible("laos", corridor_name, provider, result["krw"]):
                        raise ValueError(f"{label}'s fetched KRW ({result['krw']:,}) is implausibly far from its recent history")
                    results[key] = result
                    print(f"KRW={result['krw']:,}  fee={result['fee']:,}")
                except Exception as e:  # noqa: BLE001
                    errors[key] = e
                    print(f"FAILED ({e})")

        # Playwright fallback, only for the 2 corridors the website can do,
        # and only for whichever of those the mobile API didn't get above
        # (missing/expired secrets, or that specific REST call failing).
        fallback_keys = [k for k in GME_BROWSER_FALLBACK_KEYS if k not in results]
        if fallback_keys:
            try:
                from playwright.sync_api import sync_playwright
            except ImportError:
                rows = [GME_MOBILE_CORRIDORS[k]["row"] for k in fallback_keys]
                print(f"Playwright not installed - can't fall back for GME rows {rows}")
            else:
                playwright_instance = sync_playwright().start()
                browser_ctx = playwright_instance.chromium.launch(headless=True)
                try:
                    for key in fallback_keys:
                        provider, currency = key
                        label = f"{provider} ({currency})"
                        amount = GME_MOBILE_CORRIDORS[key]["pAmount"]
                        try:
                            print(f"Fetching {label} (browser fallback, Bank Deposit) ...", end=" ", flush=True)
                            result = fetch_gme_browser_laos(amount, currency, browser=browser_ctx)
                            # Same guard as the mobile-API path above and
                            # Thailand's fetch_gme_browser - this is the exact
                            # kind of page-interaction-driven fetch (read
                            # #numAmount after simulating input) that produced
                            # a silently-wrong-but-not-thrown value for real
                            # on Thailand's GME row.
                            corridor_name = ROW_TO_CORRIDOR[GME_MOBILE_CORRIDORS[key]["row"]]
                            if not price_history.is_plausible("laos", corridor_name, provider, result["krw"]):
                                raise ValueError(f"{label}'s fetched KRW ({result['krw']:,}) is implausibly far from its recent history")
                            results[key] = result
                            print(f"KRW={result['krw']:,}  fee={result['fee']:,}")
                        except Exception as e:  # noqa: BLE001
                            errors[key] = e
                            print(f"FAILED ({e})")
                finally:
                    browser_ctx.close()
                    playwright_instance.stop()

    if not skip_kebhana:
        try:
            print("Fetching KEB Hana (USD) (browser, live rate page) ...", end=" ", flush=True)
            import fetch_kebhana_rates

            result = fetch_kebhana_rates.fetch_kebhana_quote("USD", BASE_USD)
            results[("KEB Hana", "USD")] = result
            print(f"KRW={result['krw']:,}  fee={result['fee']:,}")
        except Exception as e:  # noqa: BLE001
            errors[("KEB Hana", "USD")] = e
            print(f"FAILED ({e})")

    for provider, currency, fn in FETCH_SPECS:
        label = f"{provider} ({currency})"
        try:
            print(f"Fetching {label} ...", end=" ", flush=True)
            result = fn(session)
            results[(provider, currency)] = result
            print(f"KRW={result['krw']:,}  fee={result['fee']:,}")
        except Exception as e:  # noqa: BLE001 - report and continue
            errors[(provider, currency)] = e
            print(f"FAILED ({e})")

    if errors:
        print("\nQuotes that failed and were left untouched in the sheet:")
        for (provider, currency), e in errors.items():
            print(f"  - {provider} ({currency}): {e}")

    return results


def rows_to_write(fetched):
    """Map each sheet row to the fetched quote it should receive, skipping
    any row whose source quote failed to fetch."""
    return {row: fetched[key] for row, key in ROW_SOURCES if key in fetched}


def stale_rows_from(fetched):
    """Rows this script actively tries to fetch every run, but whose
    underlying quote failed this time - see STALE_FILL in write_to_workbook.
    """
    return {row for row, key in ROW_SOURCES if key not in fetched}


# Flags a row whose value is left over from a previous successful run rather
# than freshly fetched. Writing "N/A" or leaving the cell blank were both
# rejected (see module docstring / plan doc): an error value in M or N makes
# O ("=M+N") error out, and the ranked table's LARGE(...) formulas don't
# ignore an error anywhere in their range - one bad cell breaks the ranking
# for the whole block. A blank cell avoids that (blank behaves as 0 in
# arithmetic) but then ranks as the cheapest option, which looks like an
# unrealistically good deal rather than looking missing. So the value stays
# untouched and only the fill color changes - a purely visual flag that
# can't affect any formula.
STALE_FILL = PatternFill(fill_type="solid", fgColor="FFC000")
NO_FILL = PatternFill(fill_type=None)


def write_to_workbook(writes, stale_rows, input_path, output_path, sheet_name=SHEET_NAME):
    wb = openpyxl.load_workbook(input_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        sys.exit(f"Sheet '{sheet_name}' not found. Sheets present: {wb.sheetnames}")
    ws = wb[sheet_name]

    for row, vals in writes.items():
        ws.cell(row=row, column=13, value=vals["krw"])  # column M
        ws.cell(row=row, column=14, value=vals["fee"])  # column N
        # column O (Total price) is left as the existing "=M{row}+N{row}" formula
        # Clear any stale flag from a previous run now that this row is fresh.
        ws.cell(row=row, column=13).fill = NO_FILL
        ws.cell(row=row, column=14).fill = NO_FILL

    for row in stale_rows:
        ws.cell(row=row, column=13).fill = STALE_FILL
        ws.cell(row=row, column=14).fill = STALE_FILL

    wb.save(output_path)
    print(f"\nSaved: {output_path}")
    print(
        "Open it in Excel (or press Ctrl+Alt+F9 to force a full recalc) and the "
        "left-hand tables will re-rank themselves by Total price automatically."
    )
    if stale_rows:
        print(
            f"Rows {sorted(stale_rows)} kept their previous value but are "
            "flagged amber - their fetch failed this run, see errors above."
        )


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", required=True, help="Path to Exchange rate daily report.xlsx")
    parser.add_argument("--output", help="Path to save the updated workbook (defaults to --input)")
    parser.add_argument("--dry-run", action="store_true", help="Fetch and print only; don't write the file")
    parser.add_argument("--skip-gme", action="store_true", help="Skip GME entirely, both the mobile API and the Playwright fallback (leaves all 6 GME cells untouched)")
    parser.add_argument("--skip-kebhana", action="store_true", help="Skip the Playwright/KEB Hana step")
    parser.add_argument(
        "--log-history", action="store_true",
        help="Append this run's prices/gaps to price_history.db (opt-in - only the scheduled automation should pass this, not ad-hoc/test runs)",
    )
    args = parser.parse_args()

    output_path = args.output or args.input

    fetched = fetch_all(skip_gme=args.skip_gme, skip_kebhana=args.skip_kebhana)

    print("\n--- Summary ---")
    for (provider, currency), vals in fetched.items():
        total = vals["krw"] + vals["fee"]
        print(f"{provider:<12} {currency:<3}  KRW={vals['krw']:>10,}  fee={vals['fee']:>7,}  total={total:>10,}")

    if args.dry_run:
        print("\n(dry run - workbook not modified)")
        return

    writes = rows_to_write(fetched)
    stale_rows = stale_rows_from(fetched)
    write_to_workbook(writes, stale_rows, args.input, output_path)

    if args.log_history:
        entries = [
            {
                "corridor": ROW_TO_CORRIDOR[row],
                "provider": key[0],
                "krw": fetched[key]["krw"],
                "fee": fetched[key]["fee"],
                # Block 1 (LAK Bank Deposit) has two GME rows - the sheet's
                # own Price Gap formulas there are anchored to Moneygram-BCEL
                # (confirmed via the workbook's own Q7 reference cell), so
                # that's the baseline here too rather than RIA.
                "is_gme_baseline": ROW_TO_CORRIDOR[row] == "LAK Bank Deposit" and key[0] == "GME (Moneygram-BCEL)",
            }
            for row, key in ROW_SOURCES
            if key in fetched
        ]
        n = price_history.log_run("laos", entries)
        print(f"Logged {n} rows to {price_history.DEFAULT_DB_PATH}")


if __name__ == "__main__":
    main()
