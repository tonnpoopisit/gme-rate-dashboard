#!/usr/bin/env python3
"""
fetch_thailand_rates.py

Fetches live "send 26,000 THB from Korea" remittance quotes from each
competitor's hidden calculator API and writes the KRW (1) / Service fee (2)
values into the INPUT AREA (columns M:N, rows 7-16) of the
"Thailand INPUT New" sheet in Exchange rate daily report.xlsx.

IMPORTANT - this script does NOT touch the left-hand "DO NOT EDIT THIS SIDE"
table. That table is already driven entirely by Excel formulas:

    D8:D16  -> =IFERROR(INDEX($L$7:$L$16,MATCH(LARGE($O$7:$O$16,ROW()-7),...
    F8:F16  -> KRW (1)      (same LARGE/MATCH/INDEX pattern)
    G8:G16  -> Service fee (2)
    H8:H16  -> =IFERROR(F+G,"")                       (Total price)
    I8:I16  -> =IF(H="","",H-VLOOKUP("GME",...))       (Price gap)
    O7:O16  -> =M+N                                    (Total price, INPUT AREA)

Those formulas auto-rank/sort by Total price and compute the price gap
against GME on their own. This script only ever writes to M7:N16 (the "edit
blue cells" INPUT AREA) - once it does, Excel recalculates everything else
the moment the file is opened (or you press Ctrl+Alt+F9).

Provider -> INPUT AREA row map (fixed; matches the existing sheet layout):
    GME          row 7   (Web UI via Playwright - fetched live like everyone
                          else, from GME's own consumer calculator at
                          online.gmeremit.com. It's a classic ASP.NET
                          WebForms page (postback to Default.aspx), not a
                          clean API, so it's driven the same way as
                          WireBarley. Fee is always written as 0 - GME's row
                          has historically used a single all-in KRW figure
                          rather than a decomposed rate+fee, and that
                          convention is intentional; see fetch_gme_browser.)
    Cross        row 8   (API)
    SBI Cosmoney row 9   (Web UI via Playwright - the calc/amount endpoint
                          works fine over plain HTTP but SBI's domain is
                          behind Cloudflare, which silently blocks direct
                          requests calls; see fetch_sbi_browser. Fee is a
                          fixed published fee, not something the site returns.)
    JRF/JPRemit  row 10  (API)
    E9Pay        row 11  (API - binary search, see fetch_e9pay)
    GmoneyTrans  row 13  (API)
    Coinshot     row 14  (API - confirmed via live capture, see fetch_coinshot)
    Hanpass      row 15  (API)
    WireBarley   row 16  (Web UI via Playwright - API request body is
                          encrypted client-side, so it cannot be called
                          directly)

Setup:
    pip install requests openpyxl playwright
    playwright install chromium

    IMPORTANT if this will run outside an interactive Claude Code session
    (e.g. via Windows Task Scheduler, as run_hourly_report.ps1 does): on a
    machine where Claude Code itself runs as a sandboxed/packaged Windows
    app, `playwright install`'s default browser cache (the ms-playwright
    folder under %LOCALAPPDATA%) is a reparse point Windows silently
    redirects into that app's isolated container storage - fully visible to
    anything launched through Claude Code, but genuinely invisible (not a
    permissions error - the path just doesn't exist from that process's
    point of view) to an ordinary process launched by something like Task
    Scheduler outside that container. Confirmed via a live diagnostic
    scheduled task, not a guess. Fix: install to a plain, non-AppData path
    both contexts can see identically, and set the same env var whenever
    this script runs:
        $env:PLAYWRIGHT_BROWSERS_PATH = "C:/PlaywrightBrowsers"
        playwright install chromium

Usage:
    python fetch_thailand_rates.py \
        --input "Exchange rate daily report.xlsx" \
        --output "Exchange rate daily report (updated).xlsx"

    # Preview fetched values without touching the file:
    python fetch_thailand_rates.py --input "Exchange rate daily report.xlsx" --dry-run

    # Skip the browser step (e.g. no Playwright installed) and keep
    # WireBarley's existing value in the sheet:
    python fetch_thailand_rates.py --input "..." --output "..." --skip-wirebarley

    # Skip GME's own live fetch too (keeps GME's existing cell untouched):
    python fetch_thailand_rates.py --input "..." --output "..." --skip-gme

    # Skip SBI's browser-driven fetch too (keeps SBI's existing cell untouched):
    python fetch_thailand_rates.py --input "..." --output "..." --skip-sbi
"""

import argparse
import json
import re
import sys

import requests

try:
    import openpyxl
    from openpyxl.styles import PatternFill
except ImportError:  # pragma: no cover
    sys.exit("Missing dependency: pip install openpyxl")

import known_fees
import price_history


BASE_THB = 26000

# Row of each provider inside the INPUT AREA (columns L:O) on
# "Thailand INPUT New". L=Provider, M=KRW(1), N=Service fee(2), O=Total(1)+(2).
INPUT_AREA_ROWS = {
    "GME": 7,
    "Cross": 8,
    "SBI": 9,
    "JRF": 10,
    "E9Pay": 11,
    "KEB Hana": 12,
    "GmoneyTrans": 13,
    "Coinshot": 14,
    "Hanpass": 15,
    "WireBarley": 16,
}

SHEET_NAME = "Thailand INPUT New"

# Cell fill applied to a provider's M/N cells when its fetch fails this run
# (the last known value is kept, not blanked - see stale_rows_from below);
# cleared back to NO_FILL the next time that provider fetches successfully.
STALE_FILL = PatternFill(fill_type="solid", fgColor="FFC000")
NO_FILL = PatternFill(fill_type=None)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}


# --------------------------------------------------------------------------
# Per-provider fetchers. Each returns {"krw": int, "fee": int}.
# --------------------------------------------------------------------------

def fetch_cross(thb=BASE_THB, session=None):
    """Cross (crossenf.com) - public GET, no auth needed.

    CORRECTION (previous version was wrong): the "+200 THB" bonus badge
    is NOT cosmetic - it genuinely subsidizes `sending_amount`. Confirmed by
    querying the live endpoint at receiving_amount = 25,800 / 26,000 / 26,200:
    `topup_amount` stays a flat 200 THB regardless of the requested amount,
    and `sending_amount` always equals `(receiving_amount - 200) * rate`.
    So requesting the nominal `thb` (26,000) returns a `sending_amount` that
    is discounted by a one-time new-customer promo (the response even flags
    `is_first_remit: true` on every anonymous query) - not the standing,
    repeat-customer rate GME should be compared against.

    Requesting `thb + 200` instead cancels the bonus out exactly:
    `(26,200 - 200) * rate == 26,000 * rate`, the full undiscounted price for
    26,000 THB. `fee` is unaffected by any of this (confirmed constant
    across all three test amounts).
    """
    s = session or requests
    url = "https://crossenf.com/v2/outbound/quote/"
    params = {
        "platform_id": 80,
        "quote_type": "receive",
        "sending_amount": "null",
        "receiving_amount": thb + 200,  # cancels out the flat +200 THB signup bonus
        "use_max_point": "true",
        "deposit_type": "Manual",
        "apply_user_limit": 0,
        "is_home": 0,
    }
    r = s.get(url, params=params, headers=HEADERS, timeout=15)
    r.raise_for_status()
    data = r.json()["data"]
    return {"krw": int(data["sending_amount"]), "fee": int(data["fee"])}


def fetch_jrf(thb=BASE_THB, session=None):
    """JRF / JPRemit - plain unauthenticated GET on the rate subdomain."""
    s = session or requests
    url = "https://rateweb.jpremit.co.kr/JrfKorea/GetRate"
    params = {"country": "TH", "payout": "B", "amount": thb, "currency": "THB", "calcby": "P"}
    r = s.get(url, params=params, headers=HEADERS, timeout=15)
    r.raise_for_status()
    data = r.json()["data"]
    return {
        "krw": int(round(float(data["collecT_AMT"]))),
        "fee": int(round(float(data["servicE_CHARGE"]))),
    }


def fetch_gmoneytrans(thb=BASE_THB, session=None):
    """GmoneyTrans - the real calculator lives on a separate subdomain
    (mapi.gmoneytrans.net), reached via the iframe on the public homepage."""
    s = session or requests
    url = "https://mapi.gmoneytrans.net/exratenew1/ajx_calcRate.asp"
    params = {
        "receive_amount": thb,
        "payout_country": "Thailand",
        "total_collected": "",
        "payment_type": "Bank Account",
        "currencyType": "THB",
    }
    r = s.get(url, params=params, headers=HEADERS, timeout=15)
    r.raise_for_status()
    text = r.text
    krw_match = re.search(r"sendAmount--td_clm--(\d+)--td_end", text)
    fee_match = re.search(r"serviceCharge--td_clm--(\d+)--td_end", text)
    if not krw_match or not fee_match:
        raise ValueError(f"Unexpected GmoneyTrans response format: {text[:200]!r}")
    return {"krw": int(krw_match.group(1)), "fee": int(fee_match.group(1))}


def fetch_hanpass(thb=BASE_THB, session=None):
    """Hanpass - real calculator is on old.hanpass.com (the corporate
    hanpass.com/en site has no calculator). Direct POST, no cookies needed."""
    s = session or requests
    url = "https://old.hanpass.com/getCost"
    payload = {
        "inputAmount": str(thb),
        "inputCurrencyCode": "THB",
        "toCurrencyCode": "KRW",
        "toCountryCode": "TH",
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


def fetch_e9pay(thb=BASE_THB, session=None, fixed_fee=None, max_iter=40):
    """E9pay - the calculator only accepts a KRW *send* amount and returns
    the resulting THB *receive* amount (the inverse of what we need), and
    the receive-amount field in the UI isn't directly editable either. So we
    binary-search the KRW input until the resulting THB output lands
    exactly on `thb`.

    `fixed_fee` defaults to whatever check_competitor_fees.py last read off
    E9Pay's own homepage widget (falls back to 5000 if that's never run) -
    see known_fees.py.
    """
    if fixed_fee is None:
        fixed_fee = known_fees.get_fee("E9Pay", default=5000)
    s = session or requests
    url = "https://www.e9pay.co.kr/cmm/calcExchangeRate.do"

    def calc(krw_amount):
        data = {
            "DEFRAY_AMOUNT": krw_amount,
            "SEND_NATN_COD": "KR",
            "CRNCY_COD": "KRW",
            "RCVER_EXPECT_NATN_COD": "TH03",
            "RCVER_EXPECT_CRNCY_COD": "THB",
            "SIMULATION_YN": "Y",
            "OVSE_FEE_PROMOTION_YN": "N",
            "LANG_COD": "",
        }
        r = s.post(url, data=data, headers=HEADERS, timeout=15)
        r.raise_for_status()
        outer = r.json()
        inner = json.loads(outer["data"])
        return float(inner["RCVER_EXPECT_RECPT_AMOUNT"])

    lo, hi = int(thb * 40), int(thb * 50)  # bracket around ~44-45 KRW/THB
    best = hi
    for _ in range(max_iter):
        mid = (lo + hi) // 2
        result = calc(mid)
        if result >= thb:
            hi = mid
            best = mid
        else:
            lo = mid
        if hi - lo <= 1:
            break
    return {"krw": best, "fee": fixed_fee}


def fetch_coinshot(thb=BASE_THB, session=None):
    """Coinshot - confirmed via live capture (driving the real page's
    currency selectors and reading the network request it fires). Two things
    the original guess got wrong:

    1. It's a Spring double-submit CSRF endpoint like SBI's, just with a
       different header name (`X-CSRF-TOKEN`) and cookie set implicitly via
       session (not readable from JS - it's a plain session cookie, not an
       XSRF-TOKEN cookie), so the token has to come from the `_csrf` meta tag
       on a GET'd page rather than from cookies directly.
    2. The body is form-encoded (matching jQuery's default, not JSON), and
       needs a `feeIncluded: false` field alongside the amount/currencies -
       omitting it or getting the encoding wrong makes the endpoint fall back
       to serving the site's "please login" HTML page instead of JSON.

    The response includes the real fee (`fromFee`) directly, so no more
    guessing a fixed_fee - the old default (2,500) was wrong; live capture
    showed 5,000 KRW.
    """
    s = session or requests.Session()
    home = s.get("https://coinshot.org/main", headers=HEADERS, timeout=15)
    m = re.search(r'name="_csrf" content="([^"]+)"', home.text)
    if not m:
        raise ValueError("Could not obtain Coinshot CSRF token")
    token = m.group(1)

    url = "https://coinshot.org/calculate/sending"
    payload = {
        "receivingAmount": str(thb),
        "sendingCurrency": "KRW",
        "receivingCurrency": "THB",
        "feeIncluded": "false",
    }
    headers = {**HEADERS, "X-CSRF-TOKEN": token, "X-Requested-With": "XMLHttpRequest"}
    r = s.post(url, data=payload, headers=headers, timeout=15)
    r.raise_for_status()
    data = r.json()
    return {"krw": int(round(float(data["fromAmount"]))), "fee": int(round(float(data["fromFee"])))}


# --------------------------------------------------------------------------
# Browser-driven fetchers (Playwright). Required for WireBarley (its API
# encrypts the request body client-side), for GME itself (its calculator
# is a classic ASP.NET WebForms postback, not a callable API), and for SBI
# (its domain sits behind Cloudflare, which silently blocks the plain
# requests-based call - see fetch_sbi_browser); available as a fallback for
# Coinshot if fetch_coinshot() above doesn't match the live payload.
# --------------------------------------------------------------------------

def fetch_sbi_browser(thb=BASE_THB, fixed_fee=5000, headless=True, browser=None):
    """SBI Cosmoney - the calc/amount endpoint only returns the raw
    KRW-per-THB exchange rate; the site multiplies client-side. The 5,000
    KRW fee is SBI's published fixed fee for Thailand bank transfer
    ("Remittance Limit & Transaction Fee" page) rather than something this
    endpoint returns - update `fixed_fee` if SBI changes its fee schedule.

    A plain `requests` POST to /calc/amount - even with the correct
    double-submit CSRF header (`X-XSRF-TOKEN` echoing the `XSRF-TOKEN`
    cookie) and a matching session - still gets silently served the site's
    normal HTML instead of JSON. SBI's domain runs behind Cloudflare
    (`cdn-cgi/...` requests are visible in its network traffic), which most
    likely fingerprints the TLS/HTTP client and serves non-browser requests
    a decoy response rather than an explicit block - the identical payload
    succeeds instantly from a real browser session. So this drives the real
    page instead: the Thailand option's `onclick` handler fires the
    calculator's own `/calc/amount` request with all the right headers and
    browser fingerprint, and we just read the response it gets back.

    `browser` lets a caller pass in an already-launched Playwright browser
    (see fetch_all) so multiple browser-driven fetchers can share one
    Chromium process/driver instead of each starting and stopping their own
    `sync_playwright()` - on Windows, starting `sync_playwright()` more than
    once per process is flaky (its internal asyncio loop teardown doesn't
    always finish cleanly, and the next invocation fails with "This event
    loop is already running"). If no `browser` is given, one is launched and
    closed locally, so this still works standalone.
    """
    from playwright.sync_api import sync_playwright

    def _run(b):
        page = b.new_page()
        try:
            page.goto("https://www.sbicosmoney.com/", wait_until="networkidle")
            with page.expect_response(lambda r: "/calc/amount" in r.url) as resp_info:
                page.evaluate(
                    'document.querySelector(\'.country-list a[data-country-id="THAILAND"]\').click();'
                )
            rate = float(resp_info.value.json()["exchangeRate"])
            return {"krw": int(thb * rate), "fee": fixed_fee}
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


def fetch_gme_browser(thb=BASE_THB, headless=True, browser=None):
    """GME's own live consumer calculator at online.gmeremit.com.

    This is GME's real, public-facing widget - the same page a customer
    would use - so fetching it live is a genuine live-vs-live comparison,
    not a static baseline. Confirmed via network capture that the widget
    posts back to https://online.gmeremit.com/Default.aspx as an ASP.NET
    WebForms ViewState postback (no clean JSON endpoint), so this drives the
    real page instead of calling anything directly - same situation as
    WireBarley.

    Manually walked through once as:
      1. open https://online.gmeremit.com/
      2. click the "Recipient Gets" currency dropdown (defaults to NPR)
      3. click "Select Your Country"
      4. type "Thailand"
      5. click "Thailand (THB)" (auto-selects "BANK DEPOSIT" as the payout
         method - leave it on that default)
      6. type the THB amount into the "Recipient Gets" field, tab out
      7. read the "You Send" figure - it already has
         "(Transfer Fees Included)" next to it, i.e. it's a single all-in
         KRW number, not rate+fee broken out.

    IMPORTANT - fee convention: GME's row has always been written as a
    single all-in figure with fee forced to 0 (KRW (1) = full "You Send"
    amount, Service fee (2) = 0), the same way it appears in the existing
    sheet - even though GME's own widget shows a ~5,000 KRW fee as a
    separate line. That's an explicit, confirmed choice (KRW=1,171,410,
    fee=0 was the reference reading), not an oversight - don't decompose it
    into rate+fee unless someone explicitly asks to change that convention.

    Confirmed via live DOM inspection (not just the manual walkthrough above)
    that text-based locators like `text=NPR` are ambiguous (5 elements match)
    and, worse, GME's sticky nav bar sits on top of the calculator at that
    scroll position and intercepts Playwright's pointer-based click even
    after scrolling the target into view. The fix is to skip pointer
    hit-testing entirely and drive the widget's actual DOM elements/handlers
    directly via `page.evaluate()`:
      - `#nCountry` is the currency/country picker's click target
        (`onclick="GetCountry()"`).
      - Its options live in `#toCurrUl li`, matched by text content.
      - `#recAmt` is the "Recipient Gets" input; `#numAmount` is "You Send".
    These ids are stable markup, not guessed positions, so this should be
    more robust to layout changes than the old text/role locators - but
    they're still GME-specific selectors, so re-verify with `headless=False`
    if GME reworks this page.

    `browser` lets a caller pass in an already-launched Playwright browser
    (see fetch_all) so multiple browser-driven fetchers can share one
    Chromium process/driver - see fetch_sbi_browser's docstring for why.
    """
    from playwright.sync_api import sync_playwright

    def _run(b):
        page = b.new_page()
        try:
            page.goto("https://online.gmeremit.com/", wait_until="domcontentloaded")
            page.wait_for_selector("#nCountry", timeout=15000)

            # Open the "Recipient Gets" currency/country picker (defaults NPR).
            page.evaluate("document.getElementById('nCountry').click()")
            page.wait_for_timeout(500)
            page.evaluate("""
                () => {
                    const li = [...document.querySelectorAll('#toCurrUl li')]
                        .find(el => el.textContent.includes('Thailand'));
                    li.click();
                }
            """)
            page.wait_for_timeout(500)  # payout method auto-switches to BANK DEPOSIT

            # "Recipient Gets" amount field - set directly and fire the events
            # the page's own JS listens for.
            page.evaluate(f"""
                () => {{
                    const el = document.getElementById('recAmt');
                    el.focus();
                    el.value = '{thb}';
                    el.dispatchEvent(new Event('input', {{bubbles: true}}));
                    el.dispatchEvent(new Event('change', {{bubbles: true}}));
                    el.dispatchEvent(new Event('blur', {{bubbles: true}}));
                }}
            """)
            page.wait_for_timeout(1500)

            # "You Send" (#numAmount) - the all-in KRW figure, fees already included.
            # Value is formatted like "1,171,728.00" - strip the decimal part
            # before stripping commas, or ".00" collapses into the digit string
            # and inflates the result 100x.
            krw_text = page.eval_on_selector("#numAmount", "el => el.value").split(".")[0]
            krw = int(re.sub(r"[^\d]", "", krw_text))

            # Confirmed for real (not theoretical): this page interaction can
            # "succeed" - no exception, no timeout - while #numAmount still
            # holds a leftover default/placeholder amount rather than the
            # real recalculated figure, if the amount-entry step didn't fully
            # take effect before this read. That silently produced a
            # suspiciously round 1,000,000 once, and because nothing raised,
            # the existing stale-row amber-flagging (which only fires on a
            # thrown exception) never caught it. Reject anything too far from
            # GME's own recent history instead of trusting every read that
            # merely didn't crash.
            if not price_history.is_plausible("thailand", "THB Bank Deposit", "GME", krw):
                raise ValueError(f"GME's fetched KRW ({krw:,}) is implausibly far from its recent history - likely read a stale/placeholder page value rather than the real one")

            # fee is intentionally always 0 - see docstring above.
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


def fetch_wirebarley_browser(thb=BASE_THB, standing_fee=None, headless=True, browser=None):
    """WireBarley - the calculator's API call
    (POST /api/kr/v2/exrate/KR/KRW) encrypts its request body client-side,
    so we drive the real page with Playwright instead.

    `standing_fee` defaults to whatever check_competitor_fees.py last read
    off WireBarley's published fee-schedule page (falls back to 3000 if
    that's never run - see known_fees.py). WireBarley's live quote page
    itself shows a 0 KRW fee plus a -3,000 KRW "signup coupon" discount,
    i.e. 3,000 KRW was the real standing fee once the one-time new-user
    coupon is gone - confirmed still accurate via live capture, and now
    re-verified daily against the published schedule instead of assumed
    fixed forever.

    Confirmed via live DOM inspection that which currency defaults to
    "sending" vs "receiving" depends on the visitor's apparent region (this
    machine's egress got sending=KRW/receiving=USD; another session got the
    opposite) - so matching selectors by their *current* currency text (the
    original code's `text=USD` / this rewrite's first attempt at `text=KRW`)
    is unreliable. Fixed by anchoring on the stable Korean row labels
    ("보내는 금액" / "받는 금액") instead and walking up to each row's
    `cursor-pointer` currency button, which works regardless of which
    currency happens to be selected by default.

    Also confirmed the amount fields render as a `<label>` inside a
    `<button>` until clicked once, at which point they become real
    `<input>`s - a `text=...` + `.type()` locator can't type into a label,
    which is why the original code timed out waiting for a second `<input>`
    to exist. The natural-looking fix - click *any* amount button once to
    reveal both inputs, then separately focus/click into the receiving
    input - turned out to race WireBarley's own blur handling: focusing (or
    even Playwright-clicking) the revealed receiving input after the fact
    collapses both inputs straight back to buttons before the value can be
    set. Clicking the *receiving* amount button directly - the one you
    actually want to edit, in a single step - sidesteps that race entirely
    and leaves it focused and editable.

    Fixed sequence, all driven via `page.evaluate()` (rather than
    Playwright's pointer-based `.click()`) since it reads the DOM directly
    instead of relying on ambiguous text matches or on-screen hit-testing:
      1. find the row containing "받는 금액" (receiving), click its currency
         button -> click the "태국" (Thailand) option in the dropdown.
      2. find the row containing "보내는 금액" (sending), click its currency
         button -> click the "KRW" option in the dropdown.
      3. click the *receiving* amount button (second visible
         `button.w-full.text-left`) directly - this both reveals and
         focuses its `<input>` in one step.
      4. set that `<input>` (THB) via the native value setter (needed
         because this is a React-controlled input - assigning `.value`
         directly doesn't trigger React's change detection) and dispatch
         `input` so React recalculates the sending side.
      5. read the sending `<input>` (KRW) value.
    Still WireBarley-specific selectors reconstructed from a live run, not
    stable ids - re-verify with `headless=False` if WireBarley reworks this
    page.

    `browser` lets a caller pass in an already-launched Playwright browser
    (see fetch_all) so multiple browser-driven fetchers can share one
    Chromium process/driver - see fetch_sbi_browser's docstring for why.
    """
    if standing_fee is None:
        standing_fee = known_fees.get_fee("WireBarley", default=3000)

    from playwright.sync_api import sync_playwright

    open_selector_by_label = """
        (labelText) => {
            for (const el of document.querySelectorAll('*')) {
                if (el.children.length === 0 && el.textContent.trim() === labelText) {
                    let row = el.closest('div');
                    for (let i = 0; i < 8 && row; i++) {
                        const btn = row.querySelector('[class*="cursor-pointer"]');
                        if (btn) { btn.click(); return true; }
                        row = row.parentElement;
                    }
                }
            }
            return false;
        }
    """
    click_dropdown_option = """
        (optionText) => {
            for (const el of document.querySelectorAll('*')) {
                if (el.children.length === 0 && el.textContent.trim() === optionText) {
                    const btn = el.closest('button');
                    if (btn) { btn.click(); return true; }
                }
            }
            return false;
        }
    """

    def _run(b):
        page = b.new_page()
        try:
            page.goto("https://www.wirebarley.com/ko", wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            # Ensure sending country is Korea (KR) if defaulted to US or other region
            header_text = page.eval_on_selector("header", "el => el.textContent") or ""
            if "US" in header_text or "KR" not in header_text:
                page.evaluate("""() => {
                    const usBtn = [...document.querySelectorAll('header button')].find(b => b.textContent.includes('US'));
                    if (usBtn) usBtn.click();
                }""")
                page.wait_for_timeout(500)
                page.evaluate("""() => {
                    const tab = [...document.querySelectorAll('button, div, span')].find(el => el.textContent.trim() === '송금국가');
                    if (tab) tab.click();
                }""")
                page.wait_for_timeout(500)
                page.evaluate("""() => {
                    const btn = [...document.querySelectorAll('li button')].find(el => el.textContent.includes('대한민국'));
                    if (btn) btn.click();
                }""")
                page.wait_for_timeout(1500)

            # Switch the "받는 금액" (receiving) row to Thailand/THB.
            page.evaluate(open_selector_by_label, "받는 금액")
            page.wait_for_timeout(500)
            page.evaluate(click_dropdown_option, "태국")
            page.wait_for_timeout(1000)

            # Click the receiving amount button directly - reveals AND
            # focuses its <input> in one step, avoiding the blur race.
            page.evaluate("""
                () => {
                    const btns = [...document.querySelectorAll('button.w-full.text-left')]
                        .filter(b => b.getBoundingClientRect().width > 0);
                    if (btns[1]) btns[1].click();
                }
            """)
            page.wait_for_timeout(500)

            # inputs[0] = sending (KRW), inputs[1] = receiving (THB) - among
            # non-checkbox inputs specifically. Confirmed live (2026-09-10)
            # that WireBarley added a real <input type="checkbox"> elsewhere
            # on the page (a plain document.querySelectorAll('input') picks
            # it up as element 0), which had silently shifted these indices
            # by one: the "receiving" write was landing on the *sending*
            # KRW field instead, and the later read of "sending" was
            # actually reading the checkbox's value ('on' - zero digits),
            # which is the real explanation for the "invalid literal for
            # int(): ''" failures on every run since ~2026-09-09, not a
            # timing issue at all (confirmed via a local diagnostic run
            # against the live page, not guessed) - excluding checkboxes
            # restores the original, correct indexing.
            page.evaluate(f"""
                () => {{
                    const inputs = document.querySelectorAll('input:not([type="checkbox"])');
                    const receivingInput = inputs[1] || inputs[0];
                    if (receivingInput) {{
                        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                        setter.call(receivingInput, '{thb}');
                        receivingInput.dispatchEvent(new Event('input', {{bubbles: true}}));
                        receivingInput.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                }}
            """)
            # Kept as a real (if now secondary) safety net: still polls for
            # a non-empty, digit-bearing value up to 10s rather than trusting
            # one fixed delay, in case recalculation is genuinely slow on a
            # given run - but the fix above (correct element, not more
            # patience) is what actually resolves the current failures.
            send_amount_text = None
            for _ in range(20):
                page.wait_for_timeout(500)
                values = page.eval_on_selector_all("input:not([type='checkbox'])", "els => els.map(e => e.value)")
                if values and re.sub(r"[^\d]", "", values[0]):
                    send_amount_text = values
                    break
            if not send_amount_text:
                raise ValueError("WireBarley: amount input still empty/unreadable after 10s of polling")

            krw = int(re.sub(r"[^\d]", "", send_amount_text[0]))

            # Validate against recent historical median before trusting the reading
            if not price_history.is_plausible("thailand", "THB Bank Deposit", "WireBarley", krw):
                raise ValueError(
                    f"WireBarley's fetched KRW ({krw:,}) is implausibly far from recent history - "
                    f"likely read default/placeholder currency value"
                )

            return {"krw": krw, "fee": standing_fee}
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


def fetch_coinshot_browser(thb=BASE_THB, standing_fee=2500, headless=True):
    """Fallback for Coinshot if fetch_coinshot()'s API guess doesn't match
    the live payload - drives the real page instead."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()
        page.goto("https://coinshot.org/main", wait_until="networkidle")

        amount_inputs = page.locator("input")
        receiving_input = amount_inputs.nth(1)  # THB receiving field
        receiving_input.click()
        receiving_input.fill("")
        receiving_input.type(str(thb))
        page.keyboard.press("Tab")
        page.wait_for_timeout(1500)

        send_amount_text = amount_inputs.nth(0).input_value()
        krw = int(re.sub(r"[^\d]", "", send_amount_text))

        browser.close()
        return {"krw": krw, "fee": standing_fee}


def fetch_kebhana(thb=BASE_THB, browser=None):
    """KEB Hana Bank - see fetch_kebhana_rates.py. That module owns the
    actual page-scraping logic (it's shared with fetch_laos_rates.py's USD
    corridor too); this is just the {"krw":..,"fee":..} adapter fetch_all()
    expects, using THB's "송금 보낼 때" (T/T selling) rate and a flat 5,000 KRW
    fee, same convention as every other non-GME provider here."""
    import fetch_kebhana_rates

    return fetch_kebhana_rates.fetch_kebhana_quote("THB", thb, browser=browser)


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

FETCHERS = {
    "Cross": fetch_cross,
    "JRF": fetch_jrf,
    "E9Pay": fetch_e9pay,
    "GmoneyTrans": fetch_gmoneytrans,
    "Coinshot": fetch_coinshot,
    "Hanpass": fetch_hanpass,
}


def fetch_all(thb=BASE_THB, skip_wirebarley=False, skip_gme=False, skip_sbi=False, skip_kebhana=False, gme_manual=None):
    """Fetch quotes for every provider, GME included. Returns
    {provider: {"krw":.., "fee":..}}. Providers whose fetch raises an
    exception are omitted (and reported), so the caller can leave those
    cells untouched rather than write bad data - this applies to GME too:
    if its live fetch fails, GME's existing cell is left alone rather than
    guessed at.

    `gme_manual`: a KRW figure to use for GME's row directly instead of
    live-fetching (e.g. GME's own site being down, confirmed 2026-08-10 -
    both the live web fetch and manually browsing it were unreachable at the
    time). Written as fresh, non-stale data - takes priority over skip_gme,
    and skips the live fetch entirely rather than trying it first.
    """
    results = {}
    errors = {}
    session = requests.Session()

    if gme_manual is not None:
        # Sanity bound, not a rate check - just wide enough (roughly
        # 35-55 KRW/THB) to catch an obvious typo like a missing or extra
        # digit before it goes out looking like a real fetched rate.
        lo, hi = round(thb * 35), round(thb * 55)
        if not (lo <= gme_manual <= hi):
            raise ValueError(
                f"Manual GME KRW ({gme_manual:,}) is outside the plausible range "
                f"({lo:,}-{hi:,}) for {thb:,} THB - check for a typo (e.g. a "
                "missing digit) before retrying"
            )

    # GME, SBI and WireBarley all need Playwright. They share a single
    # sync_playwright()/browser instance rather than each launching and
    # closing their own - on Windows, starting sync_playwright() more than
    # once per process is flaky (its asyncio loop teardown doesn't always
    # finish cleanly before the next launch), which surfaced as "This event
    # loop is already running" / "Sync API inside the asyncio loop" once a
    # third browser-driven fetcher (SBI) was added.
    need_gme_browser = not skip_gme and gme_manual is None
    need_browser = not (not need_gme_browser and skip_sbi and skip_wirebarley and skip_kebhana)
    browser_ctx = playwright_instance = None
    if need_browser:
        from playwright.sync_api import sync_playwright

        playwright_instance = sync_playwright().start()
        browser_ctx = playwright_instance.chromium.launch(headless=True)

    try:
        if gme_manual is not None:
            results["GME"] = {"krw": gme_manual, "fee": 0}
            print(f"Using manually entered GME rate (site unreachable) ... KRW={gme_manual:,}  fee=0")
        elif not skip_gme:
            try:
                print("Fetching GME (browser, own live rate) ...", end=" ", flush=True)
                results["GME"] = fetch_gme_browser(thb, browser=browser_ctx)
                print(f"KRW={results['GME']['krw']:,}  fee={results['GME']['fee']:,}")
            except Exception as e:  # noqa: BLE001
                errors["GME"] = e
                print(f"FAILED ({e})")

        if not skip_sbi:
            try:
                print("Fetching SBI (browser, Cloudflare blocks plain requests) ...", end=" ", flush=True)
                results["SBI"] = fetch_sbi_browser(thb, browser=browser_ctx)
                print(f"KRW={results['SBI']['krw']:,}  fee={results['SBI']['fee']:,}")
            except Exception as e:  # noqa: BLE001
                errors["SBI"] = e
                print(f"FAILED ({e})")

        if not skip_kebhana:
            try:
                print("Fetching KEB Hana (browser, live rate page) ...", end=" ", flush=True)
                results["KEB Hana"] = fetch_kebhana(thb, browser=browser_ctx)
                print(f"KRW={results['KEB Hana']['krw']:,}  fee={results['KEB Hana']['fee']:,}")
            except Exception as e:  # noqa: BLE001
                errors["KEB Hana"] = e
                print(f"FAILED ({e})")

        for name, fn in FETCHERS.items():
            try:
                print(f"Fetching {name} ...", end=" ", flush=True)
                results[name] = fn(thb, session=session)
                print(f"KRW={results[name]['krw']:,}  fee={results[name]['fee']:,}")
            except Exception as e:  # noqa: BLE001 - report and continue
                errors[name] = e
                print(f"FAILED ({e})")

        if not skip_wirebarley:
            try:
                print("Fetching WireBarley (browser) ...", end=" ", flush=True)
                results["WireBarley"] = fetch_wirebarley_browser(thb, browser=browser_ctx)
                print(f"KRW={results['WireBarley']['krw']:,}  fee={results['WireBarley']['fee']:,}")
            except Exception as e:  # noqa: BLE001
                errors["WireBarley"] = e
                print(f"FAILED ({e})")
    finally:
        if browser_ctx is not None:
            browser_ctx.close()
        if playwright_instance is not None:
            playwright_instance.stop()

    if errors:
        print("\nProviders that failed and were left untouched in the sheet:")
        for name, e in errors.items():
            print(f"  - {name}: {e}")

    return results


def stale_rows_from(results):
    """Rows this script tries to fetch every run, but whose provider failed
    to fetch this time - see STALE_FILL in write_to_workbook."""
    return {row for provider, row in INPUT_AREA_ROWS.items() if provider not in results}


def write_to_workbook(results, stale_rows, input_path, output_path, sheet_name=SHEET_NAME):
    wb = openpyxl.load_workbook(input_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        sys.exit(f"Sheet '{sheet_name}' not found. Sheets present: {wb.sheetnames}")
    ws = wb[sheet_name]

    for provider, row in INPUT_AREA_ROWS.items():
        if provider not in results:
            continue
        ws.cell(row=row, column=13, value=results[provider]["krw"])   # column M
        ws.cell(row=row, column=14, value=results[provider]["fee"])   # column N
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
        "left-hand table will re-rank itself by Total price automatically."
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
    parser.add_argument("--thb", type=int, default=BASE_THB, help="Base THB amount (default 26000)")
    parser.add_argument("--dry-run", action="store_true", help="Fetch and print only; don't write the file")
    parser.add_argument("--skip-wirebarley", action="store_true", help="Skip the Playwright/WireBarley step")
    parser.add_argument("--skip-gme", action="store_true", help="Skip GME's own live fetch (leaves its existing cell untouched)")
    parser.add_argument("--skip-sbi", action="store_true", help="Skip the Playwright/SBI step")
    parser.add_argument("--skip-kebhana", action="store_true", help="Skip the Playwright/KEB Hana step")
    parser.add_argument(
        "--gme-manual", type=int, default=None,
        help="Use this KRW figure for GME instead of live-fetching (e.g. GME's site is down) - written as fresh data, not flagged stale",
    )
    parser.add_argument(
        "--log-history", action="store_true",
        help="Append this run's prices/gaps to price_history.db (opt-in - only the scheduled automation should pass this, not ad-hoc/test runs)",
    )
    args = parser.parse_args()

    output_path = args.output or args.input

    results = fetch_all(
        thb=args.thb,
        skip_wirebarley=args.skip_wirebarley,
        skip_gme=args.skip_gme,
        skip_sbi=args.skip_sbi,
        skip_kebhana=args.skip_kebhana,
        gme_manual=args.gme_manual,
    )

    print("\n--- Summary (sorted by Total price, descending) ---")
    rows = sorted(results.items(), key=lambda kv: kv[1]["krw"] + kv[1]["fee"], reverse=True)
    for name, vals in rows:
        total = vals["krw"] + vals["fee"]
        print(f"{name:<12} KRW={vals['krw']:>10,}  fee={vals['fee']:>7,}  total={total:>10,}")

    if args.dry_run:
        print("\n(dry run - workbook not modified)")
        return

    stale_rows = stale_rows_from(results)
    write_to_workbook(results, stale_rows, args.input, output_path)

    if args.log_history:
        entries = [
            {"corridor": "THB Bank Deposit", "provider": provider, "krw": vals["krw"], "fee": vals["fee"]}
            for provider, vals in results.items()
        ]
        n = price_history.log_run("thailand", entries)
        print(f"Logged {n} rows to {price_history.DEFAULT_DB_PATH}")


if __name__ == "__main__":
    main()
