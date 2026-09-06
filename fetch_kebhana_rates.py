#!/usr/bin/env python3
"""
fetch_kebhana_rates.py

PROTOTYPE - standalone, not yet wired into fetch_thailand_rates.py or the
dashboard. Fetches KEB Hana Bank's live "송금 보낼 때" (telegraphic-transfer /
T/T selling) rate for a given currency and computes what it would cost to
send THB/USD via KEB Hana, for comparison against the other providers already
tracked in fetch_thailand_rates.py.

Source: https://www.hanabank.com/cont/mall/mall15/mall1501/index.jsp
The page embeds an iframe (id="bankIframe") pointing at
/cms/rate/wpfxd651_01i.do, which server-renders the full day's rate table for
every currency in one go - no separate AJAX call or params needed once the
iframe itself is loaded. Table columns (per row): currency, 현찰 사실때,
spread, 현찰 파실때, spread, 송금 보낼 때, 송금 받을 때, 외화수표 파실때,
매매기준율, 환가료율, 미화환산율.

"송금 보낼 때" (6th <td> per row, index 5 - counting the two Spread cells that
sit between 현찰/송금 pairs) is the rate the bank charges when you hand over
KRW to send foreign currency out - the one comparable to every other
provider's remittance rate. Confirmed against user-supplied reference values
for 2026-08-09 (USD 1,424.80, THB 43.11): live fetch returned 1,425.80 / 43.14
- within a few hundredths, consistent with normal intraday rate movement
between when the reference was read and when this ran.

Fee: added separately (not baked into the rate), same convention as most
other providers in the sheet. Originally assumed flat 5,000 KRW; confirmed
via check_competitor_fees.py against KEB Hana's real fee-schedule page that
the correct tier for our amounts (well under the $5,000-equivalent
electronic-banking threshold) is actually 3,000 KRW - see known_fees.json /
known_fees.py, which fetch_kebhana_quote() now reads instead of a hardcoded
constant.

Usage:
    python fetch_kebhana_rates.py
"""

import sys

try:
    from playwright.sync_api import sync_playwright
except ImportError:  # pragma: no cover
    sys.exit("Missing dependency: pip install playwright && playwright install chromium")

import known_fees

RATE_PAGE_URL = "https://www.hanabank.com/cont/mall/mall15/mall1501/index.jsp"
FIXED_FEE = 5000  # fallback only - see fetch_kebhana_quote, which reads known_fees.json first

# currency code -> (Korean name substring as shown in the table, base amount for comparison)
CURRENCIES = {
    "THB": ("태국", 26000),
    "USD": ("미국", 1000),
}


def fetch_kebhana_rates(headless=True, browser=None):
    """Returns {currency_code: send_rate} for every currency in CURRENCIES,
    read live from KEB Hana's rate table."""
    own_playwright = browser is None
    pw = sync_playwright().start() if own_playwright else None
    b = browser or pw.chromium.launch(headless=headless)
    try:
        page = b.new_page()
        page.goto(RATE_PAGE_URL, wait_until="domcontentloaded", timeout=20000)

        # The rate table renders into the iframe after its own JS runs (an
        # in-place content swap that can replace the frame document, so a
        # frame handle grabbed too early can go stale) - re-look up the
        # frame by name each poll rather than caching one handle.
        rate_table = None
        for _ in range(30):
            frame = page.frame(name="bankIframe")
            if frame is not None:
                for t in frame.query_selector_all("table"):
                    if "미국 USD" in t.inner_text():
                        rate_table = t
                        break
            if rate_table is not None:
                break
            page.wait_for_timeout(500)
        if rate_table is None:
            raise ValueError("Could not find the rate table on KEB Hana's page")

        rows = rate_table.query_selector_all("tr")
        results = {}
        for row in rows:
            cells = [c.inner_text().strip() for c in row.query_selector_all("td")]
            if len(cells) < 6:
                continue
            name_cell = cells[0]
            for code, (kr_name, _) in CURRENCIES.items():
                if kr_name in name_cell:
                    send_rate_text = cells[5].replace(",", "")
                    try:
                        results[code] = float(send_rate_text)
                    except ValueError:
                        pass
        return results
    finally:
        if browser is None:
            b.close()
        if own_playwright:
            pw.stop()


def compute_quote(rate, amount, fee=FIXED_FEE):
    krw = round(rate * amount)
    return {"rate": rate, "amount": amount, "krw": krw, "fee": fee, "total": krw + fee}


def fetch_kebhana_quote(currency, amount, fee=None, browser=None):
    """Live rate + computed quote for one currency, in the {"krw":..,"fee":..}
    shape fetch_thailand_rates.py / fetch_laos_rates.py expect from every
    provider fetcher. `browser` lets a caller pass an already-launched
    Playwright browser to share across fetchers - see fetch_gme_browser's
    docstring in fetch_thailand_rates.py for why that matters on Windows;
    if omitted, one is launched and closed just for this call.

    `fee` defaults to whatever check_competitor_fees.py last read off KEB
    Hana's own fee-schedule page (falls back to FIXED_FEE if that's never
    run) - see known_fees.py."""
    import price_history

    if fee is None:
        fee = known_fees.get_fee("KEB Hana", default=FIXED_FEE)
    rates = fetch_kebhana_rates(browser=browser)
    if currency not in rates:
        raise ValueError(f"KEB Hana: could not read the {currency} rate from the live page")
    quote = compute_quote(rates[currency], amount, fee=fee)
    krw = quote["krw"]
    report = "thailand" if currency == "THB" else "laos"
    corridor = "THB Bank Deposit" if currency == "THB" else f"{currency} Bank Deposit"
    if not price_history.is_plausible(report, corridor, "KEB Hana", krw):
        raise ValueError(f"KEB Hana: fetched KRW ({krw:,}) is implausibly far from recent history")
    return {"krw": krw, "fee": quote["fee"]}


if __name__ == "__main__":
    rates = fetch_kebhana_rates()

    reference = {"THB": 43.11, "USD": 1424.80}

    for code, (_, amount) in CURRENCIES.items():
        if code not in rates:
            print(f"{code}: FAILED to find a rate on the page")
            continue
        rate = rates[code]
        quote = compute_quote(rate, amount)
        ref = reference.get(code)
        ref_note = f" (reference: {ref})" if ref is not None else ""
        print(f"{code} send rate: {rate}{ref_note}")
        print(
            f"  {amount:,} {code} -> KRW={quote['krw']:,}  fee={quote['fee']:,}  "
            f"total={quote['total']:,}"
        )
