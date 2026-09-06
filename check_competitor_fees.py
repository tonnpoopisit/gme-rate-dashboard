#!/usr/bin/env python3
"""
check_competitor_fees.py

Daily (18:00, separate from the 2-hourly rate schedule - fee schedules
change far less often than rates) check of the three provider fees that
can't be read live from the same calculator response used for their rate:

    WireBarley (Thailand)      - live quote page shows 0 KRW fee at query
                                  time; real fee is a separate published
                                  tiered schedule (help.wirebarley.com).
    KEB Hana (Thailand + Laos) - the rate page we scrape has no fee field
                                  at all; real fee is a separate tiered
                                  schedule (hanabank.com fee-schedule page).
    E9Pay (Thailand + Laos)    - E9Pay's calculator API's REMIT_FEE field
                                  always reports 0, but E9Pay's own
                                  homepage widget displays the real fee
                                  directly as page text.

Everyone else's fee (Cross, GmoneyTrans, Hanpass, Coinshot) already comes
straight from their calculator response on every 2-hourly run - already
correct going forward, nothing to check here.

Each check function reads the provider's real fee-disclosure page (not
their quote calculator) for our actual comparison amounts and returns the
fee in KRW. main() compares each against known_fees.get_fee() and calls
known_fees.set_fee() when it differs, printing a CHANGED line the
PowerShell orchestrator (check_competitor_fees.ps1) greps for to decide
whether to fire a desktop notification.

Entirely read-only against every external site - safe to run directly,
any time, no Teams webhook involved.

Usage:
    python check_competitor_fees.py
"""

import json
import re
import sys

import requests

import known_fees

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

# Our actual comparison amounts (see BASE_THB in fetch_thailand_rates.py,
# BASE_USD in fetch_laos_rates.py) - both always land in the lowest tier of
# every tiered schedule seen so far, but each checker still does a real
# tier lookup rather than assuming that stays true forever.
THAILAND_KRW_EQUIVALENT = 1_120_000  # ~26,000 THB
LAOS_USD_EQUIVALENT = 1_000  # USD corridor base


def check_wirebarley_fee(krw_amount=THAILAND_KRW_EQUIVALENT):
    """WireBarley's Korea remittance fee schedule is published as clean
    JSON-LD (schema.org FAQPage) on a stable help-center article - far more
    reliable to parse than scraping the rendered HTML table. Tiered by KRW
    amount under the SWIFT SHA fee type (the shared-cost option, matching
    our current 3,000 KRW assumption - not SWIFT OUR, a different/pricier
    option the page also lists)."""
    url = (
        "https://help.wirebarley.com/en/support/solutions/articles/"
        "156000320230--south-korea-what-are-the-remittance-limits-and-transfer-fees-"
    )
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.raise_for_status()

    m = re.search(r'<script type="application/ld\+json">\s*(\{.*?\})\s*</script>', r.text, re.DOTALL)
    if not m:
        raise ValueError("WireBarley: could not find the JSON-LD fee block on the help article")
    data = json.loads(m.group(1))
    text = data["mainEntity"][0]["acceptedAnswer"]["text"]

    sha_section_match = re.search(r"Transfer fee\(SWIFT SHA\):(.*?)Transfer fee\(SWIFT OUR\)", text, re.DOTALL)
    if not sha_section_match:
        raise ValueError("WireBarley: could not find the SWIFT SHA fee section in the fetched text")
    sha_text = sha_section_match.group(1)

    tiers = []
    for m in re.finditer(r"Less than ([\d,]+) won:\s*([\d,]+)\s*won", sha_text):
        tiers.append((int(m.group(1).replace(",", "")), int(m.group(2).replace(",", ""))))
    for m in re.finditer(r"([\d,]+) won or less than ([\d,]+) won:\s*([\d,]+)\s*won", sha_text):
        tiers.append((int(m.group(2).replace(",", "")), int(m.group(3).replace(",", ""))))
    if not tiers:
        raise ValueError(f"WireBarley: could not parse any fee tiers from: {sha_text!r}")
    tiers.sort()

    for threshold, fee in tiers:
        if krw_amount < threshold:
            return fee
    # At/above every "less than" threshold found - the schedule's own text
    # says the top bracket is free, but don't assume that stays true
    # forever without the top-bracket text actually saying so.
    if "or more: Free" in sha_text:
        return 0
    raise ValueError(f"WireBarley: {krw_amount:,} KRW doesn't fall under any parsed tier: {tiers}")


def check_kebhana_fee(usd_equivalent=LAOS_USD_EQUIVALENT):
    """KEB Hana's fee schedule page (hanabank.com) has two separate tables
    - branch-counter (창구) and electronic-banking (전자금융/EDI 제외). Every
    provider in this comparison is an online/app service, so this reads
    the electronic-banking table, not the (pricier) branch-counter one -
    see the module note on fetch_kebhana_rates.py's flat-5,000 guess
    likely already being the wrong tier."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        page = b.new_page()
        page.goto("https://www.hanabank.com/cont/mall/mall09/mall0906/mall090603/index.jsp", timeout=30000)
        page.wait_for_timeout(2000)

        target_table = None
        for t in page.query_selector_all("table"):
            txt = t.inner_text()
            if "전자금융" in txt or ("국외 외화송금" in txt and "이하" in txt):
                # Prefer the electronic-banking table specifically; the
                # branch-counter table also contains "국외 외화송금" so check
                # the more specific marker first, else fall back only if
                # nothing better is found.
                if "전자금융" in txt:
                    target_table = t
                    break
                elif target_table is None:
                    target_table = t

        if target_table is None:
            b.close()
            raise ValueError("KEB Hana: could not find a fee table on the fee-schedule page")

        rows_text = target_table.inner_text()
        b.close()

    # Rows look like "국외 외화송금	USD 5,000불상당액 이하	3,000원" followed by
    # "USD 5,000불상당액 초과	5,000원" on the next line.
    tiers = []
    for line in rows_text.splitlines():
        m = re.search(r"USD\s*([\d,]+)불?\s*상당액?\s*(이하|초과).*?([\d,]+)\s*원", line)
        if m:
            threshold = int(m.group(1).replace(",", ""))
            is_under = m.group(2) == "이하"
            fee = int(m.group(3).replace(",", ""))
            tiers.append((threshold, is_under, fee))

    if not tiers:
        raise ValueError(f"KEB Hana: could not parse any fee tiers from the electronic-banking table: {rows_text!r}")

    for threshold, is_under, fee in tiers:
        if is_under and usd_equivalent <= threshold:
            return fee
    for threshold, is_under, fee in tiers:
        if not is_under and usd_equivalent > threshold:
            return fee
    raise ValueError(f"KEB Hana: ${usd_equivalent:,} equivalent doesn't fall under any parsed tier: {tiers}")


def check_e9pay_fee():
    """E9Pay's calculator widget lives right on their homepage and shows
    the real fee as plain page text ("송금 수수료 : 5,000 KRW"), unlike the
    calcExchangeRate.do API's REMIT_FEE field (always 0 - see
    fetch_thailand_rates.py's fetch_e9pay docstring). Confirmed live this
    fee is flat/sitewide regardless of destination country (the homepage's
    default corridor - VND - showed the same figure our THB/LAK/USD
    corridors already assume), so this doesn't need to switch corridors on
    the widget, just read whatever's showing."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        page = b.new_page()
        page.goto("https://www.e9pay.co.kr/", timeout=30000)
        page.wait_for_timeout(2000)
        text = page.inner_text("body")
        b.close()

    m = re.search(r"송금\s*수수료\s*:\s*([\d,]+)\s*KRW", text)
    if not m:
        raise ValueError("E9Pay: could not find the displayed fee text on the homepage")
    return int(m.group(1).replace(",", ""))


CHECKS = {
    "WireBarley": (check_wirebarley_fee, "https://help.wirebarley.com/en/support/solutions/articles/156000320230"),
    "KEB Hana": (check_kebhana_fee, "https://www.hanabank.com/cont/mall/mall09/mall0906/mall090603/index.jsp"),
    "E9Pay": (check_e9pay_fee, "https://www.e9pay.co.kr/"),
}


def run_checks():
    results = {}
    for provider, (fn, source_url) in CHECKS.items():
        try:
            new_fee = fn()
            old_fee = known_fees.get_fee(provider, default=None)
            changed = (old_fee is not None and new_fee != old_fee)
            known_fees.set_fee(provider, new_fee, source_url)
            results[provider] = {"ok": True, "fee": new_fee, "oldFee": old_fee, "changed": changed, "source": source_url}
        except Exception as e:  # noqa: BLE001
            results[provider] = {"ok": False, "error": str(e), "source": source_url}
    return results


def main():
    results = run_checks()
    any_error = False
    for provider, res in results.items():
        if not res.get("ok"):
            print(f"{provider}: FAILED to check ({res.get('error')})")
            any_error = True
        elif res.get("changed"):
            print(f"{provider}: CHANGED: {res['oldFee']:,} -> {res['fee']:,} KRW (source: {res['source']})")
        else:
            print(f"{provider}: unchanged ({res['fee']:,} KRW)")

    if any_error:
        sys.exit(1)


if __name__ == "__main__":
    main()
