#!/usr/bin/env python3
"""
excel_table.py

Builds the {"rows": [...]} JSON shape render_table_png.py already consumes,
matching the real Excel workbook's ranking table exactly (per-provider
highlight colors, descending-by-total-price sort, Country/Service merged
labels, per-block Date+header rows) - based on a real reference screenshot
of the workbook, not guessed. See ranking.py for the underlying data.

Sort order note: Excel's own ranking formulas use LARGE(), which ranks
descending - most expensive at the top, GME/cheapest deals at the bottom.
This is the opposite of the dashboard's own live bar-chart view (which sorts
ascending, cheapest first, for a "best deal first" glance) - both are
correct for their own context, just different conventions.
"""
from datetime import datetime
from zoneinfo import ZoneInfo

from ranking import CORRIDOR_BASE_AMOUNTS

KST = ZoneInfo("Asia/Seoul")

HEADER_BG = "#2E75B6"
HEADER_FG = "#FFFFFF"
BLUE_BG = "#2E75B6"        # Cross, JRF, WireBarley, Coinshot, KEB Hana - the "generic" providers
GME_BG = "#F4777F"
GREEN_BG = "#92D050"       # GmoneyTrans
YELLOW_BG = "#FFFF00"      # E9Pay
LIGHTBLUE_BG = "#BDD7EE"   # Hanpass
WHITE = "#FFFFFF"
BLACK = "#000000"
NEG_GAP_COLOR = "#C00000"

# Only one corridor in this whole pipeline ever has two simultaneous GME
# rows (Laos's LAK Bank Deposit: RIA-Other-banks + Moneygram-BCEL). The
# sheet's own Price Gap formulas are anchored to Moneygram-BCEL there
# (confirmed via the workbook's own Q7 reference cell - see
# fetch_laos_rates.py's --log-history entries), a fixed business choice, not
# "whichever is cheaper today" - so it's hardcoded here to match, not derived.
BASELINE_OVERRIDES = {"LAK Bank Deposit": "GME (Moneygram-BCEL)"}


def _provider_style(provider: str):
    """(background, text color, bold) for a provider's Competitor-name cell -
    matches the fixed per-provider color convention in the real workbook,
    not a rank- or corridor-based color."""
    if provider.startswith("GME"):
        return GME_BG, BLACK, True
    if provider.startswith("GmoneyTrans"):
        return GREEN_BG, BLACK, False
    if provider.startswith("E9Pay"):
        return YELLOW_BG, BLACK, False
    if provider.startswith("Hanpass"):
        return LIGHTBLUE_BG, BLACK, False
    return BLUE_BG, WHITE, True  # Cross, JRF, WireBarley, Coinshot, KEB Hana


def _cell(text, bg=WHITE, fg=BLACK, bold=False, align="center", **extra):
    return {"text": text, "bg": bg, "fg": fg, "bold": bold, "align": align, **extra}


def _country_service_currency(report: str, corridor: str):
    country = "Thailand" if report == "thailand" else "Laos"
    currency, _, service = corridor.partition(" ")
    return country, service, currency


def _pick_baseline(corridor: str, providers: list):
    gme_rows = [p for p in providers if p["isGme"]]
    if not gme_rows:
        return None
    if len(gme_rows) == 1:
        return gme_rows[0]
    override = BASELINE_OVERRIDES.get(corridor)
    for p in gme_rows:
        if p["provider"] == override:
            return p
    return gme_rows[0]


def build_table_rows(report: str, latest: dict, stale_pairs: set = None) -> dict:
    """latest: ranking.latest_rates(report, db_path) output. stale_pairs:
    set of (corridor, provider) whose fetch failed this run (gets the same
    amber border render_table_png.py already supports)."""
    stale_pairs = stale_pairs or set()
    rows = []

    for i, (corridor, block) in enumerate(latest.items()):
        country, service, currency = _country_service_currency(report, corridor)
        providers = block["providers"]

        baseline = _pick_baseline(corridor, providers)
        baseline_total = baseline["total"] if baseline else None

        ranked = sorted(providers, key=lambda p: p["total"], reverse=True)

        if i > 0:
            rows.append([_cell("") for _ in range(8)])

        # Only the first block gets a "Date:" row - one timestamp for the
        # whole table (all corridors are fetched in the same pipeline run,
        # seconds apart) rather than repeating it before every block.
        if i == 0:
            as_of = block.get("asOf")
            date_str = ""
            if as_of:
                try:
                    date_str = datetime.fromisoformat(as_of).astimezone(KST).strftime("%Y-%m-%d %H:%M:%S")
                except ValueError:
                    date_str = as_of
            rows.append(
                [_cell("") for _ in range(6)]
                + [_cell("Date:", bold=True, align="right"), _cell(date_str, align="right")]
            )
        rows.append(
            [
                _cell("Country", bg=HEADER_BG, fg=HEADER_FG, bold=True),
                _cell("Service", bg=HEADER_BG, fg=HEADER_FG, bold=True),
                _cell("Competitor", bg=HEADER_BG, fg=HEADER_FG, bold=True),
                _cell(f"FCY({currency})", bg=HEADER_BG, fg=HEADER_FG, bold=True),
                _cell("KRW①", bg=HEADER_BG, fg=HEADER_FG, bold=True),
                _cell("Service fee②", bg=HEADER_BG, fg=HEADER_FG, bold=True),
                _cell("Total price①+②", bg=HEADER_BG, fg=HEADER_FG, bold=True),
                _cell("Price gap", bg=HEADER_BG, fg=HEADER_FG, bold=True),
            ]
        )

        fcy_amount = CORRIDOR_BASE_AMOUNTS.get(currency)
        n = len(ranked)
        for idx, p in enumerate(ranked):
            bg, fg, bold = _provider_style(p["provider"])
            is_baseline = baseline is not None and p is baseline
            gap = None
            if baseline_total is not None:
                gap = 0 if is_baseline else p["total"] - baseline_total
            neg = gap is not None and gap < 0
            gap_fg = NEG_GAP_COLOR if neg else BLACK
            gap_bold = neg or is_baseline

            row = [
                _cell(country if idx == 0 else "", rowspan=n, skip=(idx != 0)),
                _cell(service if idx == 0 else "", rowspan=n, skip=(idx != 0)),
                _cell(p["provider"], bg=bg, fg=fg, bold=bold),
                _cell(f"{fcy_amount:,}" if fcy_amount else "", align="right"),
                _cell(f"{p['krw']:,}", align="right", bold=is_baseline),
                _cell(f"{p['fee']:,}", align="right", bold=is_baseline),
                _cell(f"{p['total']:,}", align="right", bold=is_baseline),
                _cell(f"{gap:,}" if gap is not None else "", align="right", bold=gap_bold, fg=gap_fg),
            ]
            if (corridor, p["provider"]) in stale_pairs:
                for c in row:
                    c["stale"] = True
            rows.append(row)

    return {"rows": rows}
