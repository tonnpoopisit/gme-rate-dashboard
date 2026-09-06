#!/usr/bin/env python3
"""
ranking.py

Standalone, Excel-free ranking/gap logic for the Thailand/Laos rate reports.
Reads price_history.db (same schema price_history.py already writes) and
produces:
  - the latest per-corridor rate table, sorted cheapest-first, with the gap
    vs GME and enough min/max to size bar charts
  - a multi-window (3d/7d/15d/30d) average-price ranking plus raw time
    series, for trend charts

This is a fresh implementation of the same logic the laptop's
dashboard_server.py already has (_latest_rates/_rank_window/_weekly_history)
- written fresh here rather than imported, so this cloud app has zero
coupling to the laptop's file and one can change without the other breaking.
Ranking that used to be delegated to Excel's IFERROR/INDEX/MATCH/LARGE
formulas is just "sort by total ascending" here - Python does this natively,
no formula engine needed.
"""
import sqlite3
import statistics
from datetime import datetime, timedelta, timezone

# Base foreign-currency amount each fetch script actually queries - see
# BASE_THB in fetch_thailand_rates.py, BASE_LAK/BASE_USD/BASE_THB in
# fetch_laos_rates.py. Needed to invert "how much KRW for X foreign
# currency" into "how much foreign currency for 1,000,000 KRW": GME's fee is
# always written as 0 in the stored krw/fee columns (GME's row uses a single
# all-in KRW figure, not a decomposed rate+fee) but that figure still
# bundles in GME's real ~5,000 KRW flat sending fee, so the fee-free rate is
# (krw - GME_FLAT_FEE_KRW) / base_amount.
CORRIDOR_BASE_AMOUNTS = {"THB": 26_000, "LAK": 15_000_000, "USD": 1_000}
REVERSE_KRW = 1_000_000
GME_FLAT_FEE_KRW = 5_000

HISTORY_DAYS = 30
RANKING_WINDOWS_DAYS = [3, 7, 15, 30]
OUTLIER_MAX_DEVIATION_PCT = 7

# Canonical display order for Laos's 5 corridors - Bank Deposit LAK/USD/THB,
# then Cash Pickup LAK/USD, matching the real Excel workbook's block order
# (see fetch_laos_rates.py's module docstring). The SQL below naturally
# returns rows alphabetical-by-corridor instead ("LAK Bank Deposit, LAK Cash
# Pickup, THB Bank Deposit, ..."), so both dashboard endpoints and the PNG
# table (which just iterates whatever order these functions hand back) are
# reordered to this before returning - one fix, propagates everywhere.
CORRIDOR_DISPLAY_ORDER = [
    "LAK Bank Deposit", "USD Bank Deposit", "THB Bank Deposit",
    "LAK Cash Pickup", "USD Cash Pickup",
]

# Anchors price gap comparisons for corridors with multiple GME products (matches excel_table.py)
BASELINE_OVERRIDES = {"LAK Bank Deposit": "GME (Moneygram-BCEL)"}


def _ordered(by_corridor: dict) -> dict:
    """Reorders a {corridor: ...} dict to CORRIDOR_DISPLAY_ORDER, falling
    back to whatever order SQL produced for any corridor not listed there -
    e.g. Thailand's single "THB Bank Deposit" corridor, unaffected either way."""
    keys = [c for c in CORRIDOR_DISPLAY_ORDER if c in by_corridor]
    keys += [c for c in by_corridor if c not in keys]
    return {c: by_corridor[c] for c in keys}


def _reverse_conversions(corridor: str, providers: list):
    """For each GME row in this corridor, how much foreign currency would
    1,000,000 KRW get you. A corridor can have more than one GME row (Laos's
    LAK Bank Deposit lists both RIA and Moneygram-BCEL - genuinely different
    products, not duplicates), so this returns one entry per GME row."""
    currency = corridor.split()[0]
    base = CORRIDOR_BASE_AMOUNTS.get(currency)
    if base is None:
        return []
    decimals = 0 if currency == "LAK" else 2
    out = []
    for p in providers:
        if not p["isGme"] or not p["krw"]:
            continue
        true_rate_krw = p["krw"] - GME_FLAT_FEE_KRW
        if true_rate_krw <= 0:
            continue
        amount = round(REVERSE_KRW * base / true_rate_krw, decimals)
        out.append({"provider": p["provider"], "currency": currency, "amount": amount})
    return out


def latest_rates(report: str, db_path):
    """Latest row per (corridor, provider) for this report, grouped by
    corridor, sorted by total price ascending (cheapest first)."""
    if not db_path.exists():
        return {}
    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.execute(
            """
            SELECT corridor, provider, is_gme, krw, fee, total, price_gap_krw, timestamp
            FROM price_history
            WHERE report = ?
              AND timestamp = (
                  SELECT MAX(timestamp) FROM price_history p2
                  WHERE p2.report = price_history.report AND p2.corridor = price_history.corridor
                    AND p2.provider = price_history.provider
              )
            ORDER BY corridor, total ASC
            """,
            (report,),
        )
        rows = cur.fetchall()
    finally:
        conn.close()

    by_corridor = {}
    for corridor, provider, is_gme, krw, fee, total, gap, ts in rows:
        by_corridor.setdefault(corridor, {"asOf": ts, "providers": []})
        by_corridor[corridor]["providers"].append(
            {
                "provider": provider,
                "isGme": bool(is_gme),
                "krw": krw,
                "fee": fee,
                "total": total,
                "gap": gap,
            }
        )

    for corridor, block in by_corridor.items():
        baseline_name = BASELINE_OVERRIDES.get(corridor)
        baseline_p = next((p for p in block["providers"] if p["provider"] == baseline_name), None) if baseline_name else None
        if baseline_p is None:
            baseline_p = next((p for p in block["providers"] if p["isGme"]), None)
        baseline_total = baseline_p["total"] if baseline_p else None

        for p in block["providers"]:
            if baseline_p and p["provider"] == baseline_p["provider"]:
                p["gap"] = 0
            elif baseline_total is not None and (p.get("gap") is None or baseline_name):
                p["gap"] = p["total"] - baseline_total

        totals = [p["total"] for p in block["providers"]]
        block["minTotal"] = min(totals)
        block["maxTotal"] = max(totals)
        block["reverseConversions"] = _reverse_conversions(corridor, block["providers"])

    return _ordered(by_corridor)


def _rank_window(corridor: str, provider_points: dict, days_n: int, now: datetime):
    """provider_points: {provider: {"isGme": bool, "kept": [(ts, total), ...]}}
    (already outlier-filtered). Slices to the last `days_n` days, rolls up to
    a daily mean before averaging across days (so extra manual runs on one
    day don't skew it), and ranks by average price. Each window computes its
    own GME baseline, anchored to BASELINE_OVERRIDES if specified (e.g.
    GME Moneygram-BCEL in Laos LAK Bank Deposit)."""
    cutoff = (now - timedelta(days=days_n)).isoformat()
    ranking = []
    for provider, block in provider_points.items():
        pts = [(ts, total) for ts, total in block["kept"] if ts >= cutoff]
        if not pts:
            continue
        per_day = {}
        for ts, total in pts:
            per_day.setdefault(ts[:10], []).append(total)
        day_means = [sum(v) / len(v) for v in per_day.values()]
        avg_total = sum(day_means) / len(day_means)
        ranking.append(
            {
                "provider": provider,
                "isGme": block["isGme"],
                "avgTotal": round(avg_total, 2),
                "nDays": len(day_means),
                "nPoints": len(pts),
            }
        )

    ranking.sort(key=lambda r: r["avgTotal"])

    # Anchor to the designated GME baseline for this corridor
    baseline_provider = BASELINE_OVERRIDES.get(corridor)
    target = next((r for r in ranking if r["provider"] == baseline_provider), None) if baseline_provider else None
    if target is not None:
        gme_baseline = target["avgTotal"]
    else:
        gme_entries = [r for r in ranking if r["isGme"]]
        gme_baseline = gme_entries[0]["avgTotal"] if gme_entries else None

    for i, r in enumerate(ranking, start=1):
        r["rank"] = i
        r["gapVsGmeKrw"] = round(r["avgTotal"] - gme_baseline, 2) if gme_baseline is not None else None
    return ranking


def weekly_history(report: str, db_path, days: int = HISTORY_DAYS):
    """Per-corridor time series + multi-window average-price ranking, for a
    trend chart. A single wildly-off reading (e.g. a placeholder value) is
    dropped per-provider via its own median before anything here sees it -
    same spirit as price_history.is_plausible, just applied after the fact
    to a whole stored window instead of gating one incoming fetch."""
    if not db_path.exists():
        return {}
    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.execute(
            """
            SELECT corridor, provider, is_gme, total, timestamp
            FROM price_history
            WHERE report = ? AND timestamp >= datetime('now', ?)
            ORDER BY corridor, provider, timestamp
            """,
            (report, f"-{days} days"),
        )
        rows = cur.fetchall()
    finally:
        conn.close()

    by_corridor = {}
    for corridor, provider, is_gme, total, ts in rows:
        provider_block = by_corridor.setdefault(corridor, {}).setdefault(
            provider, {"isGme": bool(is_gme), "raw": []}
        )
        provider_block["raw"].append((ts, total))

    now = datetime.now(timezone.utc)
    out = {}
    for corridor, providers in by_corridor.items():
        kept_by_provider = {}
        for provider, block in providers.items():
            totals = [t for _, t in block["raw"]]
            med = statistics.median(totals)
            kept_by_provider[provider] = [
                (ts, total)
                for ts, total in block["raw"]
                if med == 0 or abs(total - med) / med * 100 <= OUTLIER_MAX_DEVIATION_PCT
            ]

        # Per-timestamp GME reference, anchored to BASELINE_OVERRIDES if present
        target_baseline = BASELINE_OVERRIDES.get(corridor)
        gme_by_ts = {}
        if target_baseline and target_baseline in kept_by_provider:
            for ts, total in kept_by_provider[target_baseline]:
                gme_by_ts[ts] = total
        else:
            for provider, block in providers.items():
                if not block["isGme"]:
                    continue
                for ts, total in kept_by_provider[provider]:
                    gme_by_ts.setdefault(ts, []).append(total)
            gme_by_ts = {ts: vals[0] for ts, vals in gme_by_ts.items()}

        series = {}
        provider_points = {}
        for provider, block in providers.items():
            kept = kept_by_provider[provider]
            series[provider] = {
                "isGme": block["isGme"],
                "points": [
                    {
                        "t": ts,
                        "total": total,
                        "gap": round(total - gme_by_ts[ts], 2) if ts in gme_by_ts else None,
                    }
                    for ts, total in kept
                ],
            }
            provider_points[provider] = {"isGme": block["isGme"], "kept": kept}

        windows = {
            str(n): {"days": n, "ranking": _rank_window(corridor, provider_points, n, now)}
            for n in RANKING_WINDOWS_DAYS
        }
        out[corridor] = {"series": series, "windows": windows}

    return _ordered(out)
