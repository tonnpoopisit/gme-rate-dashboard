#!/usr/bin/env python3
"""
price_history.py

Shared, report-agnostic logger for the Thailand/Laos rate-report pipelines.
Appends one row per (corridor, provider) to a local SQLite database every
time a fetch script is run with --log-history, building up the history
needed to answer "how many days has GME been more expensive than competitor
X, by how much, by what %" - see
.claude/skills/thailand-rate-report/references/price_history_queries.md for
ready-to-run example queries.

Opt-in by design (see --log-history in fetch_thailand_rates.py /
fetch_laos_rates.py): only the scheduled automation passes it, so ad-hoc/
manual/debugging runs against scratch copies don't pollute the real history.
"""
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_DB_PATH = Path(__file__).with_name("price_history.db")

_SCHEMA = """
CREATE TABLE IF NOT EXISTS price_history (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp       TEXT NOT NULL,
    report          TEXT NOT NULL,
    corridor        TEXT NOT NULL,
    provider        TEXT NOT NULL,
    is_gme          INTEGER NOT NULL,
    krw             INTEGER NOT NULL,
    fee             INTEGER NOT NULL,
    total           INTEGER NOT NULL,
    price_gap_krw   INTEGER,
    price_gap_pct   REAL
);
CREATE INDEX IF NOT EXISTS idx_price_history_lookup
    ON price_history(report, corridor, provider, timestamp);
"""

# Separate table from price_history: market spot rates are one row per
# (currency, timestamp) with no provider/corridor/fee concept at all, so
# forcing them into price_history's per-provider-quote schema would just
# mean a pile of NULL columns - see fetch_market_rates.py.
_MARKET_SCHEMA = """
CREATE TABLE IF NOT EXISTS market_rates (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp       TEXT NOT NULL,
    currency        TEXT NOT NULL,
    krw_per_unit    REAL NOT NULL,
    source          TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_market_rates_lookup
    ON market_rates(currency, timestamp);
"""


def log_run(report, entries, db_path=DEFAULT_DB_PATH, timestamp=None):
    """Append one history row per entry for this run.

    `entries`: list of dicts, one per (corridor, provider) actually fetched
    this run - {"corridor": str, "provider": str, "krw": int, "fee": int,
    "is_gme_baseline": bool (optional)}. A provider whose result is reused
    across multiple corridors (e.g. Laos's Hanpass, whose API doesn't
    distinguish Bank Deposit from Cash Pickup) should appear once per
    corridor it's written into, since "days more expensive" is a
    corridor-specific question.

    Sign convention matches the sheet's own Price Gap column: price_gap_krw
    = total - <corridor's GME baseline total>, so positive means GME is more
    expensive, negative means GME is cheaper.

    GME baseline selection per corridor: if any entry is explicitly marked
    is_gme_baseline=True, that one is used (needed where a corridor has more
    than one GME row - e.g. Laos's LAK Bank Deposit has both "GME
    (RIA-Other banks)" and "GME (Moneygram-BCEL)", and the sheet's own
    Price Gap formulas are anchored to Moneygram-BCEL there, confirmed by
    reading the workbook's own Q7 reference cell rather than assumed).
    Otherwise, the first entry whose provider name starts with "GME" is used
    - correct for every corridor that only ever has one GME row. If a
    corridor has no GME entry at all this run (GME's own fetch failed),
    every provider in it logs with a NULL gap rather than a misleading
    comparison against a stale baseline.
    """
    ts = timestamp or datetime.now(timezone.utc).isoformat()

    by_corridor = {}
    for e in entries:
        by_corridor.setdefault(e["corridor"], []).append(e)

    rows = []
    for corridor, corridor_entries in by_corridor.items():
        baseline = next((e for e in corridor_entries if e.get("is_gme_baseline")), None)
        if baseline is None:
            baseline = next((e for e in corridor_entries if e["provider"].startswith("GME")), None)
        gme_total = (baseline["krw"] + baseline["fee"]) if baseline else None

        for e in corridor_entries:
            total = e["krw"] + e["fee"]
            is_gme = e["provider"].startswith("GME")
            is_baseline = (baseline and e["provider"] == baseline["provider"])
            if is_baseline:
                gap_krw = 0
                gap_pct = 0.0
            elif gme_total is not None:
                gap_krw = total - gme_total
                gap_pct = (gap_krw / gme_total) * 100 if gme_total else None
            else:
                gap_krw = gap_pct = None
            rows.append((
                ts, report, corridor, e["provider"], int(is_gme),
                e["krw"], e["fee"], total, gap_krw, gap_pct,
            ))

    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(_SCHEMA)
        conn.executemany(
            """INSERT INTO price_history
               (timestamp, report, corridor, provider, is_gme, krw, fee, total, price_gap_krw, price_gap_pct)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            rows,
        )
        conn.commit()
    finally:
        conn.close()

    return len(rows)


def ensure_market_schema(db_path=DEFAULT_DB_PATH):
    """Creates market_rates (and its index) if it doesn't exist yet, without
    inserting anything - lets callers query the table before any row has
    ever been logged into it."""
    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(_MARKET_SCHEMA)
        conn.commit()
    finally:
        conn.close()


def log_market_rates(rates, source, db_path=DEFAULT_DB_PATH, timestamp=None):
    """Append one row per currency to market_rates.

    `rates`: dict of {currency: krw_per_unit}, e.g.
    {"USD": 1439.6, "THB": 42.9, "LAK": 0.0646} - how many KRW one unit of
    that currency buys, matching the direction remittance quotes are
    naturally read in (send foreign currency, receive/compare in KRW).
    """
    ts = timestamp or datetime.now(timezone.utc).isoformat()
    rows = [(ts, currency, krw_per_unit, source) for currency, krw_per_unit in rates.items()]

    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(_MARKET_SCHEMA)
        conn.executemany(
            "INSERT INTO market_rates (timestamp, currency, krw_per_unit, source) VALUES (?, ?, ?, ?)",
            rows,
        )
        conn.commit()
    finally:
        conn.close()

    return len(rows)


def is_plausible(report, corridor, provider, new_krw, db_path=DEFAULT_DB_PATH, lookback=5, max_deviation_pct=7):
    """Sanity-check a freshly fetched KRW value against this exact
    (report, corridor, provider)'s own recent history before trusting it.

    Exists because a fetcher can "succeed" - no exception, no timeout - while
    still returning the wrong number, if a page interaction silently doesn't
    take effect (confirmed for real: GME's Thailand fetcher read a page's
    leftover default/placeholder amount of exactly 1,000,000 KRW instead of
    the real "You Send" figure at least once, and because nothing raised, the
    existing stale-row amber-flagging never fired - the bad number went out
    looking exactly as trustworthy as a real one). A hardcoded absolute
    plausible range would need periodic retuning as real exchange rates
    genuinely drift over months; comparing against this specific provider's
    own recent fetches doesn't.

    Uses the MEDIAN of the last `lookback` logged values, not the mean, so a
    single already-bad historical entry in that window (like the 1,000,000
    one above, if it already got logged before this check existed) doesn't
    skew the baseline it's being compared against - confirmed against real
    history: GME's Thailand fetches cluster within <1% of each other run to
    run, while that one bad value was ~12% off, so max_deviation_pct=7 has
    comfortable margin on both sides (rejects the actual bad value seen,
    passes every real fetch-to-fetch fluctuation observed so far).

    No history yet for this (report, corridor, provider) - nothing to
    compare against, so this returns True (accept it; there's no way to
    validate a first-ever reading).
    """
    if not Path(db_path).exists():
        return True
    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.execute(
            """SELECT krw FROM price_history
               WHERE report = ? AND corridor = ? AND provider = ?
               ORDER BY timestamp DESC LIMIT ?""",
            (report, corridor, provider, lookback),
        )
        history = [row[0] for row in cur.fetchall()]
    finally:
        conn.close()

    if not history:
        return True

    history.sort()
    median = history[len(history) // 2]
    if median == 0:
        return True
    deviation_pct = abs(new_krw - median) / median * 100
    return deviation_pct <= max_deviation_pct
