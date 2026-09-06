#!/usr/bin/env python3
"""
pipeline.py

Cloud-native replacement for run_hourly_report.ps1 / run_hourly_laos_report.ps1.
Fetches live rates for one report, logs them to the (GCS-synced)
price_history.db, renders the ranking-table PNG, and (when post_to_teams=True)
posts it to Teams via teams_post.py - no Excel, no PowerShell.

Entry-building logic below is copied verbatim from each fetch script's own
main() --log-history block (fetch_thailand_rates.py / fetch_laos_rates.py) -
that logic never touched openpyxl to begin with, so it moves here unchanged.

Usage:
    python pipeline.py --report thailand
    python pipeline.py --report laos
    python pipeline.py --report thailand --gme-manual-krw 1123346
    python pipeline.py --report thailand --no-log-history   # dry-run-ish, don't persist
"""
import argparse
import json
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

import excel_table
import gcs_sync
import price_history
import ranking
import teams_post

PROJECT_ROOT = Path(__file__).parent
PNG_NAMES = {"thailand": "Thailand_left_table.png", "laos": "Laos_left_table.png"}
_PIPELINE_LOCK = threading.Lock()


def run_thailand(gme_manual_krw=None, log_history=True):
    import fetch_thailand_rates as th

    results = th.fetch_all(gme_manual=gme_manual_krw)
    fetched, failed = len(results), len(th.INPUT_AREA_ROWS) - len(results)
    stale_pairs = {("THB Bank Deposit", p) for p in th.INPUT_AREA_ROWS if p not in results}

    if log_history:
        db_path = gcs_sync.download_db(force=True)
        entries = [
            {"corridor": "THB Bank Deposit", "provider": provider, "krw": vals["krw"], "fee": vals["fee"]}
            for provider, vals in results.items()
        ]
        n = price_history.log_run("thailand", entries, db_path=db_path)
        gcs_sync.upload_db(db_path)
    else:
        n = 0

    return {"fetched": fetched, "failed": failed, "loggedRows": n, "results": results, "stalePairs": stale_pairs}


def run_laos(log_history=True):
    import fetch_laos_rates as laos

    fetched_raw = laos.fetch_all()
    fetched, failed = len(fetched_raw), len(laos.ROW_SOURCES) - len(
        [k for _, k in laos.ROW_SOURCES if k in fetched_raw]
    )
    stale_pairs = {
        (laos.ROW_TO_CORRIDOR[row], key[0]) for row, key in laos.ROW_SOURCES if key not in fetched_raw
    }

    if log_history:
        db_path = gcs_sync.download_db(force=True)
        entries = [
            {
                "corridor": laos.ROW_TO_CORRIDOR[row],
                "provider": key[0],
                "krw": fetched_raw[key]["krw"],
                "fee": fetched_raw[key]["fee"],
                # Block 1 (LAK Bank Deposit) has two GME rows - the sheet's
                # own Price Gap formulas there are anchored to
                # Moneygram-BCEL, so that's the baseline here too.
                "is_gme_baseline": laos.ROW_TO_CORRIDOR[row] == "LAK Bank Deposit" and key[0] == "GME (Moneygram-BCEL)",
            }
            for row, key in laos.ROW_SOURCES
            if key in fetched_raw
        ]
        n = price_history.log_run("laos", entries, db_path=db_path)
        gcs_sync.upload_db(db_path)
    else:
        n = 0

    # JSON-friendly key for the results echoed back in the status/summary.
    results = {f"{p} ({c})": v for (p, c), v in fetched_raw.items()}
    return {"fetched": fetched, "failed": failed, "loggedRows": n, "results": results, "stalePairs": stale_pairs}


def render_png(report: str, stale_pairs: set) -> str:
    """Builds the Excel-style ranking table PNG from whatever's now in
    price_history.db (so this only makes sense right after a log_history=True
    run) via the existing render_table_png.py renderer - unchanged, just fed
    JSON built fresh in Python instead of read out of Excel."""
    db_path = gcs_sync.download_db()
    latest = ranking.latest_rates(report, db_path)
    table = excel_table.build_table_rows(report, latest, stale_pairs)

    json_path = PROJECT_ROOT / f"table_data_{report}.json"
    png_path = PROJECT_ROOT / PNG_NAMES[report]
    json_path.write_text(json.dumps(table))

    args = [sys.executable, str(PROJECT_ROOT / "render_table_png.py"), str(json_path), str(png_path)]
    if report != "laos":
        # --compact: this PNG doubles as the Teams webhook attachment (see
        # teams_post.py), which hard-rejects bodies over ~28KB once wrapped -
        # a default-quality Thailand render alone runs 130KB+, so render
        # compact unconditionally rather than only when actually posting.
        args += ["--compact"]
    if report == "laos":
        # 16 was too aggressive: with ~10 distinct flat colors (6 provider
        # backgrounds + white/black/border-grey/negative-gap-red) plus every
        # anti-aliased text/border edge adding its own blended shades, the
        # adaptive palette dropped GME's rarely-used salmon (#F4777F, only
        # 1-2 small cells) in favor of more pixel-frequent colors - it came
        # out as a wrong, muddy (204,147,154) instead (confirmed via direct
        # pixel sampling, not guessed). 64 keeps real file-size savings
        # (Laos's table is much smaller than Thailand's, which uses the
        # unquantized default of 32) while giving the palette enough room to
        # keep every semantic color true.
        args += ["--compact", "--font-size", "11", "--pad-v", "2", "--pad-h", "5", "--colors", "64"]
    subprocess.run(args, check=True, cwd=str(PROJECT_ROOT))

    gcs_sync.upload_png(report, png_path)
    return str(png_path)


def run(report: str, gme_manual_krw=None, log_history=True, post_to_teams=False) -> dict:
    with _PIPELINE_LOCK:
        started = time.time()
        gcs_sync.write_status(report, {"state": "running", "startedAt": started})
        try:
            if report == "thailand":
                summary = run_thailand(gme_manual_krw=gme_manual_krw, log_history=log_history)
            elif report == "laos":
                summary = run_laos(log_history=log_history)
            else:
                raise ValueError(f"Unknown report '{report}'")
            stale_pairs = summary.pop("stalePairs", set())
            if log_history:
                render_png(report, stale_pairs)
                if post_to_teams:
                    # A post failure fails this whole run (propagates out of
                    # the try below) - same as the laptop pipeline treating a
                    # failed Send-TeamsImage as a whole-run failure, since the
                    # entire point of a scheduled run is to post.
                    teams_post.post_to_teams(report, PROJECT_ROOT / f"table_data_{report}.json")
            status = {
                "state": "success",
                "startedAt": started,
                "finishedAt": time.time(),
                **summary,
            }
        except Exception as e:  # noqa: BLE001 - always record status, even on a hard failure
            status = {
                "state": "error",
                "startedAt": started,
                "finishedAt": time.time(),
                "error": str(e),
                "traceback": traceback.format_exc(),
            }
            gcs_sync.write_status(report, status)
            raise
        gcs_sync.write_status(report, status)
        return status


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, choices=["thailand", "laos"])
    parser.add_argument("--gme-manual-krw", type=int, default=None, help="Thailand only - see fetch_thailand_rates.py --gme-manual")
    parser.add_argument("--no-log-history", action="store_true", help="Fetch and print only, don't write to price_history.db")
    args = parser.parse_args()

    if args.gme_manual_krw is not None and args.report != "thailand":
        sys.exit("--gme-manual-krw only applies to --report thailand")

    status = run(args.report, gme_manual_krw=args.gme_manual_krw, log_history=not args.no_log_history)

    print(f"\n--- {args.report} pipeline: {status['state']} ---")
    if status["state"] == "success":
        print(f"Fetched {status['fetched']}, failed {status['failed']}, logged {status['loggedRows']} rows")
        for name, vals in status["results"].items():
            total = vals["krw"] + vals["fee"]
            print(f"  {name:<28} KRW={vals['krw']:>10,}  fee={vals['fee']:>7,}  total={total:>10,}")
    else:
        print(status.get("error"))
        sys.exit(1)


if __name__ == "__main__":
    main()
