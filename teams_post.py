#!/usr/bin/env python3
"""
teams_post.py

Posts a rendered ranking-table PNG to a Microsoft Teams channel via a Power
Automate webhook - the Cloud Run equivalent of run_hourly_report.ps1's
Send-TeamsImage function (the laptop pipeline this replaces). Kept as its own
module rather than folded into pipeline.py since it's the one piece that
talks to an external, secret-bearing URL.

Webhook URLs are bearer-credential-equivalent (anyone with the URL can post
to the channel) and this repo is public, so they're never read from source or
committed - only from a Secret-Manager-mounted file (TEAMS_WEBHOOK_*_PATH env
vars, same pattern as AUTH_PATH/GME_SECRETS_PATH in dashboard_server.py) or,
for local dev with no secrets mounted, the same webhook_config.json /
webhook_config_laos.json filenames the laptop pipeline already uses.

Body-size handling: this URL is a Power Platform "manual trigger" webhook,
confirmed (on the laptop pipeline, see run_hourly_report.ps1) to hard-reject
any request body over ~28KB, and the request body must be a complete Adaptive
Card wrapped in a Bot Framework message envelope - not a custom JSON schema.
Rather than trust one fixed render size (verified independently: the actual
Linux/Playwright container renders meaningfully larger PNGs than the same
settings produce on a laptop, almost certainly a Calibri-vs-fallback-font
difference - so a size tuned by testing locally can't be trusted blind), this
re-renders at progressively more aggressive settings until the wrapped body
actually fits, checked for real each time rather than assumed.
"""
import base64
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

PROJECT_ROOT = Path(__file__).parent
KST = ZoneInfo("Asia/Seoul")

# Hard limit is ~28KB; stay a safety margin under it (same 27KB the laptop
# pipeline used) since the JSON wrapper adds a little on top of the base64
# image itself.
BODY_LIMIT_BYTES = 27 * 1024

# Each report's cascade of render settings, most-preferred (best visual
# quality) first. Thailand's table is small enough that plain --compact has
# always been enough (laptop pipeline default, ~9-13KB even accounting for
# the Linux font-rendering gap noted above); Laos's 32-row/5-block table
# needs more aggressive shrinking and gets a real fallback ladder. --colors
# 64 (not the smaller values used further down the ladder) is the preferred
# Laos setting because anything lower was confirmed to drop GME's
# rarely-used salmon (#F4777F) in favor of a wrong muddy color (see
# pipeline.py's render_png) - only fall back to fewer colors if 64 doesn't
# fit, trading that color accuracy for actually fitting the payload.
RENDER_PRESETS = {
    "thailand": [
        ["--compact"],
        ["--compact", "--colors", "16"],
    ],
    "laos": [
        ["--compact", "--font-size", "11", "--pad-v", "2", "--pad-h", "5", "--colors", "64"],
        # font-size 11 / pad 2,5 / colors 16 is render_table_png.py's own
        # documented, previously-verified-in-production Laos setting (before
        # colors got bumped to 64 to fix GME's salmon getting quantized away)
        # - a known-good fallback rather than another guess, just missing
        # that color-accuracy fix.
        ["--compact", "--font-size", "11", "--pad-v", "2", "--pad-h", "5", "--colors", "16"],
        ["--compact", "--font-size", "10", "--pad-v", "1", "--pad-h", "3", "--colors", "24"],
        ["--compact", "--font-size", "9", "--pad-v", "1", "--pad-h", "3", "--colors", "16"],
    ],
}

REPORT_LABELS = {"thailand": "Thailand", "laos": "Laos"}


def _webhook_config_path(report: str) -> Path:
    env_key = f"TEAMS_WEBHOOK_{report.upper()}_PATH"
    override = os.environ.get(env_key)
    if override:
        return Path(override)
    # Local dev fallback - same filenames the laptop pipeline already reads
    # from (webhook_config.json / webhook_config_laos.json), so dropping
    # those two files into this folder is all local testing needs.
    return PROJECT_ROOT / ("webhook_config.json" if report == "thailand" else "webhook_config_laos.json")


def _webhook_url(report: str) -> str:
    path = _webhook_config_path(report)
    if not path.exists():
        raise RuntimeError(f"No Teams webhook config found for '{report}' at {path}")
    url = json.loads(path.read_text()).get("teams_webhook_url", "").strip()
    if not url:
        raise RuntimeError(f"teams_webhook_url missing/empty in {path}")
    return url


def _build_card_body(png_path: Path, title: str) -> bytes:
    b64 = base64.b64encode(png_path.read_bytes()).decode("ascii")
    adaptive_card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [
            {"type": "TextBlock", "text": title, "weight": "Bolder", "size": "Medium"},
            {"type": "Image", "url": f"data:image/png;base64,{b64}", "size": "Stretch"},
        ],
    }
    body = {
        "type": "message",
        "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive", "content": adaptive_card}],
    }
    return json.dumps(body).encode("utf-8")


def _render(report: str, table_json_path: Path, out_path: Path, extra_args: list) -> Path:
    args = [sys.executable, str(PROJECT_ROOT / "render_table_png.py"), str(table_json_path), str(out_path), *extra_args]
    subprocess.run(args, check=True, cwd=str(PROJECT_ROOT))
    return out_path


def render_for_teams(report: str, table_json_path: Path, title: str) -> bytes:
    """Tries each of this report's render presets in order (best quality
    first) against a scratch PNG, actually measuring the wrapped Adaptive
    Card body each time, and returns the body of the first one that fits.
    Raises if even the most aggressive preset doesn't - a loud, visible
    failure (surfaces via pipeline.run's existing error status) rather than
    ever sending a body Power Automate would just reject anyway."""
    scratch_png = PROJECT_ROOT / f"_teams_render_{report}.png"
    last_size = None
    for preset in RENDER_PRESETS[report]:
        _render(report, table_json_path, scratch_png, preset)
        body = _build_card_body(scratch_png, title)
        last_size = len(body)
        if last_size <= BODY_LIMIT_BYTES:
            return body
    raise RuntimeError(
        f"No render preset for '{report}' fit the Teams webhook's ~28KB limit "
        f"(smallest attempt was {last_size / 1024:.1f}KB) - needs a more aggressive preset in teams_post.RENDER_PRESETS"
    )


def post_to_teams(report: str, table_json_path: Path) -> None:
    """Renders (with fallback to smaller presets if needed) and posts this
    report's ranking table to its configured Teams webhook. Raises on any
    failure - fitting the payload, the HTTP call itself, or a non-2xx
    response - so a scheduled run that can't actually post is recorded as a
    failed run (see pipeline.run), same as the laptop pipeline treated a
    failed Send-TeamsImage as a whole-run failure worth alerting on.
    """
    label = REPORT_LABELS[report]
    title = f"{label} rate report {datetime.now(KST).strftime('%Y-%m-%d %H:%M')}"
    body = render_for_teams(report, table_json_path, title)
    url = _webhook_url(report)
    resp = requests.post(url, data=body, headers={"Content-Type": "application/json; charset=utf-8"}, timeout=60)
    resp.raise_for_status()
