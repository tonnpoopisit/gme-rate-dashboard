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

The card's image is referenced by URL (gcs_sync.upload_public_png), not
embedded as base64: this webhook is a Power Platform "manual trigger" URL,
confirmed (on the laptop pipeline, see run_hourly_report.ps1) to hard-reject
any request body over ~28KB, and a full-quality render of even a plain
11-row table alone runs 60-90KB once base64-encoded - long before this
report grows further. An earlier version of this file re-rendered at
progressively smaller/uglier settings to squeeze under that limit; the
Laos table had already grown enough to need illegibly small text even then.
Referencing a URL instead makes the request body tiny regardless of image
size, so the image can just be full quality.
"""
import json
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

import gcs_sync

PROJECT_ROOT = Path(__file__).parent
KST = ZoneInfo("Asia/Seoul")

REPORT_LABELS = {"thailand": "Thailand", "laos": "Laos"}


def _webhook_config_path(report: str) -> Path:
    import os

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


def _build_card_body(image_url: str, title: str) -> bytes:
    adaptive_card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [
            {"type": "TextBlock", "text": title, "weight": "Bolder", "size": "Medium"},
            {"type": "Image", "url": image_url, "size": "Stretch"},
        ],
    }
    body = {
        "type": "message",
        "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive", "content": adaptive_card}],
    }
    return json.dumps(body).encode("utf-8")


def post_to_teams(report: str, png_path: Path) -> None:
    """Publishes png_path to the public images bucket and posts a Teams card
    referencing that URL. Raises on any failure - publishing, the HTTP call
    itself, or a non-2xx response - so a scheduled run that can't actually
    post is recorded as a failed run (see pipeline.run), same as the laptop
    pipeline treated a failed Send-TeamsImage as a whole-run failure worth
    alerting on."""
    label = REPORT_LABELS[report]
    title = f"{label} rate report {datetime.now(KST).strftime('%Y-%m-%d %H:%M')}"
    image_url = gcs_sync.upload_public_png(png_path)
    body = _build_card_body(image_url, title)
    url = _webhook_url(report)
    resp = requests.post(url, data=body, headers={"Content-Type": "application/json; charset=utf-8"}, timeout=60)
    resp.raise_for_status()
