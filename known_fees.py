#!/usr/bin/env python3
"""
known_fees.py

Shared store for the handful of provider fees that can't be read live from
the same calculator response used for their rate (see
check_competitor_fees.py's module docstring for which providers and why).
Both the daily fee checker and the regular 2-hourly rate fetchers read this
file - the checker writes it when it finds a real change, the rate fetchers
read it every run instead of using a hardcoded literal.

known_fees.json (created on first write) looks like:
    {
      "WireBarley": {"fee_krw": 3000, "checked_at": "2026-08-10T18:00:03+09:00", "source": "https://..."},
      ...
    }
"""

import json
from datetime import datetime
from pathlib import Path

import gcs_sync

DEFAULT_PATH = Path(__file__).with_name("known_fees.json")


def get_fee(provider, default, path=DEFAULT_PATH):
    """Returns the currently-known fee for `provider`, or `default` if the
    file doesn't exist yet or has no entry for it - so every caller works
    unchanged before the checker has ever run once. Pulls the latest copy
    from GCS first (no-op locally / when no bucket is configured) so a fee
    the daily checker found on a different, since-recycled container
    instance is still picked up here."""
    gcs_sync.download_known_fees(path)
    if not path.exists():
        return default
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return default
    entry = data.get(provider)
    if not entry or "fee_krw" not in entry:
        return default
    return entry["fee_krw"]


def set_fee(provider, fee_krw, source, path=DEFAULT_PATH):
    """Records a freshly-checked fee for `provider`. Read-modify-write of
    the whole file (fine at this size/update frequency - once a day at
    most, three providers total). Pushes the updated file back to GCS
    afterward so the change survives this container instance being
    recycled before the next rate-fetch run needs it."""
    data = {}
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            data = {}
    data[provider] = {
        "fee_krw": fee_krw,
        "checked_at": datetime.now().astimezone().isoformat(),
        "source": source,
    }
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    gcs_sync.upload_known_fees(path)
