#!/usr/bin/env python3
"""
gcs_sync.py

Keeps price_history.db durable across Cloud Run's ephemeral filesystem by
downloading/uploading it as a whole GCS object around each pipeline run,
rather than mounting it live (SQLite's locking model doesn't hold up well
over a network filesystem like GCS FUSE - see the migration plan).

Local dev / no bucket configured: every function is a no-op and the DB just
lives at its normal local path (price_history.DEFAULT_DB_PATH) - this lets
the whole fetch -> rank -> dashboard loop be exercised on a laptop with zero
GCP setup before ever touching Cloud Run.
"""
import os
import time
from pathlib import Path

import price_history

GCS_BUCKET = os.environ.get("GCS_BUCKET", "").strip()

# Cache downloads for a short window so a burst of dashboard page loads
# doesn't re-download the DB from GCS on every single request.
_DOWNLOAD_CACHE_SECONDS = 30
_last_download_at = 0.0


def _client():
    from google.cloud import storage

    return storage.Client()


def db_path() -> Path:
    """The local path price_history.py itself reads/writes by default -
    fetch_laos_rates.py's internal is_plausible() check uses this same
    default path (it doesn't take a db_path override), so keeping the
    GCS-synced copy at exactly this location is what makes that check see
    up-to-date data."""
    return price_history.DEFAULT_DB_PATH


def download_db(force: bool = False) -> Path:
    """Pull the latest price_history.db from GCS to the local default path,
    if a bucket is configured. Cached for _DOWNLOAD_CACHE_SECONDS unless
    force=True (pipeline runs should force; dashboard reads can use the
    cache)."""
    global _last_download_at
    path = db_path()
    if not GCS_BUCKET:
        return path

    now = time.time()
    if not force and (now - _last_download_at) < _DOWNLOAD_CACHE_SECONDS and path.exists():
        return path

    bucket = _client().bucket(GCS_BUCKET)
    blob = bucket.blob("price_history.db")
    if blob.exists():
        blob.download_to_filename(str(path))
    _last_download_at = now
    return path


def upload_db(path: Path = None):
    """Push the local price_history.db back to GCS after a pipeline run."""
    if not GCS_BUCKET:
        return
    path = path or db_path()
    if not path.exists():
        return
    bucket = _client().bucket(GCS_BUCKET)
    blob = bucket.blob("price_history.db")
    blob.upload_from_filename(str(path))


def upload_png(report: str, local_path: Path):
    if not GCS_BUCKET:
        return
    bucket = _client().bucket(GCS_BUCKET)
    bucket.blob(local_path.name).upload_from_filename(str(local_path))


def download_png(report: str, local_path: Path) -> bool:
    """Pulls the latest PNG from GCS to local_path if a bucket is
    configured. Returns whether a file is actually available at local_path
    afterward (either just downloaded, or already there for local dev)."""
    if not GCS_BUCKET:
        return local_path.exists()
    bucket = _client().bucket(GCS_BUCKET)
    blob = bucket.blob(local_path.name)
    if not blob.exists():
        return local_path.exists()
    blob.download_to_filename(str(local_path))
    return True


def write_status(report: str, status: dict):
    """Small JSON status blob per report, for the dashboard's job-polling
    endpoint to read - stateless across cold starts, unlike an in-memory
    dict. Local dev: written to a local file instead, same read path."""
    import json

    payload = json.dumps(status)
    local_path = Path(__file__).with_name(f"status_{report}.json")
    local_path.write_text(payload)
    if not GCS_BUCKET:
        return
    bucket = _client().bucket(GCS_BUCKET)
    bucket.blob(f"status/{report}.json").upload_from_string(payload, content_type="application/json")


# A real run takes 30-90s; anything still "running" past this either died
# mid-flight (process killed/restarted/crashed) without writing a final
# status, or is Cloud Run's own retry rerunning after a real failure - never
# trust a "running" state indefinitely.
STALE_RUNNING_SECONDS = 600


def _resolve_stale(status: dict) -> dict:
    if status.get("state") == "running" and time.time() - status.get("startedAt", 0) > STALE_RUNNING_SECONDS:
        return {"state": "unknown", "message": "Previous run didn't finish (interrupted or crashed) - try Refresh again."}
    return status


def read_status(report: str) -> dict:
    import json

    if GCS_BUCKET:
        bucket = _client().bucket(GCS_BUCKET)
        blob = bucket.blob(f"status/{report}.json")
        if blob.exists():
            return _resolve_stale(json.loads(blob.download_as_text()))
        return {"state": "unknown"}

    local_path = Path(__file__).with_name(f"status_{report}.json")
    if local_path.exists():
        return _resolve_stale(json.loads(local_path.read_text()))
    return {"state": "unknown"}
