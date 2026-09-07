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

# Separate, dedicated bucket (public objectViewer for allUsers) just for the
# two rendered PNGs Teams needs to fetch directly - see teams_post.py. Kept
# entirely separate from GCS_BUCKET (which holds price_history.db and run
# status, both private) rather than making individual objects public there,
# since GCS_BUCKET has uniform bucket-level access enabled and so can only
# grant public read at the whole-bucket level, not per-object.
PUBLIC_GCS_BUCKET = os.environ.get("PUBLIC_GCS_BUCKET", "").strip()

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


def download_known_fees(path: Path):
    """Pulls the latest known_fees.json from GCS to `path`, if a bucket is
    configured - mirrors download_db's persistence-across-cold-starts fix,
    which this file never got originally (confirmed: no known_fees.json ever
    showed up in the bucket, so a fee change the daily checker found was
    being silently lost whenever the container scaled to zero before the
    next rate-fetch run read it back). No caching (unlike download_db): this
    file is tiny and read far less often, so a fresh pull every call is
    cheap and always correct. Takes path explicitly (rather than importing
    known_fees for its DEFAULT_PATH) to avoid a circular import, since
    known_fees.py is the one importing this module."""
    if not GCS_BUCKET:
        return
    blob = _client().bucket(GCS_BUCKET).blob("known_fees.json")
    if blob.exists():
        blob.download_to_filename(str(path))


def upload_known_fees(path: Path):
    """Pushes a freshly-written known_fees.json back to GCS."""
    if not GCS_BUCKET or not path.exists():
        return
    _client().bucket(GCS_BUCKET).blob("known_fees.json").upload_from_filename(str(path))


def upload_public_png(local_path: Path, object_name: str) -> str:
    """Uploads a copy of local_path to PUBLIC_GCS_BUCKET under object_name
    and returns its public URL - used for the Teams card's Image url so
    Teams' own servers fetch the image directly, rather than embedding it as
    base64 in the webhook body (a hard ~28KB limit the rendered table had
    already outgrown once - see teams_post.py's git history).

    object_name must be unique per call (caller passes a timestamped name -
    see teams_post.py) rather than a fixed name like local_path.name:
    confirmed live that reusing the same URL for every post causes Teams to
    display a stale, client-cached image instead of re-fetching - a
    Thailand post showed the *previous day's* rate despite the underlying
    pipeline having fetched and rendered correctly. A genuinely new URL per
    post has nothing to serve from cache.

    Raises if PUBLIC_GCS_BUCKET isn't configured (local dev has no public
    bucket to publish to)."""
    if not PUBLIC_GCS_BUCKET:
        raise RuntimeError("PUBLIC_GCS_BUCKET not configured - can't publish an image for Teams to fetch")
    blob = _client().bucket(PUBLIC_GCS_BUCKET).blob(object_name)
    # Belt-and-suspenders alongside the unique name - discourages any
    # intermediate proxy/CDN from caching this object under its own URL too.
    blob.cache_control = "no-store, max-age=0"
    blob.upload_from_filename(str(local_path))
    return f"https://storage.googleapis.com/{PUBLIC_GCS_BUCKET}/{object_name}"


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
