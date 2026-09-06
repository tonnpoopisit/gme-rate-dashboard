#!/usr/bin/env python3
"""
dashboard_server.py

Cloud-native rate dashboard: fetch Thailand/Laos rates on demand, rank them
in Python (no Excel), show a table + charts, render a PNG, and (only for
/api/cron/run's scheduled runs, not the dashboard's own manual Refresh - see
teams_post.py) post that PNG to Teams. This is a separate, standalone app
from the laptop's dashboard_server.py - no Excel/PowerShell orchestration.

Same two-tier PIN auth as the laptop dashboard (team / owner), same
generate-once-and-print-to-console pattern for first run. Secrets
(dashboard_auth.json, dashboard_secret_key.txt, gme_mobile_secrets.json) are
plain local files - on Cloud Run these paths are backed by mounted Secret
Manager secrets (--set-secrets), so this code is identical between local dev
and cloud deployment.

Run locally:
    python dashboard_server.py
Then open http://127.0.0.1:5151
"""
import base64
import hashlib
import json
import os
import secrets
import threading
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory, abort, redirect, session
from werkzeug.middleware.proxy_fix import ProxyFix

import fetch_laos_rates as laos_fetch
import gcs_sync
import pipeline
import ranking

PROJECT_ROOT = Path(__file__).parent

# Overridable via env var because Cloud Run's --set-secrets can't mount two
# different secrets as files in the same directory (each secret owns its
# whole mount directory) - the deploy command points each of these at its
# own subdirectory. Local dev sets none of these, so all three fall back to
# the plain PROJECT_ROOT files exactly as before.
AUTH_PATH = Path(os.environ.get("AUTH_PATH", str(PROJECT_ROOT / "dashboard_auth.json")))
SECRET_KEY_PATH = Path(os.environ.get("SECRET_KEY_PATH", str(PROJECT_ROOT / "dashboard_secret_key.txt")))
GME_SECRETS_PATH = Path(os.environ.get("GME_SECRETS_PATH", str(PROJECT_ROOT / "gme_mobile_secrets.json")))

# On Cloud Run, K_SERVICE is always set (part of the platform's own runtime
# contract) - used here to detect "am I running as the mounted-secrets
# deployment" vs. plain local dev, since GME_SECRETS_PATH is a read-only
# Secret Manager volume mount in the former case (see _persist_gme_secrets).
CLOUD_RUN_SERVICE = os.environ.get("K_SERVICE", "")
GME_SECRETS_NAME = os.environ.get("GME_SECRETS_SECRET_NAME", "gme-mobile-secrets")


def _persist_gme_secrets(existing: dict):
    """Writes the updated GME secrets (Authorization refreshed, cookie/
    username carried over). Locally this is just the plain JSON file at
    GME_SECRETS_PATH; on Cloud Run that same path is a read-only
    Secret-Manager-mounted volume (--set-secrets in the deploy command), so
    a write there raises - a new secret version is added via the API
    instead. Cloud Run refreshes a mounted volume to the latest version on
    its own (no redeploy needed), just not instantly - fine for a token
    that's refreshed by hand roughly daily, not for something latency-sensitive."""
    payload = json.dumps(existing, indent=2)
    if CLOUD_RUN_SERVICE:
        from google.cloud import secretmanager
        import google.auth

        client = secretmanager.SecretManagerServiceClient()
        project = os.environ.get("GOOGLE_CLOUD_PROJECT") or os.environ.get("PROJECT_ID")
        if not project:
            _, project = google.auth.default()
        parent = f"projects/{project}/secrets/{GME_SECRETS_NAME}"
        client.add_secret_version(request={"parent": parent, "payload": {"data": payload.encode()}})
    else:
        GME_SECRETS_PATH.write_text(payload)

REPORTS = {
    "thailand": {"label": "Thailand", "png": PROJECT_ROOT / "Thailand_left_table.png"},
    "laos": {"label": "Laos", "png": PROJECT_ROOT / "Laos_left_table.png"},
}

app = Flask(__name__, static_folder=None)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=bool(CLOUD_RUN_SERVICE),
)

# --------------------------------------------------------------------------
# Auth - two-tier PIN gate, same scheme as the laptop dashboard.
# --------------------------------------------------------------------------


def _hash_pin(pin: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", pin.encode(), salt.encode(), 100_000).hex()


def _verify_pin(pin: str, salt: str, expected_hash: str) -> bool:
    if secrets.compare_digest(_hash_pin(pin, salt), expected_hash):
        return True
    legacy_hash = hashlib.sha256((salt + pin).encode()).hexdigest()
    return secrets.compare_digest(legacy_hash, expected_hash)


def _generate_pin() -> str:
    return f"{secrets.randbelow(1_000_000):06d}"


def _load_or_create_auth() -> dict:
    if AUTH_PATH.exists():
        return json.loads(AUTH_PATH.read_text())

    team_pin = _generate_pin()
    owner_pin = _generate_pin()
    team_salt = secrets.token_hex(16)
    owner_salt = secrets.token_hex(16)
    data = {
        "team_pin_hash": _hash_pin(team_pin, team_salt),
        "team_pin_salt": team_salt,
        "owner_pin_hash": _hash_pin(owner_pin, owner_salt),
        "owner_pin_salt": owner_salt,
    }
    AUTH_PATH.write_text(json.dumps(data, indent=2))
    print("=" * 64)
    print("Generated new dashboard PINs - each is shown only this once.")
    print(f"  Team PIN  (share with teammates): {team_pin}")
    print(f"  Owner PIN (keep private, yours only): {owner_pin}")
    print(f"Both are stored (hashed, salted) in {AUTH_PATH.name}.")
    print("To regenerate either, delete that file and restart the server.")
    print("=" * 64)
    return data


def _load_or_create_secret_key() -> str:
    if SECRET_KEY_PATH.exists():
        return SECRET_KEY_PATH.read_text().strip()
    key = secrets.token_hex(32)
    SECRET_KEY_PATH.write_text(key)
    return key


AUTH = _load_or_create_auth()
app.secret_key = _load_or_create_secret_key()

_login_attempts = {}
_login_attempts_lock = threading.Lock()
LOGIN_MAX_ATTEMPTS = 5
LOGIN_WINDOW_SECONDS = 300

OPEN_PATHS = {"/login", "/api/cron/run"}


def _is_locked_out(ip: str) -> bool:
    with _login_attempts_lock:
        cutoff = time.time() - LOGIN_WINDOW_SECONDS
        attempts = [t for t in _login_attempts.get(ip, []) if t > cutoff]
        _login_attempts[ip] = attempts
        return len(attempts) >= LOGIN_MAX_ATTEMPTS


def _record_failed_attempt(ip: str):
    with _login_attempts_lock:
        _login_attempts.setdefault(ip, []).append(time.time())


@app.before_request
def _require_auth():
    if request.path in OPEN_PATHS or request.path.startswith("/static/"):
        return None
    if session.get("authenticated"):
        return None
    if request.path.startswith("/api/"):
        return jsonify({"error": "Not authenticated"}), 401
    return redirect("/login")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return send_file(str(PROJECT_ROOT / "static" / "login.html"))

    ip = request.remote_addr
    if _is_locked_out(ip):
        return jsonify({"error": "Too many failed attempts - locked out for 5 minutes"}), 429

    data = request.get_json(force=True, silent=True) or {}
    pin = str(data.get("pin", "")).strip()

    if pin and _verify_pin(pin, AUTH["owner_pin_salt"], AUTH["owner_pin_hash"]):
        session["authenticated"] = True
        session["role"] = "owner"
        return jsonify({"ok": True, "role": "owner"})
    if pin and _verify_pin(pin, AUTH["team_pin_salt"], AUTH["team_pin_hash"]):
        session["authenticated"] = True
        session["role"] = "team"
        return jsonify({"ok": True, "role": "team"})

    _record_failed_attempt(ip)
    return jsonify({"error": "Incorrect PIN"}), 401


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/whoami")
def whoami():
    return jsonify({"role": session.get("role", "none")})


# --------------------------------------------------------------------------
# GME Laos mobile-API token - owner-only. No automatable refresh exists (see
# laos-rate-report's references/gme-mobile-token.md - it needs the account
# password to automate, deliberately rejected), so this is a permanent
# human-in-the-loop step; this just gives that human a form instead of a
# terminal. Local file for now, same as the laptop dashboard - once this
# moves to Cloud Run, this needs to write a new Secret Manager version
# instead (mounted secrets are read-only in the container), not yet built.
# --------------------------------------------------------------------------


def _decode_jwt_exp(bearer_token: str) -> datetime:
    if not bearer_token.startswith("Bearer "):
        raise ValueError("must start with 'Bearer '")
    jwt = bearer_token.removeprefix("Bearer ").strip()
    payload_b64 = jwt.split(".")[1]
    padded = payload_b64 + "=" * (-len(payload_b64) % 4)
    claims = json.loads(base64.urlsafe_b64decode(padded))
    return datetime.fromtimestamp(claims["exp"], tz=timezone.utc)


@app.route("/api/gme-token-status")
def api_gme_token_status():
    if session.get("role") != "owner":
        return jsonify({"error": "Owner access required"}), 403

    if not GME_SECRETS_PATH.exists():
        return jsonify({"lastUpdatedAt": None, "expiresAt": None, "hasCookie": False, "hasUsername": False})

    last_updated = datetime.fromtimestamp(GME_SECRETS_PATH.stat().st_mtime, tz=timezone.utc)
    expires_at = None
    existing = {}
    try:
        existing = json.loads(GME_SECRETS_PATH.read_text())
        expires_at = _decode_jwt_exp(existing.get("authorization", "")).isoformat()
    except Exception:  # noqa: BLE001 - malformed/missing token, just omit expiresAt
        pass

    return jsonify({
        "lastUpdatedAt": last_updated.isoformat(),
        "expiresAt": expires_at,
        "hasCookie": bool(existing.get("cookie")),
        "hasUsername": bool(existing.get("username")),
    })


@app.route("/api/gme-token", methods=["POST"])
def api_gme_token():
    """Updates just the Bearer token - cookie/username don't change between
    refreshes (see references/gme-mobile-token.md) and must already exist on
    disk from a one-time initial setup; this form only ever refreshes the
    token on top of that. Then makes one live GME mobile API call to confirm
    it actually works, not just that it decodes - a plain read-only quote
    lookup, not the webhook."""
    if session.get("role") != "owner":
        return jsonify({"error": "Owner access required"}), 403

    data = request.get_json(force=True, silent=True) or {}
    token = str(data.get("authorization", "")).strip()

    try:
        expiry = _decode_jwt_exp(token)
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"Couldn't decode token: {e}"}), 400

    existing = json.loads(GME_SECRETS_PATH.read_text()) if GME_SECRETS_PATH.exists() else {}
    missing = [k for k in ("cookie", "username") if not existing.get(k)]
    if missing:
        return jsonify({
            "error": f"{' and '.join(missing)} not set up yet on this instance - this form only refreshes the "
                     f"token, it needs cookie/username seeded once first. Ask the assistant to seed them."
        }), 400

    existing["authorization"] = token
    _persist_gme_secrets(existing)

    try:
        corridor = laos_fetch.GME_MOBILE_CORRIDORS[("GME (RIA-Other banks)", "LAK")]
        result = laos_fetch.fetch_gme_mobile_laos(corridor, existing)
        test_result = {"ok": True, "krw": result["krw"], "fee": result["fee"]}
    except Exception as e:  # noqa: BLE001
        test_result = {"ok": False, "message": str(e)}

    return jsonify({"ok": True, "expiresAt": expiry.isoformat(), "test": test_result})


# --------------------------------------------------------------------------
# Rates / charts
# --------------------------------------------------------------------------


@app.route("/")
def index():
    return send_file(str(PROJECT_ROOT / "static" / "index.html"))


@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(PROJECT_ROOT / "static", filename)


@app.route("/api/overview")
def api_overview():
    db_path = gcs_sync.download_db()
    out = {}
    for key, cfg in REPORTS.items():
        has_png = gcs_sync.download_png(key, cfg["png"])
        out[key] = {
            "label": cfg["label"],
            "status": gcs_sync.read_status(key),
            "rates": ranking.latest_rates(key, db_path),
            "hasPng": has_png,
            "pngUpdatedAt": cfg["png"].stat().st_mtime if has_png and cfg["png"].exists() else None,
        }
    resp = jsonify(out)
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/image/<report>")
def api_image(report):
    cfg = REPORTS.get(report)
    if not cfg:
        abort(404)
    if not gcs_sync.download_png(report, cfg["png"]) or not cfg["png"].exists():
        abort(404)
    resp = send_file(str(cfg["png"]), mimetype="image/png")
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/api/history/<report>")
def api_history(report):
    if report not in REPORTS:
        abort(404)
    db_path = gcs_sync.download_db()
    resp = jsonify(ranking.weekly_history(report, db_path))
    resp.headers["Cache-Control"] = "no-store"
    return resp


# --------------------------------------------------------------------------
# Refresh (fetch + rank + log + render PNG). Never posts to Teams even
# though pipeline.run() supports it - a manual dashboard click is a preview,
# same "don't post what wasn't reviewed" reasoning the laptop dashboard's
# separate Refresh/Send-to-Workflow buttons used. Only /api/cron/run's
# scheduled runs pass post_to_teams=True.
# --------------------------------------------------------------------------

GME_MANUAL_KRW_BOUNDS = (26_000 * 35, 26_000 * 55)


def _run_job(job_id: str, report: str, gme_manual_krw=None):
    try:
        pipeline.run(report, gme_manual_krw=gme_manual_krw, log_history=True)
    except Exception:  # noqa: BLE001 - status is already recorded by pipeline.run itself
        pass


@app.route("/api/action", methods=["POST"])
def api_action():
    data = request.get_json(force=True, silent=True) or {}
    report = data.get("report")
    if report not in REPORTS:
        return jsonify({"error": f"Unknown report '{report}'"}), 400

    gme_manual_krw = None
    if data.get("gmeManualKrw") not in (None, ""):
        if report != "thailand":
            return jsonify({"error": "gmeManualKrw only applies to Thailand"}), 400
        try:
            gme_manual_krw = int(data["gmeManualKrw"])
        except (TypeError, ValueError):
            return jsonify({"error": "gmeManualKrw must be a whole number"}), 400
        lo, hi = GME_MANUAL_KRW_BOUNDS
        if not (lo <= gme_manual_krw <= hi):
            return jsonify({"error": f"gmeManualKrw ({gme_manual_krw:,}) looks like a typo - expected roughly {lo:,}-{hi:,}"}), 400

    job_id = f"{report}-{uuid.uuid4().hex[:8]}"
    threading.Thread(target=_run_job, args=(job_id, report, gme_manual_krw), daemon=True).start()
    return jsonify({"jobId": job_id, "report": report})


@app.route("/api/job/<job_id>")
def api_job(job_id):
    report = job_id.split("-")[0]
    if report not in REPORTS:
        abort(404)
    return jsonify(gcs_sync.read_status(report))


def _send_to_teams_job(report):
    try:
        pipeline.send_to_teams(report)
    except Exception:  # noqa: BLE001 - status is already recorded by pipeline.send_to_teams
        pass


@app.route("/api/send-to-teams", methods=["POST"])
def api_send_to_teams():
    data = request.get_json(force=True, silent=True) or {}
    report = data.get("report")
    if report not in REPORTS:
        return jsonify({"error": f"Unknown report '{report}'"}), 400

    job_id = f"{report}-{uuid.uuid4().hex[:8]}"
    threading.Thread(target=_send_to_teams_job, args=(report,), daemon=True).start()
    return jsonify({"jobId": job_id, "report": report})


@app.route("/api/cron/run", methods=["POST", "GET"])
def api_cron_run():
    cron_key = request.headers.get("X-Cron-Key") or request.args.get("key")
    expected_key = os.environ.get("CRON_SECRET", app.secret_key)
    is_scheduler = bool(request.headers.get("X-CloudScheduler") or request.headers.get("X-Google-CloudScheduler"))
    is_authed = session.get("authenticated") or is_scheduler or (cron_key and secrets.compare_digest(cron_key, expected_key))
    if not is_authed:
        return jsonify({"error": "Unauthorized"}), 401

    data = request.get_json(force=True, silent=True) or {}
    action = request.args.get("action") or data.get("action", "thailand")

    if action == "thailand":
        res = pipeline.run("thailand", log_history=True, post_to_teams=True)
        return jsonify({"ok": True, "action": "thailand", "result": res})
    elif action == "laos":
        res = pipeline.run("laos", log_history=True, post_to_teams=True)
        return jsonify({"ok": True, "action": "laos", "result": res})
    elif action in ("fees", "competitor-fees", "sanity-check"):
        import check_competitor_fees

        res = check_competitor_fees.run_checks()
        return jsonify({"ok": True, "action": "fees", "result": res})
    else:
        return jsonify({"error": f"Unknown action '{action}'"}), 400


if __name__ == "__main__":
    print("Dashboard running at http://127.0.0.1:5151")
    app.run(host="0.0.0.0", port=5151, debug=False, threaded=True)
