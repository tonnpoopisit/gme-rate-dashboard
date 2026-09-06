#!/usr/bin/env bash
# Refreshes the copied fetcher source from claude-skills/pipeline-scripts
# (the single source of truth - fix a broken provider fetcher there, then
# rerun this before the next deploy) and builds+deploys the Cloud Run
# service. Requires: gcloud CLI authenticated, PROJECT_ID/REGION/GCS_BUCKET
# set below or as env vars.
set -euo pipefail

SKILLS_DIR="${SKILLS_DIR:-$HOME/Documents/claude-skills/pipeline-scripts}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

for f in fetch_thailand_rates.py fetch_laos_rates.py fetch_kebhana_rates.py known_fees.py price_history.py check_competitor_fees.py; do
  [ -f "$SKILLS_DIR/$f" ] && cp "$SKILLS_DIR/$f" "$HERE/$f" || true
done
echo "Copied latest fetcher source from $SKILLS_DIR"

SERVICE_NAME="${SERVICE_NAME:-gme-rate-dashboard}"
PROJECT_ID="${PROJECT_ID:-gme-related-project}"
REGION="${REGION:-asia-southeast1}"
GCS_BUCKET="${GCS_BUCKET:-gme-related-project-rate-data}"

gcloud run deploy "$SERVICE_NAME" \
  --source "$HERE" \
  --project "$PROJECT_ID" \
  --region "$REGION" \
  --allow-unauthenticated \
  --no-cpu-throttling \
  --memory 2Gi \
  --cpu 2 \
  --set-env-vars "GCS_BUCKET=$GCS_BUCKET,GOOGLE_CLOUD_PROJECT=$PROJECT_ID,AUTH_PATH=/app/secrets/auth/dashboard_auth.json,SECRET_KEY_PATH=/app/secrets/key/dashboard_secret_key.txt,GME_SECRETS_PATH=/app/secrets/gme/gme_mobile_secrets.json,GME_SECRETS_SECRET_NAME=gme-mobile-secrets,TEAMS_WEBHOOK_THAILAND_PATH=/app/secrets/webhook-th/webhook_config.json,TEAMS_WEBHOOK_LAOS_PATH=/app/secrets/webhook-lo/webhook_config_laos.json" \
  --set-secrets "/app/secrets/auth/dashboard_auth.json=dashboard-auth-json:latest,/app/secrets/key/dashboard_secret_key.txt=dashboard-secret-key:latest,/app/secrets/gme/gme_mobile_secrets.json=gme-mobile-secrets:latest,/app/secrets/webhook-th/webhook_config.json=teams-webhook-thailand:latest,/app/secrets/webhook-lo/webhook_config_laos.json=teams-webhook-laos:latest"
