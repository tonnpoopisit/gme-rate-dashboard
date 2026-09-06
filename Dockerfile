# Playwright's base image bundles Chromium + every system lib it needs -
# sidesteps the Windows-only browser-cache-virtualization problem the
# laptop pipeline had to work around (see automation-scheduling.md).
FROM mcr.microsoft.com/playwright/python:v1.47.0-jammy

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=8080
EXPOSE 8080
CMD ["gunicorn", "-b", "0.0.0.0:8080", "--workers", "1", "--threads", "8", "--timeout", "180", "dashboard_server:app"]
