#!/bin/bash
set -e

MAX_DOWNLOAD_BYTES="${MAX_DOWNLOAD_BYTES:-536870912}"
JSON_CLI="/var/www/onlyoffice/documentserver/npm/json"
DEFAULT_JSON="/etc/onlyoffice/documentserver/default.json"
LOCAL_JSON="/etc/onlyoffice/documentserver/local.json"
RUNTIME_JSON="/var/www/onlyoffice/Data/runtime.json"

echo "[onlyoffice-bootstrap] Starting ONLYOFFICE Document Server..."
/app/ds/run-document-server.sh &
main_pid=$!

trap 'kill -TERM "$main_pid" 2>/dev/null || true' TERM INT

for _ in $(seq 1 120); do
  if [ -x "$JSON_CLI" ] && [ -f "$DEFAULT_JSON" ]; then
    break
  fi
  sleep 1
done

if [ -x "$JSON_CLI" ] && [ -f "$DEFAULT_JSON" ]; then
  echo "[onlyoffice-bootstrap] Setting maxDownloadBytes=${MAX_DOWNLOAD_BYTES}..."
  "$JSON_CLI" -I -f "$DEFAULT_JSON" -e "this.FileConverter.converter.maxDownloadBytes=${MAX_DOWNLOAD_BYTES}" || true
  "$JSON_CLI" -I -f "$DEFAULT_JSON" -e "this.services.CoAuthoring.server.limits_tempfile_upload=${MAX_DOWNLOAD_BYTES}" || true

  if [ ! -f "$LOCAL_JSON" ]; then
    echo "{}" > "$LOCAL_JSON"
  fi
  "$JSON_CLI" -I -f "$LOCAL_JSON" -e "this.FileConverter=this.FileConverter||{};this.FileConverter.converter=this.FileConverter.converter||{};this.FileConverter.converter.maxDownloadBytes=${MAX_DOWNLOAD_BYTES}" || true
  "$JSON_CLI" -I -f "$LOCAL_JSON" -e "this.services=this.services||{};this.services.CoAuthoring=this.services.CoAuthoring||{};this.services.CoAuthoring.server=this.services.CoAuthoring.server||{};this.services.CoAuthoring.server.limits_tempfile_upload=${MAX_DOWNLOAD_BYTES}" || true

  if [ -f "$RUNTIME_JSON" ]; then
    "$JSON_CLI" -I -f "$RUNTIME_JSON" -e "this.FileConverter=this.FileConverter||{};this.FileConverter.converter=this.FileConverter.converter||{};this.FileConverter.converter.maxDownloadBytes=${MAX_DOWNLOAD_BYTES}" || true
    "$JSON_CLI" -I -f "$RUNTIME_JSON" -e "this.services=this.services||{};this.services.CoAuthoring=this.services.CoAuthoring||{};this.services.CoAuthoring.server=this.services.CoAuthoring.server||{};this.services.CoAuthoring.server.limits_tempfile_upload=${MAX_DOWNLOAD_BYTES}" || true
  fi
fi

if command -v supervisorctl >/dev/null 2>&1; then
  for _ in $(seq 1 60); do
    if supervisorctl status >/dev/null 2>&1; then
      echo "[onlyoffice-bootstrap] Reloading ONLYOFFICE services..."
      supervisorctl restart all >/dev/null 2>&1 || true
      break
    fi
    sleep 1
  done
fi

echo "[onlyoffice-bootstrap] Effective settings:"
grep -n "maxDownloadBytes\\|limits_tempfile_upload" "$DEFAULT_JSON" "$LOCAL_JSON" "$RUNTIME_JSON" 2>/dev/null || true

wait "$main_pid"
