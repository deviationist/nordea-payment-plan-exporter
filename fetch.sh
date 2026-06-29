#!/usr/bin/env bash
# Refresh the raw Nordea JSON sources by driving a real login (Playwright),
# then rebuild. Reads SSN + BANKID_PWD from .env; you approve the BankID push
# on your phone. Direct API fetches 401, so the script intercepts the bank's
# own authenticated responses.
#
#   ./fetch.sh             # headless (default; the phone-push is the only manual step)
#   ./fetch.sh --headed    # visible window (fallback if Nordea ever blocks headless)
set -euo pipefail
cd "$(dirname "$0")"

VENV=".venv-fetch"
PY="$VENV/bin/python"

if [[ ! -x "$PY" ]]; then
  echo "Setting up $VENV (installs Playwright + Chromium, ~150 MB first time) ..."
  python3 -m venv "$VENV"
  "$VENV/bin/pip" install --quiet --upgrade pip
  "$VENV/bin/pip" install --quiet -r requirements-fetch.txt
  "$VENV/bin/playwright" install chromium
fi

exec "$PY" fetch.py "$@"
