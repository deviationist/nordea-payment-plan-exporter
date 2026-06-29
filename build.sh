#!/usr/bin/env bash
# Build the Nordea payment-plan spreadsheet(s) from the raw JSON sources.
#
#   ./build.sh                    # both CSV and XLSX (committed baseline)
#   ./build.sh csv                # CSV only
#   ./build.sh xlsx               # XLSX only
#   ./build.sh test               # run the integration tests against the bank's plan
#   ./build.sh --capture <ts>     # build from a captures/ snapshot (see fetch.sh)
#   ./build.sh xlsx --capture <ts>
#
# Creates/uses a local .venv automatically — no manual setup, no Nordea login.
set -euo pipefail
cd "$(dirname "$0")"

VENV=".venv"
PY="$VENV/bin/python"

if [[ ! -x "$PY" ]]; then
  echo "Setting up $VENV ..."
  python3 -m venv "$VENV"
  "$VENV/bin/pip" install --quiet --upgrade pip
  "$VENV/bin/pip" install --quiet -r requirements.txt
fi

if [[ "${1:-}" == "test" ]]; then
  exec "$PY" -m unittest -v test_payplan test_compare
fi
FORMAT="both"
case "${1:-}" in
  csv|xlsx|both) FORMAT="$1"; shift ;;
esac
exec "$PY" build_sheet.py --format "$FORMAT" "$@"
