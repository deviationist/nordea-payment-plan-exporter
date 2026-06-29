#!/usr/bin/env bash
# Compare the initial plan vs the reduced (post-1,1 M) plan -> CSV + XLSX.
#   ./compare.sh                 # estimate (re-annuitised minus the downpayment)
#   ./compare.sh --post <ts>     # hard comparison vs a captured reduced plan
#   ./compare.sh --downpayment 500000 --format xlsx
# Reuses the build venv (openpyxl); no Nordea login.
set -euo pipefail
cd "$(dirname "$0")"

VENV=".venv"
PY="$VENV/bin/python"
if [[ ! -x "$PY" ]]; then
  python3 -m venv "$VENV"
  "$VENV/bin/pip" install --quiet --upgrade pip
  "$VENV/bin/pip" install --quiet -r requirements.txt
fi
exec "$PY" compare.py "$@"
