#!/usr/bin/env bash
# Idempotent Cloud Agent setup for SEC--data-analysys.
# Recurring builds clone main; keep this file on the default branch
# if the dashboard install command is `bash .cursor/install.sh`.
set -euo pipefail

cd "$(dirname "$0")/.."

if ! python3 -c "import ensurepip" >/dev/null 2>&1; then
  sudo apt-get update -qq
  sudo apt-get install -y -qq python3.12-venv
fi

if [ ! -x ".venv/bin/python" ]; then
  python3 -m venv .venv
fi

./.venv/bin/python -m pip install --upgrade pip
./.venv/bin/python -m pip install -r requirements.txt

# Sample outputs the Streamlit dashboard reads on first boot.
./.venv/bin/python scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv

echo "install.sh completed"
