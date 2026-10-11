#!/usr/bin/env bash
# Idempotent Cloud Agent bootstrap for SEC--data-analysys.
# Prepares a Python 3.12 venv, installs pinned dependencies, and generates the
# sample outputs the Streamlit dashboard reads.
set -euo pipefail

cd "$(dirname "$0")/.."

# ensurepip / venv support is not in the base Debian/Ubuntu python by default.
if ! python3 -c "import ensurepip" >/dev/null 2>&1; then
  sudo apt-get update -qq
  sudo apt-get install -y -qq python3-venv
fi

if [ ! -x .venv/bin/python ]; then
  python3 -m venv .venv
fi

.venv/bin/python -m pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

# Regenerate the sample ranking so the dashboard has fresh data on first boot.
PYTHONPATH=. .venv/bin/python scripts/gerar_fluxos.py \
  --input data/sample_operacoes_com_agente.csv --stem fluxos_amostra

echo "[install] SEC--data-analysys environment ready."
