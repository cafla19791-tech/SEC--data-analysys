#!/usr/bin/env bash
# Idempotent Cloud Agent setup for the SEC--data-analysys project.
# Creates a virtualenv, installs pinned dependencies, and generates the
# sample outputs the Streamlit dashboard reads on first load.
set -euo pipefail

cd "$(dirname "$0")/.."

# The Cursor default image ships Python 3.12 but not the venv module.
if ! python3 -c "import ensurepip" >/dev/null 2>&1; then
  sudo apt-get update -qq
  sudo apt-get install -y --no-install-recommends python3-venv
fi

if [ ! -x .venv/bin/python ]; then
  python3 -m venv .venv
fi

.venv/bin/python -m pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

# Pre-generate the sample flows + agent ranking so `streamlit run app.py`
# shows real data immediately (uses the committed sample + local SELIC file,
# so no network access is required).
PYTHONPATH=. .venv/bin/python scripts/gerar_fluxos.py \
  --input data/sample_operacoes_com_agente.csv --stem fluxos_amostra
