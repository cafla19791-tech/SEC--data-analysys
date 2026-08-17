#!/usr/bin/env bash
# Idempotent Cloud Agent setup for the SEC--data-analysys project.
# Runs from the repository root after the source tree is checked out.
set -euo pipefail

# The default image ships Python 3.12 but not the venv module.
if ! python3 -c "import ensurepip" >/dev/null 2>&1; then
  sudo apt-get update -qq
  sudo apt-get install -y -qq python3.12-venv
fi

# Create the virtualenv only when it does not already exist.
if [ ! -x ".venv/bin/python" ]; then
  python3 -m venv .venv
fi

./.venv/bin/python -m pip install --upgrade pip
./.venv/bin/python -m pip install -r requirements.txt

# Generate the sample outputs the Streamlit dashboard reads on boot.
# Deterministic, terminates, and safe to re-run.
./.venv/bin/python scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv

echo "install.sh completed"
