#!/usr/bin/env bash
# Launches the Streamlit dashboard for SEC--data-analysys.
set -euo pipefail

cd "$(dirname "$0")/.."

export PYTHONPATH=.
exec .venv/bin/streamlit run app.py \
  --server.port 8501 \
  --server.address 0.0.0.0 \
  --server.headless true
