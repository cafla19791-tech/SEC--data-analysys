#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"

if [[ -x "$ROOT/.venv/bin/python" ]]; then
  PYTHON="$ROOT/.venv/bin/python"
else
  PYTHON="${PYTHON:-python3}"
fi

export PYTHONPATH="$ROOT/src${PYTHONPATH:+:$PYTHONPATH}"
export MARKET_DATA_PROVIDER="${MARKET_DATA_PROVIDER:-yahoo}"
# stdio = Cloud Agent / IDE local process
# streamable-http = remote HTTP MCP (MCP_HOST/MCP_PORT)
export MCP_TRANSPORT="${MCP_TRANSPORT:-stdio}"

exec "$PYTHON" -m nyse_mcp.server
