#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"

if [[ -x "$ROOT/.venv/bin/python" ]]; then
  PYTHON="$ROOT/.venv/bin/python"
else
  PYTHON="${PYTHON:-python3}"
fi

export PYTHONPATH="$ROOT/src${PYTHONPATH:+:$PYTHONPATH}"
export BCB_USER_AGENT="${BCB_USER_AGENT:-SEC-data-analysys-bcb-mcp/0.1 (cafla19791@gmail.com)}"
export MCP_TRANSPORT="${MCP_TRANSPORT:-stdio}"

exec "$PYTHON" -m bcb_mcp.server
