# nyse-mcp

Skeleton MCP server for **US equity market data** (NYSE/NASDAQ tickers), made to work **without Cursor Desktop**.

Default provider: **Yahoo Finance** (free, delayed, no API key).  
Optional: **Alpha Vantage** (`MARKET_DATA_PROVIDER=alphavantage` + `ALPHA_VANTAGE_API_KEY`).

> Market **data** only — does not place trades.

## Tools

| Tool | Purpose |
|------|---------|
| `get_quote` | Latest delayed quote |
| `get_history` | OHLCV bars |
| `get_fundamentals` | Market cap, P/E, sector, etc. |
| `search_ticker` | Find ticker by company name |
| `market_status` | Indicative US session status |

## Corporate / no local Cursor

If you cannot install Cursor Desktop, follow:

- **[CADASTRO_MCP_CLOUD.md](./CADASTRO_MCP_CLOUD.md)** — exact Cloud Agent MCP fields
- **[CLOUD_SETUP.md](./CLOUD_SETUP.md)** — overview + Windows CLI

Short version:

1. Use [cursor.com/agents](https://cursor.com/agents)
2. Register `nyse-mcp` in the **MCP** dropdown (stdio → `bash ./nyse-mcp/run_mcp.sh`), **or**
3. Skip MCP and use the CLI in the Cloud Agent VM / Windows:

```bash
# Cloud Agent
cd nyse-mcp && python3 -m venv .venv && source .venv/bin/activate && pip install -e .
nyse-mcp-cli quote JPM
```

```powershell
# Windows (VS Code / ContAgil WinPython) — no Cursor Desktop needed
cd nyse-mcp
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_e_rodar_cli.ps1 -Symbol JPM
.\nyse_mcp_cli.bat quote XOM
```

## Local install (optional)

```bash
cd nyse-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Transports

| Mode | Env | Use case |
|------|-----|----------|
| `stdio` (default) | `MCP_TRANSPORT=stdio` | Cloud Agent / IDE process |
| `streamable-http` | `MCP_TRANSPORT=streamable-http` | Remote HTTP MCP (`/mcp`) |

```bash
# stdio (default)
bash ./run_mcp.sh

# HTTP
MCP_TRANSPORT=streamable-http MCP_PORT=8000 bash ./run_mcp.sh
```

## Alpha Vantage (optional)

```bash
export MARKET_DATA_PROVIDER=alphavantage
export ALPHA_VANTAGE_API_KEY=YOUR_KEY
```

## Limits

- Yahoo/Alpha Vantage data is typically **delayed**, not a direct NYSE feed.
- Free Alpha Vantage tiers have strict rate limits.
- `market_status` is indicative only.
