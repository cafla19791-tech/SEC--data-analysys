# nyse-mcp

Skeleton MCP server that gives Cursor tools for **US equity market data** (NYSE/NASDAQ tickers).

Default provider: **Yahoo Finance** (free, delayed, no API key).  
Optional provider: **Alpha Vantage** (free tier with API key).

> This is market **data** only — it does not place trades.

## Tools

| Tool | Purpose |
|------|---------|
| `get_quote` | Latest delayed quote |
| `get_history` | OHLCV bars |
| `get_fundamentals` | Market cap, P/E, sector, etc. |
| `search_ticker` | Find ticker by company name |
| `market_status` | Indicative US session status |

## Setup

```bash
cd nyse-mcp
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e .
```

## Cursor config

Project config is already at [`.cursor/mcp.json`](../.cursor/mcp.json).

1. Install deps (`pip install -e .` from `nyse-mcp/`).
2. Restart Cursor (or reload MCP servers in **Settings → Tools & MCP**).
3. Ask things like: *“Qual a cotação da JPM?”* / *“Histórico de 1 ano da XOM”*.

### Alpha Vantage (optional)

Edit `.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "nyse-mcp": {
      "command": "python",
      "args": ["-m", "nyse_mcp.server"],
      "cwd": "${workspaceFolder}/nyse-mcp/src",
      "env": {
        "MARKET_DATA_PROVIDER": "alphavantage",
        "ALPHA_VANTAGE_API_KEY": "YOUR_KEY_HERE",
        "PYTHONPATH": "${workspaceFolder}/nyse-mcp/src"
      }
    }
  }
}
```

If your shell uses a venv, prefer:

```json
"command": "${workspaceFolder}/nyse-mcp/.venv/bin/python"
```

## Manual smoke test

```bash
cd nyse-mcp
source .venv/bin/activate
PYTHONPATH=src python -c "from nyse_mcp.providers import get_quote; print(get_quote('AAPL'))"
```

## Limits / caveats

- Yahoo/Alpha Vantage data is typically **delayed**, not a direct NYSE feed.
- Free Alpha Vantage tiers have strict rate limits.
- `market_status` is indicative only.
