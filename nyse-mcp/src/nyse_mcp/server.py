"""MCP server entrypoint: exposes NYSE/US equity tools to Cursor Cloud Agents."""

from __future__ import annotations

import json
import os
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import providers

HOST = os.getenv("MCP_HOST", "0.0.0.0")
PORT = int(os.getenv("MCP_PORT", "8000"))

mcp = FastMCP(
    "nyse-mcp",
    instructions=(
        "US equity market-data tools for NYSE/NASDAQ tickers. "
        "Default provider is Yahoo Finance (free, delayed). "
        "Set MARKET_DATA_PROVIDER=alphavantage and ALPHA_VANTAGE_API_KEY for Alpha Vantage."
    ),
    host=HOST,
    port=PORT,
    # Recommended for remote/cloud HTTP MCP clients.
    stateless_http=True,
    json_response=True,
)


def _ok(data: dict[str, Any]) -> str:
    return json.dumps(data, indent=2, default=str)


def _err(exc: Exception) -> str:
    return json.dumps({"error": str(exc)}, indent=2)


@mcp.tool()
def get_quote(symbol: str) -> str:
    """Get the latest delayed quote for a US equity ticker (e.g. AAPL, JPM, XOM)."""
    try:
        return _ok(providers.get_quote(symbol))
    except Exception as exc:  # noqa: BLE001 - surface provider errors to the model
        return _err(exc)


@mcp.tool()
def get_history(symbol: str, period: str = "1mo", interval: str = "1d") -> str:
    """Get OHLCV history for a ticker.

    period examples: 5d, 1mo, 3mo, 6mo, 1y, 5y, max
    interval examples: 1d, 1wk, 1mo (Yahoo); Alpha Vantage free tier prefers daily.
    """
    try:
        return _ok(providers.get_history(symbol, period=period, interval=interval))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_fundamentals(symbol: str) -> str:
    """Get basic fundamentals: market cap, P/E, sector, dividend yield, 52w range."""
    try:
        return _ok(providers.get_fundamentals(symbol))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def search_ticker(query: str) -> str:
    """Search tickers/company names (Yahoo Finance search)."""
    try:
        return _ok(providers.search_ticker(query))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def market_status() -> str:
    """Return an indicative US market open/closed status (not an official NYSE feed)."""
    try:
        return _ok(providers.market_status())
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


def main() -> None:
    """Run MCP over stdio (Cloud Agent / IDE) or streamable-http (remote)."""
    transport = os.getenv("MCP_TRANSPORT", "stdio").strip().lower()
    if transport in {"http", "streamable-http", "streamable_http"}:
        mcp.run(transport="streamable-http")
        return
    if transport == "sse":
        # Kept for local experiments; Cursor Cloud Agents do not support SSE.
        mcp.run(transport="sse")
        return
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
