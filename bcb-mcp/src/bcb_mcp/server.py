"""MCP server: Banco Central do Brasil SGS + OLINDA for Cursor Cloud Agents."""

from __future__ import annotations

import json
import os
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import providers

HOST = os.getenv("MCP_HOST", "0.0.0.0")
PORT = int(os.getenv("MCP_PORT", "8000"))

mcp = FastMCP(
    "bcb-mcp",
    instructions=(
        "Tools for Banco Central do Brasil open data: SGS time series "
        "(Selic, CDI, IPCA, FX) and OLINDA (PTAX, Focus expectations). "
        "No API key. Prefer aliases (selic, ipca, dolar) or numeric SGS codes."
    ),
    host=HOST,
    port=PORT,
    stateless_http=True,
    json_response=True,
)


def _ok(data: dict[str, Any]) -> str:
    return json.dumps(data, indent=2, default=str, ensure_ascii=True)


def _err(exc: Exception) -> str:
    return json.dumps({"error": str(exc)}, indent=2, ensure_ascii=True)


@mcp.tool()
def list_known_series() -> str:
    """List local aliases for common BCB SGS series (Selic, CDI, IPCA, FX...)."""
    try:
        return _ok(providers.list_known_series())
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_sgs_series(
    code_or_alias: str,
    date_from: str = "",
    date_to: str = "",
    last: int = 0,
) -> str:
    """Fetch a BCB SGS time series by alias (selic, ipca, dolar) or numeric code.

    date_from / date_to: YYYY-MM-DD or DD/MM/YYYY. If last > 0, returns last N points.
    Windows longer than ~10 years are auto-chunked (BCB limit).
    """
    try:
        return _ok(
            providers.get_sgs_series(
                code_or_alias,
                date_from=date_from or None,
                date_to=date_to or None,
                last=last if last and last > 0 else None,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_ptax(date_from: str = "", date_to: str = "", last_days: int = 30) -> str:
    """USD/BRL PTAX quotes from OLINDA (buy/sell, bulletin type)."""
    try:
        return _ok(
            providers.get_ptax(
                date_from=date_from or None,
                date_to=date_to or None,
                last_days=last_days,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_expectativas(indicator: str = "IPCA", top: int = 20) -> str:
    """Market Focus expectations (annual) from OLINDA. indicator: IPCA, Selic, Cambio..."""
    try:
        return _ok(providers.get_expectativas(indicator, top=top))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


def main() -> None:
    transport = os.getenv("MCP_TRANSPORT", "stdio").strip().lower()
    if transport in {"http", "streamable-http", "streamable_http"}:
        mcp.run(transport="streamable-http")
        return
    if transport == "sse":
        mcp.run(transport="sse")
        return
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
