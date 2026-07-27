"""MCP server: SEC EDGAR filings + XBRL facts for Cursor Cloud Agents."""

from __future__ import annotations

import json
import os
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import providers

HOST = os.getenv("MCP_HOST", "0.0.0.0")
PORT = int(os.getenv("MCP_PORT", "8000"))

mcp = FastMCP(
    "sec-edgar-mcp",
    instructions=(
        "Tools to access SEC EDGAR company filings and XBRL financial facts "
        "via data.sec.gov (free, no API key). Always set SEC_USER_AGENT with "
        "an app name and contact email. Forms: 10-K, 10-Q, 8-K, 20-F, etc."
    ),
    host=HOST,
    port=PORT,
    stateless_http=True,
    json_response=True,
)


def _ok(data: dict[str, Any]) -> str:
    # Avoid dumping huge nested blobs from submissions helper
    if "_data" in data:
        data = {k: v for k, v in data.items() if k != "_data"}
    return json.dumps(data, indent=2, default=str)


def _err(exc: Exception) -> str:
    return json.dumps({"error": str(exc)}, indent=2)


@mcp.tool()
def lookup_ticker(ticker: str) -> str:
    """Resolve a stock ticker (e.g. AAPL, PBR, KO) to SEC CIK and company title."""
    try:
        return _ok(providers.lookup_ticker(ticker))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def list_filings(ticker_or_cik: str, form: str = "", limit: int = 20) -> str:
    """List recent SEC filings for a ticker or CIK.

    form examples: 10-K, 10-Q, 8-K, 20-F, 6-K (empty = all forms).
    Returns documentUrl / indexUrl for each filing.
    """
    try:
        return _ok(
            providers.list_filings(
                ticker_or_cik,
                form=form or None,
                limit=limit,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_company_facts(
    ticker_or_cik: str,
    concepts: str = "",
    limit_per_concept: int = 8,
) -> str:
    """Get key XBRL financial facts (revenues, net income, assets, etc.).

    concepts: optional comma-separated us-gaap tags
    (e.g. 'NetIncomeLoss,Assets'). Empty = default set.
    """
    try:
        concept_list = (
            [c.strip() for c in concepts.split(",") if c.strip()]
            if concepts
            else None
        )
        return _ok(
            providers.get_company_facts(
                ticker_or_cik,
                concepts=concept_list,
                limit_per_concept=limit_per_concept,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_concept(
    ticker_or_cik: str,
    concept: str,
    taxonomy: str = "us-gaap",
    limit: int = 20,
) -> str:
    """Time series for one XBRL concept (e.g. NetIncomeLoss, Revenues)."""
    try:
        return _ok(
            providers.get_concept(
                ticker_or_cik,
                concept=concept,
                taxonomy=taxonomy,
                limit=limit,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_company_profile(ticker_or_cik: str) -> str:
    """Company profile from SEC submissions (name, SIC, tickers, exchanges)."""
    try:
        meta = providers.get_submissions(ticker_or_cik)
        meta.pop("_data", None)
        meta.pop("raw_keys", None)
        return _ok(meta)
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
