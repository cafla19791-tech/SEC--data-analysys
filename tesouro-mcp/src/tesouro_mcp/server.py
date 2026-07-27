"""MCP server: Tesouro Nacional fiscal statistics for Cursor Cloud Agents."""

from __future__ import annotations

import json
import os
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import providers

HOST = os.getenv("MCP_HOST", "0.0.0.0")
PORT = int(os.getenv("MCP_PORT", "8000"))

mcp = FastMCP(
    "tesouro-mcp",
    instructions=(
        "Tools for Tesouro Nacional fiscal statistics: Resultado do Tesouro "
        "Nacional (RTN) monthly series via ARIA, headline Grandes Numeros, and "
        "CKAN open datasets. Values are typically in R$ milhoes. Themes: "
        "10=resultado fiscal, 13=investimento, 20=custeio."
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
def list_temas() -> str:
    """List RTN themes available in ARIA (10 fiscal, 13 investment, 20 admin costs)."""
    try:
        return _ok(providers.list_temas())
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def list_known_aliases() -> str:
    """List local aliases for common fiscal series (resultado_primario, receita_total...)."""
    try:
        return _ok(providers.list_known_aliases())
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def list_series(tema: str = "") -> str:
    """Catalog RTN series from Tesouro ARIA. Optional tema filter: 10, 13, or 20."""
    try:
        return _ok(providers.list_series(tema=tema or None))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def search_series(query: str, tema: str = "", limit: int = 30) -> str:
    """Search RTN series by name/code (e.g. 'primario', 'receita', '10.04')."""
    try:
        return _ok(providers.search_series(query, tema=tema or None, limit=limit))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_serie(
    alias_or_code: str,
    data_inicio: str = "",
    data_fim: str = "",
    correcao_ipca: bool = False,
) -> str:
    """Monthly fiscal series by alias (resultado_primario) or code (10.04.1).

    Dates: MM/AAAA or YYYY-MM. Values in R$ milhoes.
    """
    try:
        return _ok(
            providers.get_serie(
                alias_or_code,
                data_inicio=data_inicio or None,
                data_fim=data_fim or None,
                correcao_ipca=correcao_ipca,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_resultado_fiscal(
    tema: str = "10",
    data_inicio: str = "",
    data_fim: str = "",
    codigo_serie: str = "",
    correcao_ipca: bool = False,
) -> str:
    """Query RTN resultado-fiscal table (theme 10/13/20), optionally one series.

    Returns monthly points in R$ milhoes from Boletim Resultado do Tesouro Nacional.
    """
    try:
        return _ok(
            providers.get_resultado_fiscal(
                tema=tema,
                data_inicio=data_inicio or None,
                data_fim=data_fim or None,
                codigo_serie=codigo_serie or None,
                correcao_ipca=correcao_ipca,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_grandes_numeros(metric: str = "") -> str:
    """Headline figures from Tesouro Transparente (resultado_primario, estoque_dpf...)."""
    try:
        return _ok(providers.get_grandes_numeros(metric or None))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def ckan_package_search(query: str = "resultado do tesouro", rows: int = 10) -> str:
    """Search open datasets on Tesouro Transparente (CKAN), including RTN spreadsheets."""
    try:
        return _ok(providers.ckan_package_search(query, rows=rows))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def ckan_package_show(package_id: str = "resultado-do-tesouro-nacional") -> str:
    """Show one CKAN package with download URLs (XLSX historical RTN, metadata, API docs)."""
    try:
        return _ok(providers.ckan_package_show(package_id))
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
