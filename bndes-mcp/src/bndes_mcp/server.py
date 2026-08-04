"""MCP server minimo para consulta de operacoes BNDES."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import excel_export, providers

mcp = FastMCP("bndes-operacoes")


@mcp.tool()
def bndes_operacoes_por_documento(documento: str, rows: int = 10000) -> str:
    """Busca operacoes BNDES por CNPJ (14) ou CPF (11) e devolve resumo + amostra."""
    data = providers.fetch_operacoes(documento, rows=rows)
    docs = list(data.get("response", {}).get("docs") or [])
    summary = providers.summarize(docs)
    amostra = [providers.flatten_operacao(d) for d in docs[:20]]
    return json.dumps(
        {"resumo": summary, "numFound": data.get("response", {}).get("numFound"), "amostra": amostra},
        ensure_ascii=False,
        indent=2,
        default=str,
    )


@mcp.tool()
def bndes_excel_por_documento(documento: str, out_path: str = "", rows: int = 10000) -> str:
    """Busca operacoes BNDES e grava Excel (Resumo/Operacoes/Subcreditos)."""
    data = providers.fetch_operacoes(documento, rows=rows)
    docs = list(data.get("response", {}).get("docs") or [])
    dig = providers.digits_only(documento)
    path = Path(out_path) if out_path else Path.cwd() / f"bndes_{dig}.xlsx"
    excel_export.write_excel(docs, path)
    summary = providers.summarize(docs)
    summary["excel"] = str(path.resolve())
    return json.dumps(summary, ensure_ascii=False, indent=2, default=str)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
