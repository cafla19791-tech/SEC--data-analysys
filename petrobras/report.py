"""Formatação de tabelas e texto de análise."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def _money_brl(value: float) -> str:
    sign = "-" if value < 0 else ""
    abs_v = abs(value)
    return f"{sign}R$ {abs_v / 1_000_000_000:,.2f} bi".replace(",", "X").replace(".", ",").replace("X", ".")


def _money_usd(value: float) -> str:
    sign = "-" if value < 0 else ""
    abs_v = abs(value)
    return f"{sign}US$ {abs_v / 1_000_000_000:,.2f} bi"


def _pct(value: float | None) -> str:
    if value is None:
        return "—"
    return f"{value * 100:+.1f}%"


def format_markdown_table(result: dict[str, Any], heading_level: int = 1) -> str:
    hashes = "#" * max(1, min(heading_level, 4))
    lines = [
        f"{hashes} Lucro líquido — {result['name']} ({result.get('ticker', 'PBR')})",
        "",
        f"**Métrica:** {result['metric']}",
        f"**CIK:** {result['cik']:010d}",
        "",
        "| Ano | Lucro líquido (R$) | Lucro líquido (US$) | YoY (R$) | YoY (US$) | FX médio |",
        "|----:|-------------------:|--------------------:|---------:|----------:|---------:|",
    ]
    for row in result["years"]:
        lines.append(
            "| {year} | {brl} | {usd} | {yoy_brl} | {yoy_usd} | {fx:.4f} |".format(
                year=row["year"],
                brl=_money_brl(row["net_income_brl"]),
                usd=_money_usd(row["net_income_usd"]),
                yoy_brl=_pct(row["yoy_brl_pct"]),
                yoy_usd=_pct(row["yoy_usd_pct"]),
                fx=row["usd_brl_avg"],
            )
        )

    notes = result.get("currency_notes", {})
    lines.extend(
        [
            "",
            "## Notas metodológicas",
            "",
            f"- **USD:** {notes.get('usd', '')}",
            f"- **BRL:** {notes.get('brl', '')}",
            "- **YoY:** variação percentual ano a ano; base negativa usa denominador em módulo.",
            "",
        ]
    )
    return "\n".join(lines)


def write_analysis_stub(result: dict[str, Any], path: Path) -> None:
    """Gera/atualiza o arquivo de análise com a tabela embutida (texto analítico fixo no repo)."""
    # Se já existe análise completa, apenas atualiza o bloco de tabela.
    # heading_level=2 evita H1 duplicado dentro de reports/*_analysis.md
    table = format_markdown_table(result, heading_level=2)
    marker_start = "<!-- AUTO-TABLE:START -->"
    marker_end = "<!-- AUTO-TABLE:END -->"
    block = f"{marker_start}\n{table}\n{marker_end}"

    if path.exists():
        text = path.read_text(encoding="utf-8")
        if marker_start in text and marker_end in text:
            before = text.split(marker_start)[0]
            after = text.split(marker_end, 1)[1]
            path.write_text(before + block + after, encoding="utf-8")
            return

    # Bootstrap mínimo — a análise narrativa completa é versionada em reports/
    path.write_text(
        "# Análise — Lucro líquido Petrobras\n\n" + block + "\n",
        encoding="utf-8",
    )
