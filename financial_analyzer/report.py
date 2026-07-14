"""Formatação de relatórios em texto e exportação CSV/JSON."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .metrics import FinancialMetrics, YearRow


def _money(value: float | None) -> str:
    if value is None:
        return "—"
    abs_v = abs(value)
    sign = "-" if value < 0 else ""
    if abs_v >= 1_000_000_000:
        return f"{sign}${abs_v / 1_000_000_000:,.2f}B"
    if abs_v >= 1_000_000:
        return f"{sign}${abs_v / 1_000_000:,.2f}M"
    if abs_v >= 1_000:
        return f"{sign}${abs_v / 1_000:,.2f}K"
    return f"{sign}${abs_v:,.2f}"


def _pct(value: float | None) -> str:
    if value is None:
        return "—"
    return f"{value * 100:.1f}%"


def format_report(metrics: FinancialMetrics) -> str:
    lines: list[str] = []
    header = f"{metrics.name}"
    if metrics.ticker:
        header += f" ({metrics.ticker})"
    if metrics.cik:
        header += f" | CIK {metrics.cik:010d}"

    lines.append("=" * 72)
    lines.append("ANÁLISE FINANCEIRA")
    lines.append(header)
    lines.append("=" * 72)

    if not metrics.rows:
        lines.append("Nenhum dado anual encontrado.")
        return "\n".join(lines)

    s = metrics.summary
    lines.append("")
    lines.append(f"Ano mais recente : {s.get('latest_year', '—')}")
    lines.append(f"Receita          : {_money(s.get('revenue'))}")
    lines.append(f"Lucro líquido    : {_money(s.get('net_income'))}")
    lines.append(f"Margem líquida   : {_pct(s.get('profit_margin'))}")
    lines.append(f"ROE              : {_pct(s.get('roe'))}")
    lines.append(f"ROA              : {_pct(s.get('roa'))}")
    de = s.get("debt_to_equity")
    de_text = f"{de:.2f}x" if de is not None else "—"
    lines.append(f"Dívida / PL      : {de_text}")
    current_ratio = s.get("current_ratio")
    cr_text = f"{current_ratio:.2f}" if current_ratio is not None else "—"
    lines.append(f"Liquidez corrente: {cr_text}")
    lines.append(f"CAGR receita     : {_pct(s.get('revenue_cagr'))}")
    lines.append("")

    lines.append("-" * 72)
    lines.append(
        f"{'Ano':<6}{'Receita':>14}{'Lucro Líq.':>14}"
        f"{'Margem':>10}{'ROE':>10}{'Δ Receita':>12}"
    )
    lines.append("-" * 72)
    for row in metrics.rows:
        lines.append(
            f"{row.year:<6}"
            f"{_money(row.revenue):>14}"
            f"{_money(row.net_income):>14}"
            f"{_pct(row.profit_margin):>10}"
            f"{_pct(row.roe):>10}"
            f"{_pct(row.revenue_growth):>12}"
        )
    lines.append("-" * 72)
    lines.append("")
    lines.append(_interpretation(metrics))
    return "\n".join(lines)


def _interpretation(metrics: FinancialMetrics) -> str:
    s = metrics.summary
    notes: list[str] = ["Leitura rápida:"]

    margin = s.get("profit_margin")
    if margin is not None:
        if margin >= 0.20:
            notes.append("• Margem líquida elevada — operação bastante rentável.")
        elif margin >= 0.08:
            notes.append("• Margem líquida saudável.")
        elif margin >= 0:
            notes.append("• Margem líquida positiva, porém apertada.")
        else:
            notes.append("• Empresa no prejuízo no período mais recente.")

    growth = s.get("revenue_cagr")
    if growth is not None:
        if growth >= 0.15:
            notes.append("• Crescimento de receita forte no período analisado.")
        elif growth >= 0.05:
            notes.append("• Crescimento de receita moderado.")
        elif growth >= 0:
            notes.append("• Receita praticamente estável.")
        else:
            notes.append("• Receita em contração no período.")

    de = s.get("debt_to_equity")
    if de is not None:
        if de > 2:
            notes.append("• Alavancagem elevada (passivo / patrimônio).")
        elif de > 1:
            notes.append("• Alavancagem moderada.")
        else:
            notes.append("• Estrutura de capital relativamente conservadora.")

    cr = s.get("current_ratio")
    if cr is not None:
        if cr < 1:
            notes.append("• Liquidez corrente abaixo de 1 — atenção ao curto prazo.")
        elif cr < 1.5:
            notes.append("• Liquidez corrente adequada.")
        else:
            notes.append("• Boa folga de liquidez de curto prazo.")

    return "\n".join(notes)


def rows_to_records(metrics: FinancialMetrics) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for row in metrics.rows:
        records.append(_row_dict(row))
    return records


def _row_dict(row: YearRow) -> dict[str, Any]:
    return {
        "year": row.year,
        "revenue": row.revenue,
        "net_income": row.net_income,
        "total_assets": row.total_assets,
        "total_liabilities": row.total_liabilities,
        "equity": row.equity,
        "operating_income": row.operating_income,
        "current_assets": row.current_assets,
        "current_liabilities": row.current_liabilities,
        "cash": row.cash,
        "profit_margin": row.profit_margin,
        "operating_margin": row.operating_margin,
        "roe": row.roe,
        "roa": row.roa,
        "debt_to_equity": row.debt_to_equity,
        "current_ratio": row.current_ratio,
        "revenue_growth": row.revenue_growth,
        "net_income_growth": row.net_income_growth,
    }


def export_json(metrics: FinancialMetrics, path: str | Path) -> None:
    payload = {
        "ticker": metrics.ticker,
        "cik": metrics.cik,
        "name": metrics.name,
        "summary": metrics.summary,
        "years": rows_to_records(metrics),
    }
    Path(path).write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def export_csv(metrics: FinancialMetrics, path: str | Path) -> None:
    import pandas as pd

    df = pd.DataFrame(rows_to_records(metrics))
    df.to_csv(path, index=False)
