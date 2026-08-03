"""Cálculo de indicadores e crescimento a partir de séries financeiras."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class YearRow:
    year: str
    revenue: float | None = None
    net_income: float | None = None
    total_assets: float | None = None
    total_liabilities: float | None = None
    equity: float | None = None
    operating_income: float | None = None
    current_assets: float | None = None
    current_liabilities: float | None = None
    cash: float | None = None

    # derivados
    profit_margin: float | None = None
    operating_margin: float | None = None
    roe: float | None = None
    roa: float | None = None
    debt_to_equity: float | None = None
    current_ratio: float | None = None
    revenue_growth: float | None = None
    net_income_growth: float | None = None


@dataclass
class FinancialMetrics:
    ticker: str
    cik: int
    name: str
    rows: list[YearRow] = field(default_factory=list)
    summary: dict[str, Any] = field(default_factory=dict)


def _latest_value(series: list[dict[str, Any]], end: str) -> float | None:
    for item in series:
        if item.get("end") == end:
            return item.get("value")
    return None


def _pct(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator is None or denominator == 0:
        return None
    return numerator / denominator


def _growth(current: float | None, previous: float | None) -> float | None:
    if current is None or previous is None or previous == 0:
        return None
    return (current - previous) / abs(previous)


def _year_label(end: str) -> str:
    return end[:4] if end else "?"


def compute_metrics(snapshot: dict[str, Any]) -> FinancialMetrics:
    series = snapshot.get("series", {})
    ends: set[str] = set()
    for values in series.values():
        for item in values:
            if item.get("end"):
                ends.add(item["end"])

    sorted_ends = sorted(ends)
    rows: list[YearRow] = []

    for end in sorted_ends:
        row = YearRow(
            year=_year_label(end),
            revenue=_latest_value(series.get("revenue", []), end),
            net_income=_latest_value(series.get("net_income", []), end),
            total_assets=_latest_value(series.get("total_assets", []), end),
            total_liabilities=_latest_value(series.get("total_liabilities", []), end),
            equity=_latest_value(series.get("equity", []), end),
            operating_income=_latest_value(series.get("operating_income", []), end),
            current_assets=_latest_value(series.get("current_assets", []), end),
            current_liabilities=_latest_value(
                series.get("current_liabilities", []), end
            ),
            cash=_latest_value(series.get("cash", []), end),
        )
        row.profit_margin = _pct(row.net_income, row.revenue)
        row.operating_margin = _pct(row.operating_income, row.revenue)
        row.roe = _pct(row.net_income, row.equity)
        row.roa = _pct(row.net_income, row.total_assets)
        row.debt_to_equity = _pct(row.total_liabilities, row.equity)
        row.current_ratio = _pct(row.current_assets, row.current_liabilities)
        rows.append(row)

    for i, row in enumerate(rows):
        prev = rows[i - 1] if i > 0 else None
        if prev:
            row.revenue_growth = _growth(row.revenue, prev.revenue)
            row.net_income_growth = _growth(row.net_income, prev.net_income)

    summary: dict[str, Any] = {}
    if rows:
        latest = rows[-1]
        summary = {
            "latest_year": latest.year,
            "revenue": latest.revenue,
            "net_income": latest.net_income,
            "profit_margin": latest.profit_margin,
            "roe": latest.roe,
            "roa": latest.roa,
            "debt_to_equity": latest.debt_to_equity,
            "current_ratio": latest.current_ratio,
            "revenue_cagr": _cagr(
                [r.revenue for r in rows if r.revenue is not None]
            ),
            "years_covered": len(rows),
        }

    return FinancialMetrics(
        ticker=snapshot.get("ticker", ""),
        cik=int(snapshot.get("cik", 0)),
        name=snapshot.get("name", ""),
        rows=rows,
        summary=summary,
    )


def _cagr(values: list[float]) -> float | None:
    if len(values) < 2 or values[0] == 0:
        return None
    periods = len(values) - 1
    if values[0] < 0 or values[-1] < 0:
        return None
    return (values[-1] / values[0]) ** (1 / periods) - 1


def snapshot_from_dataframe(df, ticker: str = "", name: str = "") -> dict[str, Any]:
    """Converte um DataFrame/CSV no formato esperado por compute_metrics."""
    import pandas as pd

    if not isinstance(df, pd.DataFrame):
        raise TypeError("df deve ser um pandas.DataFrame")

    required = {"year"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes: {sorted(missing)}")

    metric_cols = [
        "revenue",
        "net_income",
        "total_assets",
        "total_liabilities",
        "equity",
        "operating_income",
        "current_assets",
        "current_liabilities",
        "cash",
    ]

    series: dict[str, list[dict[str, Any]]] = {m: [] for m in metric_cols}
    for _, row in df.sort_values("year").iterrows():
        year = str(int(row["year"])) if pd.notna(row["year"]) else None
        if not year:
            continue
        end = f"{year}-12-31"
        for metric in metric_cols:
            if metric in df.columns and pd.notna(row[metric]):
                series[metric].append(
                    {
                        "end": end,
                        "value": float(row[metric]),
                        "filed": end,
                        "concept": "csv",
                        "form": "CSV",
                    }
                )

    return {
        "ticker": ticker or str(df.attrs.get("ticker", "CSV")),
        "cik": int(df.attrs.get("cik", 0)),
        "name": name or str(df.attrs.get("name", "Empresa (CSV)")),
        "series": series,
    }
