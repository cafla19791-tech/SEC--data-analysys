"""Testes básicos do analisador financeiro (modo offline)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from financial_analyzer.metrics import compute_metrics, snapshot_from_dataframe
from financial_analyzer.report import format_report, rows_to_records


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "data" / "sample_companies.csv"


def test_sample_aapl_metrics():
    df = pd.read_csv(SAMPLE)
    df = df[df["ticker"] == "AAPL"]
    snapshot = snapshot_from_dataframe(df, ticker="AAPL", name="Apple Inc.")
    metrics = compute_metrics(snapshot)

    assert metrics.ticker == "AAPL"
    assert len(metrics.rows) == 5
    assert metrics.summary["latest_year"] == "2024"
    assert metrics.summary["revenue"] == 391035000000
    assert metrics.summary["profit_margin"] is not None
    assert 0.2 < metrics.summary["profit_margin"] < 0.3

    report = format_report(metrics)
    assert "Apple Inc." in report
    assert "Margem líquida" in report

    records = rows_to_records(metrics)
    assert records[-1]["year"] == "2024"
    assert records[1]["revenue_growth"] is not None


def test_cli_demo_exit_zero():
    from analyze_finance import main

    assert main(["--demo", "--company", "MSFT"]) == 0
