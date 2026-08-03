"""Testes unitários da extração de lucro líquido Petrobras."""

from __future__ import annotations

import pytest

from petrobras.net_income import build_table, extract_annual_usd, _yoy


SAMPLE_FACTS = {
    "entityName": "PETROBRAS - PETROLEO BRASILEIRO SA",
    "facts": {
        "ifrs-full": {
            "ProfitLossAttributableToOwnersOfParent": {
                "units": {
                    "USD": [
                        {
                            "end": "2022-12-31",
                            "val": 36623000000,
                            "form": "20-F",
                            "fp": "FY",
                            "frame": "CY2022",
                            "filed": "2025-04-03",
                        },
                        {
                            "end": "2023-12-31",
                            "val": 24884000000,
                            "form": "20-F",
                            "fp": "FY",
                            "frame": "CY2023",
                            "filed": "2025-04-03",
                        },
                        {
                            "end": "2024-06-30",
                            "val": 4734000000,
                            "form": "6-K/A",
                            "fp": "Q2",
                            "frame": "CY2024Q2",
                            "filed": "2025-08-26",
                        },
                        {
                            "end": "2024-12-31",
                            "val": 7528000000,
                            "form": "20-F",
                            "fp": "FY",
                            "frame": "CY2024",
                            "filed": "2025-04-03",
                        },
                    ]
                }
            }
        }
    },
}


def test_extract_annual_ignores_interim_and_keeps_fy():
    series = extract_annual_usd(SAMPLE_FACTS, years=10)
    years = [r["year"] for r in series]
    assert years == [2022, 2023, 2024]
    assert series[-1]["net_income_usd"] == 7528000000.0
    assert series[-1]["form"] == "20-F"


def test_yoy_from_loss_to_profit_uses_abs_denominator():
    assert _yoy(100, -50) == 3.0


def test_build_table_computes_brl_and_yoy():
    usd_series = extract_annual_usd(SAMPLE_FACTS, years=10)
    fx = {
        "2022": {"avg_usd_brl": 5.0, "source": "test"},
        "2023": {"avg_usd_brl": 5.0, "source": "test"},
        "2024": {"avg_usd_brl": 5.0, "source": "test"},
    }
    rows = build_table(usd_series, fx)
    assert rows[0]["yoy_usd_pct"] is None
    assert rows[1]["net_income_brl"] == 24884000000 * 5.0
    expected = (24884000000 - 36623000000) / 36623000000
    assert rows[1]["yoy_usd_pct"] == pytest.approx(expected)
