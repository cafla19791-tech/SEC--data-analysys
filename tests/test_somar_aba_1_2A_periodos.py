"""Testes da soma por período da aba 1.2-A."""

from __future__ import annotations

from pathlib import Path

import pytest

from scripts.somar_aba_1_2A_periodos import PERIODS, sum_periods, write_outputs

DATA = Path(__file__).resolve().parents[1] / "data" / "rtn" / "serie_historica_mai26.xlsx"


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_period_month_counts_and_totals():
    table = sum_periods(DATA)
    assert table["sheet"] == "1.2-A"
    assert "Mai/2026" in (table["unit"] or "")
    months = [r["meses"] for r in table["resumo"]]
    assert months == [72, 161, 31, 48, 41]
    # Totals in R$ bi (approx)
    bi = [r["soma_linhas_6_177_R$bi"] for r in table["resumo"]]
    assert bi[0] == pytest.approx(41824.35, abs=0.1)
    assert bi[4] == pytest.approx(72315.77, abs=0.1)
    assert len(table["detail"]) == 172  # rows 6..177


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_write_xlsx(tmp_path):
    table = sum_periods(DATA)
    out = tmp_path / "out.xlsx"
    write_outputs(table, out)
    assert out.exists() and out.stat().st_size > 1000
    assert len(PERIODS) == 5
