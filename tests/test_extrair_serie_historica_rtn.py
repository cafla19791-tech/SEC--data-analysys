"""Testes do extrator RTN serie_historica."""

from __future__ import annotations

from pathlib import Path

import pytest

from scripts.extrair_serie_historica_rtn import (
    ROW_NEEDLES_1_1,
    annual_sum_from_monthly_sheet,
    extract_annual_rtn,
    write_csv,
)

DATA = Path(__file__).resolve().parents[1] / "data" / "rtn" / "serie_historica_mai26.xlsx"


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_extract_ipca_mai26():
    table = extract_annual_rtn(DATA, year_from=2020, year_to=2025, constantes_ipca=True)
    assert table["sheet"] == "1.1-A"
    assert "Mai/2026" in (table["unit"] or "")
    assert table["missing"] == []
    by_year = {r["ano"]: r for r in table["rows"]}
    assert by_year[2020]["resultado_primario_R$bi"] == pytest.approx(-1060.26, abs=0.05)
    assert by_year[2022]["resultado_primario_R$bi"] == pytest.approx(58.20, abs=0.05)
    assert by_year[2025]["resultado_primario_R$bi"] == pytest.approx(-62.52, abs=0.05)


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_extract_corrente_and_csv(tmp_path):
    table = extract_annual_rtn(DATA, year_from=2024, year_to=2025, constantes_ipca=False)
    assert table["sheet"] == "1.1"
    assert "Correntes" in (table["unit"] or "")
    out = tmp_path / "out.csv"
    write_csv(table["rows"], out)
    text = out.read_text(encoding="utf-8")
    assert "resultado_primario_R$bi" in text
    assert "2025" in text


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_month_headers_are_datetimes():
    core = annual_sum_from_monthly_sheet(
        DATA, "1.1-A", ROW_NEEDLES_1_1, year_from=2025, year_to=2025, min_months=1
    )
    assert 2025 in core["series"]["resultado_primario"]
    # sanity: needles matched
    assert set(core["labels"]) == set(ROW_NEEDLES_1_1)
