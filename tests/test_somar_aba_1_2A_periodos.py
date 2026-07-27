"""Testes: totais por item (não soma vertical das 172 linhas)."""

from __future__ import annotations

from pathlib import Path

import pytest

from scripts.somar_aba_1_2A_periodos import PERIODS, totals_por_item, write_outputs

DATA = Path(__file__).resolve().parents[1] / "data" / "rtn" / "serie_historica_mai26.xlsx"


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_172_itens_e_colunas_periodo():
    table = totals_por_item(DATA)
    assert table["n_itens"] == 172
    assert len(table["periodos"]) == 5
    months = [p["meses"] for p in table["periodos"]]
    assert months == [72, 161, 31, 48, 41]
    # jan/97–dez/02 = B..BU (=SUM(B6:BU6)), não até BV
    assert table["periodos"][0]["coluna_inicio"] == "B"
    assert table["periodos"][0]["coluna_fim"] == "BU"
    assert table["periodos"][0]["formula_exemplo_linha6"] == "=SUM(B6:BU6)"
    assert table["periodos"][1]["coluna_inicio"] == "BV"


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_receita_total_linha6_periodo1():
    table = totals_por_item(DATA)
    item6 = next(i for i in table["itens"] if i["linha"] == 6)
    assert "RECEITA TOTAL" in str(item6["item"])
    # SUM(B6:BU6) ≈ 6.433.212,26 R$ mi
    assert item6["jan/97–dez/02"] == pytest.approx(6433212.26, abs=0.5)
    assert item6["jan/97–dez/02_R$bi"] == pytest.approx(6433.21, abs=0.01)


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_itens_diferentes_nao_sao_iguais():
    """Garante que estamos somando por linha, não um único total vertical."""
    table = totals_por_item(DATA)
    a6 = next(i for i in table["itens"] if i["linha"] == 6)
    a7 = next(i for i in table["itens"] if i["linha"] == 7)
    assert a6["jan/97–dez/02"] != a7["jan/97–dez/02"]


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_write_xlsx(tmp_path):
    table = totals_por_item(DATA)
    out = tmp_path / "out.xlsx"
    write_outputs(table, out)
    assert out.exists() and out.stat().st_size > 1000
    assert len(PERIODS) == 5
