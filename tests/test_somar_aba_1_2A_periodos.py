"""Testes: totais por item (A6–A177) da aba 1.2-A."""

from __future__ import annotations

from pathlib import Path

import pytest

from scripts.somar_aba_1_2A_periodos import PERIODS, totails_por_item, write_outputs

DATA = Path(__file__).resolve().parents[1] / "data" / "rtn" / "serie_historica_mai26.xlsx"


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_172_itens_e_periodos():
    table = totails_por_item(DATA)
    assert table["n_itens"] == 172
    assert table["itens"][0]["linha"] == 6
    assert table["itens"][-1]["linha"] == 177
    assert [p["meses"] for p in table["periodos"]] == [72, 161, 31, 48, 41]
    assert table["periodos"][0]["colunas_excel"] == "B:BU"
    assert table["periodos"][1]["colunas_excel"] == "BV:HZ"


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_a6_periodo1_eh_soma_b_bu():
    """A6 jan/97–dez/02 = SOMA(B6:BU6), não soma entre linhas."""
    table = totails_por_item(DATA)
    a6 = table["itens"][0]
    assert "RECEITA TOTAL" in (a6["item"] or "")
    # R$ mi IPCA Mai/2026
    assert a6["jan/97–dez/02"] == pytest.approx(6433212.26, abs=0.05)
    a7 = table["itens"][1]
    assert a7["jan/97–dez/02"] == pytest.approx(4204612.16, abs=0.05)
    # itens distintos — NÃO são iguais a uma soma agregada única
    assert a6["jan/97–dez/02"] != a7["jan/97–dez/02"]


@pytest.mark.skipif(not DATA.exists(), reason="serie_historica_mai26.xlsx ausente")
def test_write_xlsx(tmp_path):
    table = totails_por_item(DATA)
    out = tmp_path / "out.xlsx"
    write_outputs(table, out)
    assert out.exists() and out.stat().st_size > 1000
    assert len(PERIODS) == 5
