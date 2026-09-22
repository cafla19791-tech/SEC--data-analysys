"""Testes do relatório do Acórdão TCU 1.839/2018-Plenário."""

from __future__ import annotations

from scripts.gerar_relatorio_acordao_1839_2018 import (
    IGP,
    TABELA2_TOTAIS,
    escrever_markdown,
    tabela2_rows,
)
from scripts.petrobras_divida_bruta_20f import montar_dataframe as df_divida
from scripts.petrobras_juros_pagos_20f import montar_dataframe as df_juros
from scripts.petrobras_lucro_liquido_20f import montar_dataframe as df_lucro


def test_tabela2_totais_do_acordao():
    assert TABELA2_TOTAIS == {
        "Reuniões": 188,
        "Arquivos": 957,
        "Páginas": 13175,
        "Ocorrências": 719,
    }
    rows = tabela2_rows()
    assert rows[0][1] == "2004"
    assert rows[1][-1] == "**188**"


def test_igp_total_24_sendo_23_rnest():
    assert IGP[1][-1] == "23"
    assert IGP[-1][-1] == "**24**"


def test_relatorio_apresenta_acordao_tabelas_e_20f():
    rel = escrever_markdown(df_divida(), df_juros(), df_lucro(), "teste")
    assert "Acórdão" in rel
    assert "Dívida Bruta do Governo Geral" in rel
    assert "188" in rel
    assert "83,69" in rel or "83.69" in rel
    assert "132.158" in rel
    assert "7.308" in rel
    assert "−8.450" in rel or "-8.450" in rel
    assert "R$ 43" in rel
    assert "Tabela 2" in rel
    assert "IG-P" in rel
