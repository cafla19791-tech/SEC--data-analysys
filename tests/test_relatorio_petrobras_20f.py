"""Testes do relatório formal Petrobras (dívida, caixa, juros e lucro)."""

from __future__ import annotations

from scripts.gerar_relatorio_petrobras_20f import (
    escrever_markdown,
    extrair_tabela_evolucao,
)
from scripts.petrobras_caixa_operacional_20f import (
    escrever_markdown as md_caixa,
    montar_dataframe as df_caixa,
)
from scripts.petrobras_divida_bruta_20f import (
    escrever_markdown as md_divida,
    montar_dataframe as df_divida,
)
from scripts.petrobras_juros_pagos_20f import (
    escrever_markdown as md_juros,
    montar_dataframe as df_juros,
)
from scripts.petrobras_lucro_liquido_20f import (
    escrever_markdown as md_lucro,
    montar_dataframe as df_lucro,
)


def test_relatorio_reproduz_tabelas_dos_discriminativos():
    div, cx, jur, luc = df_divida(), df_caixa(), df_juros(), df_lucro()
    rel = escrever_markdown(div, cx, jur, luc, "teste")
    assert extrair_tabela_evolucao(md_divida(div, "x")) in rel
    assert extrair_tabela_evolucao(md_caixa(cx, "x")) in rel
    assert extrair_tabela_evolucao(md_juros(jur, "x")) in rel
    assert extrair_tabela_evolucao(md_lucro(luc, "x")) in rel


def test_caixa_operacional_vem_logo_apos_a_divida_bruta():
    rel = escrever_markdown(df_divida(), df_caixa(), df_juros(), df_lucro(), "teste")
    i_div = rel.index("## 3. Dívida bruta")
    i_cx = rel.index("## 4. Geração operacional de caixa")
    i_jur = rel.index("## 5. Juros pagos")
    i_luc = rel.index("## 6. Lucro líquido")
    assert i_div < i_cx < i_jur < i_luc


def test_relatorio_apresenta_series_e_totais():
    rel = escrever_markdown(df_divida(), df_caixa(), df_juros(), df_lucro(), "teste")
    assert "evolução da dívida bruta" in rel.lower()
    assert "geração operacional de caixa" in rel.lower()
    assert "juros pagos" in rel.lower()
    assert "lucro líquido" in rel.lower()
    assert "não são a Dívida Bruta do Governo Geral" in rel
    assert "77.996" in rel
    assert "79.066" in rel
    assert "253.447" in rel
    assert "270.074" in rel
    assert "643.420" in rel
    assert "664.069" in rel
    assert "55.113" in rel
    assert "Métrica" in rel
    assert "| 2016 | ano | 2017-04-27 | 7.308 |" in rel
    assert "| 2022 | ano | 2023-03-29 | 36.623 |" in rel
    assert "| 2014 | 2015-05-15 | 132.158 |" in rel
    assert "| 2022 | ano | 2023-03-29 | 49.717 |" in rel
