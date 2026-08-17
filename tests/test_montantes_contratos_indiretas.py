"""Testes da consolidação anual de contratos indiretos do BNDES."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.montantes_contratos_indiretas import (
    agregar_registros,
    ano_da_data,
    escrever_markdown,
    montar_resumo,
    _num,
)


def test_ano_da_data_iso_e_invalido():
    assert ano_da_data("2002-01-02T00:00:00") == 2002
    assert ano_da_data("2026-06-30") == 2026
    assert ano_da_data(None) is None
    assert ano_da_data("abc") is None
    assert ano_da_data("1899-01-01") is None


def test_num_aceita_float_e_ptbr():
    assert _num(10.5) == 10.5
    assert _num(None) == 0.0
    assert _num("1.234,56") == 1234.56
    assert _num("2975.0") == 2975.0
    assert _num("") == 0.0


def test_agregar_registros_soma_por_ano_e_ignora_fora_do_periodo():
    recs = [
        {
            "data_da_contratacao": "2002-03-01",
            "valor_da_operacao_em_reais": 100,
            "valor_desembolsado_reais": 80,
        },
        {
            "data_da_contratacao": "2002-12-31",
            "valor_da_operacao_em_reais": 50,
            "valor_desembolsado_reais": 50,
        },
        {
            "data_da_contratacao": "2001-12-31",
            "valor_da_operacao_em_reais": 999,
            "valor_desembolsado_reais": 999,
        },
        {
            "data_da_contratacao": "2027-01-01",
            "valor_da_operacao_em_reais": 1,
            "valor_desembolsado_reais": 1,
        },
    ]
    acc = agregar_registros(
        recs,
        col_data="data_da_contratacao",
        col_contratado="valor_da_operacao_em_reais",
        col_desembolsado="valor_desembolsado_reais",
    )
    assert acc[2002]["qtd"] == 2
    assert acc[2002]["contratado"] == 150
    assert acc[2002]["desembolsado"] == 130
    assert 2001 not in acc
    assert 2027 not in acc


def test_montar_resumo_soma_automaticas_e_nao_automaticas():
    auto = {2009: {"qtd": 2, "contratado": 1_000_000, "desembolsado": 800_000}}
    nao = {2009: {"qtd": 1, "contratado": 500_000, "desembolsado": 400_000}}
    aprov = pd.DataFrame(
        {
            "ano": [2009],
            "aprovado_oficial_r_milhoes": [1.8],
            "aprovado_oficial_reais": [1_800_000],
        }
    )
    out = montar_resumo(auto, nao, aprov)
    row = out.loc[out["ano"] == 2009].iloc[0]
    assert row["qtd_total"] == 3
    assert row["contratado_total"] == 1_500_000
    assert row["desembolsado_total"] == 1_200_000
    assert abs(row["contratado_total_r_milhoes"] - 1.5) < 1e-9
    assert row["aprovado_oficial_r_milhoes"] == 1.8
    assert set(range(2002, 2027)).issubset(set(out["ano"]))


def test_escrever_markdown_contem_anos_e_total(tmp_path: Path):
    auto = {2002: {"qtd": 1, "contratado": 2_000_000_000, "desembolsado": 1_000_000_000}}
    nao = {2002: {"qtd": 0, "contratado": 0, "desembolsado": 0}}
    aprov = pd.DataFrame(
        {
            "ano": [2002],
            "aprovado_oficial_r_milhoes": [2100.0],
            "aprovado_oficial_reais": [2_100_000_000],
        }
    )
    resumo = montar_resumo(auto, nao, aprov)
    path = tmp_path / "out.md"
    escrever_markdown(resumo, path, "2026-08-17 00:00 UTC")
    texto = path.read_text(encoding="utf-8")
    assert "2002" in texto
    assert "modalidade indireta" in texto
    assert "Total" in texto
