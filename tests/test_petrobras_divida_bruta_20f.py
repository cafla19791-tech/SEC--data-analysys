"""Testes do discriminativo de dívida bruta Petrobras (Forms 20-F)."""

from __future__ import annotations

from scripts.petrobras_divida_bruta_20f import (
    LINHAS,
    edgar_url,
    escrever_markdown,
    montar_dataframe,
)


def test_serie_cobre_2002_a_2025_sem_furo():
    anos = [row["ano"] for row in LINHAS]
    assert anos == list(range(2002, 2026))


def test_valores_ancora_conferidos_nos_20f():
    por_ano = {row["ano"]: row["divida_bruta_usd_milhoes"] for row in LINHAS}
    assert por_ano[2002] == 14_680
    assert por_ano[2014] == 132_158
    assert por_ano[2018] == 84_360
    assert por_ano[2019] == 87_121
    assert por_ano[2022] == 53_799
    assert por_ano[2025] == 69_793


def test_pico_em_2014_e_minimo_recente_em_2022():
    df = montar_dataframe()
    assert int(df.loc[df["divida_bruta_usd_milhoes"].idxmax(), "ano"]) == 2014
    recente = df[df["ano"] >= 2019]
    assert int(recente.loc[recente["divida_bruta_usd_milhoes"].idxmin(), "ano"]) == 2022


def test_variacao_primeiro_ano_vazia_e_yoy_2015():
    df = montar_dataframe()
    assert pd_isna(df.loc[0, "variacao_usd_milhoes"])
    row_2015 = df.loc[df["ano"] == 2015].iloc[0]
    assert row_2015["variacao_usd_milhoes"] == 126_216 - 132_158


def test_urls_edgar_e_markdown_tem_pagina():
    row = LINHAS[-1]
    url = edgar_url(row["accession"], row["arquivo"])
    assert url.endswith("pbrform20f_2025.htm")
    assert "1119639" in url
    md = escrever_markdown(montar_dataframe(), "teste")
    assert "página 180" in md.lower() or "180 / F-22" in md
    assert "87.121" in md or "87,121" in md or "87.121" in md.replace(",", ".")


def pd_isna(value) -> bool:
    import pandas as pd

    return bool(pd.isna(value))
