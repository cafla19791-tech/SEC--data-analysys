"""Testes do discriminativo de geração operacional de caixa da Petrobras (20-F / 6-K)."""

from __future__ import annotations

import pandas as pd

from scripts.petrobras_caixa_operacional_20f import (
    LINHAS,
    edgar_url,
    escrever_markdown,
    montar_dataframe,
)


def test_serie_cobre_2002_a_2026_sem_furo() -> None:
    anos = [linha["ano"] for linha in LINHAS]
    assert anos == list(range(2002, 2027))
    assert len(LINHAS) == 25


def test_valores_e_paginas_ancoram_anos_chave() -> None:
    df = montar_dataframe()
    por_ano = {int(row["ano"]): row for row in df.to_dict(orient="records")}

    assert por_ano[2002]["caixa_operacional_usd_milhoes"] == 6287
    assert por_ano[2002]["pagina"] == "F-7"
    assert por_ano[2002]["norma"] == "US GAAP"

    assert por_ano[2004]["caixa_operacional_usd_milhoes"] == 8833
    assert por_ano[2004]["pagina"] == "F-9"

    assert por_ano[2010]["caixa_operacional_usd_milhoes"] == 28495
    assert por_ano[2010]["pagina"] == "F-9"
    assert "000129281411001552" in por_ano[2010]["url_documento"]

    assert por_ano[2011]["caixa_operacional_usd_milhoes"] == 33698
    assert por_ano[2011]["norma"] == "IFRS"
    assert por_ano[2011]["pagina"] == "F-10"

    assert por_ano[2014]["caixa_operacional_usd_milhoes"] == 26632
    assert por_ano[2014]["pagina"] == "F-7"

    assert por_ano[2022]["caixa_operacional_usd_milhoes"] == 49717
    assert por_ano[2022]["pagina"] == "F-6"

    assert por_ano[2025]["caixa_operacional_usd_milhoes"] == 36047
    assert por_ano[2025]["pagina"] == "F-6"
    assert por_ano[2025]["tipo"] == "20-F"


def test_2026_e_parcial_6k_e_nao_tem_variacao_yoy() -> None:
    df = montar_dataframe()
    row = df.loc[df["ano"] == 2026].iloc[0]
    assert row["tipo"] == "6-K"
    assert row["periodo"] == "1S (jan–jun)"
    assert row["caixa_operacional_usd_milhoes"] == 20649
    assert row["pagina"] == "6"
    assert pd.isna(row["variacao_pct"])
    assert pd.isna(row["variacao_usd_milhoes"])
    assert "pbrfs2q26usd_6k.htm" in row["url_documento"]
    assert "000129281426004133" in row["url_documento"]


def test_pico_2022_e_minimo_2002_e_variacao_anual() -> None:
    df = montar_dataframe()
    anuais = df[df["periodo"] == "ano"]
    assert int(anuais.loc[anuais["caixa_operacional_usd_milhoes"].idxmax(), "ano"]) == 2022
    assert int(anuais.loc[anuais["caixa_operacional_usd_milhoes"].idxmin(), "ano"]) == 2002
    assert pd.isna(df.loc[df["ano"] == 2002, "variacao_pct"].iloc[0])
    assert df.loc[df["ano"] == 2003, "variacao_pct"].iloc[0] == 0.3630
    assert df.loc[df["ano"] == 2005, "variacao_pct"].iloc[0] == 0.7112
    assert df.loc[df["ano"] == 2022, "variacao_pct"].iloc[0] == 0.3156
    assert df.loc[df["ano"] == 2025, "variacao_pct"].iloc[0] == -0.0510


def test_totais_e_reapresentacoes_nao_substituem_proprio_ano() -> None:
    df = montar_dataframe()
    anuais = df[df["periodo"] == "ano"]
    total_anos = int(anuais["caixa_operacional_usd_milhoes"].sum())
    total_com_1s = int(df["caixa_operacional_usd_milhoes"].sum())
    assert total_anos == 643420
    assert total_com_1s == 664069
    por_ano = {int(row["ano"]): row for row in df.to_dict(orient="records")}
    assert por_ano[2004]["caixa_operacional_usd_milhoes"] == 8833
    assert por_ano[2010]["caixa_operacional_usd_milhoes"] == 28495
    assert "8,155" in por_ano[2004]["trecho"] or "8.155" in por_ano[2004]["trecho"]
    assert "30,110" in por_ano[2010]["trecho"] or "30.110" in por_ano[2010]["trecho"]


def test_urls_edgar_e_markdown_tem_pagina() -> None:
    row = LINHAS[-1]
    url = edgar_url(row["accession"], row["arquivo"])
    assert url.endswith("pbrfs2q26usd_6k.htm")
    assert "1119639" in url
    md = escrever_markdown(montar_dataframe(), "teste")
    assert "49.717" in md or "49,717" in md
    assert "F-6" in md
    assert "6-K" in md
    assert "643.420" in md
    assert "664.069" in md
    assert "Total 2002–2025" in md
    assert "8.833" in md or "8,833" in md
    assert "28.495" in md or "28,495" in md
