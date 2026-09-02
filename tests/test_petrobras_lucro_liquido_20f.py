"""Testes do discriminativo de lucro líquido da Petrobras (20-F / 6-K)."""

from __future__ import annotations

import pandas as pd

from scripts.petrobras_lucro_liquido_20f import (
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

    assert por_ano[2002]["lucro_liquido_usd_milhoes"] == 2311
    assert por_ano[2002]["pagina"] == "F-5"
    assert por_ano[2002]["norma"] == "US GAAP"

    assert por_ano[2010]["lucro_liquido_usd_milhoes"] == 19184
    assert "000129281411001552" in por_ano[2010]["url_documento"]

    assert por_ano[2011]["lucro_liquido_usd_milhoes"] == 20121
    assert por_ano[2011]["norma"] == "IFRS"
    assert por_ano[2011]["pagina"] == "F-7"

    assert por_ano[2014]["lucro_liquido_usd_milhoes"] == -7367
    assert por_ano[2015]["lucro_liquido_usd_milhoes"] == -8450
    assert por_ano[2017]["lucro_liquido_usd_milhoes"] == -91

    assert por_ano[2022]["lucro_liquido_usd_milhoes"] == 36623
    assert por_ano[2022]["pagina"] == "F-4"

    assert por_ano[2025]["lucro_liquido_usd_milhoes"] == 19634
    assert por_ano[2025]["pagina"] == "F-4"


def test_2026_e_parcial_6k_e_nao_tem_variacao_yoy() -> None:
    df = montar_dataframe()
    row = df.loc[df["ano"] == 2026].iloc[0]
    assert row["tipo"] == "6-K"
    assert row["periodo"] == "1S (jan–jun)"
    assert row["lucro_liquido_usd_milhoes"] == 16627
    assert row["pagina"] == "4"
    assert pd.isna(row["variacao_pct"])
    assert pd.isna(row["variacao_usd_milhoes"])
    assert "pbrfs2q26usd_6k.htm" in row["url_documento"]


def test_pico_2022_e_minimo_2015_e_variacao_com_prejuizo() -> None:
    df = montar_dataframe()
    anuais = df[df["periodo"] == "ano"]
    assert int(anuais.loc[anuais["lucro_liquido_usd_milhoes"].idxmax(), "ano"]) == 2022
    assert int(anuais.loc[anuais["lucro_liquido_usd_milhoes"].idxmin(), "ano"]) == 2015
    assert pd.isna(df.loc[df["ano"] == 2014, "variacao_pct"].iloc[0])
    assert df.loc[df["ano"] == 2015, "variacao_pct"].iloc[0] == -0.1470
    assert df.loc[df["ano"] == 2016, "variacao_pct"].iloc[0] == 0.4275
    assert pd.isna(df.loc[df["ano"] == 2018, "variacao_pct"].iloc[0])


def test_urls_edgar_e_markdown_tem_pagina() -> None:
    row = LINHAS[-1]
    url = edgar_url(row["accession"], row["arquivo"])
    assert url.endswith("pbrfs2q26usd_6k.htm")
    md = escrever_markdown(montar_dataframe(), "teste")
    assert "36.623" in md or "36,623" in md
    assert "F-4" in md
    assert "descontinuadas" in md
    assert "6-K" in md
    assert "253.447" in md
    assert "270.074" in md
    assert "Total 2002–2025" in md
