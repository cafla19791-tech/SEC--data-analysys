"""Testes do discriminativo de juros pagos da Petrobras (20-F / 6-K)."""

from __future__ import annotations

import pandas as pd

from scripts.petrobras_juros_pagos_20f import (
    LINHAS,
    edgar_url,
    escrever_markdown,
    montar_dataframe,
)


def test_serie_cobre_2002_a_2026_e_tem_26_linhas() -> None:
    anos = [linha["ano"] for linha in LINHAS]
    assert anos == list(range(2002, 2027))
    assert len(LINHAS) == 25


def test_valores_e_paginas_ancoram_anos_chave() -> None:
    df = montar_dataframe()
    por_ano = {int(row["ano"]): row for row in df.to_dict(orient="records")}

    assert por_ano[2002]["juros_pagos_usd_milhoes"] == 200
    assert por_ano[2002]["pagina"] == "F-8"
    assert "supplemental" in por_ano[2002]["metrica"].lower() or "Cash paid" in por_ano[2002]["metrica"]

    assert por_ano[2004]["juros_pagos_usd_milhoes"] == 995
    assert "net of amount capitalized" in por_ano[2004]["metrica"]

    assert por_ano[2010]["juros_pagos_usd_milhoes"] == 3700
    assert por_ano[2010]["pagina"] == "F-9"
    assert "000129281411001552" in por_ano[2010]["url_documento"]

    assert por_ano[2011]["juros_pagos_usd_milhoes"] == 4574
    assert "Repayment of interest" in por_ano[2011]["metrica"]

    assert por_ano[2016]["juros_pagos_usd_milhoes"] == 7308
    assert por_ano[2016]["pagina"] == "F-8"

    assert por_ano[2025]["juros_pagos_usd_milhoes"] == 1836
    assert por_ano[2025]["pagina"] == "F-6"
    assert por_ano[2025]["tipo"] == "20-F"


def test_2026_e_parcial_6k_e_nao_tem_variacao_yoy() -> None:
    df = montar_dataframe()
    row = df.loc[df["ano"] == 2026].iloc[0]
    assert row["tipo"] == "6-K"
    assert row["periodo"] == "1S (jan–jun)"
    assert row["juros_pagos_usd_milhoes"] == 1070
    assert row["pagina"] == "6"
    assert pd.isna(row["variacao_pct"])
    assert pd.isna(row["variacao_usd_milhoes"])
    assert "pbrfs2q26usd_6k.htm" in row["url_documento"]
    assert "000129281426004133" in row["url_documento"]


def test_variacao_yoy_so_entre_anos_completos() -> None:
    df = montar_dataframe()
    assert pd.isna(df.loc[df["ano"] == 2002, "variacao_pct"].iloc[0])
    assert df.loc[df["ano"] == 2003, "variacao_pct"].iloc[0] == 2.11
    assert df.loc[df["ano"] == 2016, "variacao_pct"].iloc[0] == 0.1591
    assert df.loc[df["ano"] == 2025, "variacao_pct"].iloc[0] == -0.0428


def test_urls_edgar_e_markdown_tem_pagina() -> None:
    row = LINHAS[-1]
    url = edgar_url(row["accession"], row["arquivo"])
    assert url.endswith("pbrfs2q26usd_6k.htm")
    assert "1119639" in url
    md = escrever_markdown(montar_dataframe(), "teste")
    assert "1.070" in md or "1,070" in md
    assert "F-6" in md
    assert "6-K" in md
    assert "net of amount capitalized" in md
    assert "77.996" in md
    assert "79.066" in md
    assert "Total 2002–2025" in md
