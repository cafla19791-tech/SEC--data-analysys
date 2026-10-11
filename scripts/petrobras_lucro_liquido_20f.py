#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Discriminativo do lucro líquido da Petrobras (Forms 20-F 2002–2025 + 6-K 1S2026).

Série anual em US$ milhões, atribuível aos acionistas da Petrobras, extraída
da DRE consolidada do 20-F original (CIK 0001119639). 2026 ainda não tem 20-F;
usa o 6-K das demonstrações interinas de 30/06/2026 (jan–jun).

Uso::

  python scripts/petrobras_lucro_liquido_20f.py
  python scripts/petrobras_lucro_liquido_20f.py --saida-dir output
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CIK = "1119639"
STEM = "petrobras_lucro_liquido_20f_2002_2026"


def edgar_url(accession: str, arquivo: str) -> str:
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{acc}/{arquivo}"


def index_url(accession: str) -> str:
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{acc}/{accession}-index.htm"


# Lucro líquido atribuível aos acionistas, do próprio 20-F do exercício
# (salvo 2026: 6-K 1S). US GAAP até 2010; IFRS a partir de 2011.
LINHAS: list[dict] = [
    {
        "ano": 2002,
        "tipo": "20-F",
        "data_protocolo": "2003-06-19",
        "accession": "0000950123-03-007204",
        "arquivo": "y87469e20vf.htm",
        "lucro_liquido_usd_milhoes": 2311,
        "periodo": "ano",
        "pagina": "F-5",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year",
        "trecho": "Net income for the year 2,311  [2002]  3,491 [2001]  5,342 [2000]",
    },
    {
        "ano": 2003,
        "tipo": "20-F",
        "data_protocolo": "2004-06-30",
        "accession": "0001193125-04-112315",
        "arquivo": "d20f.htm",
        "lucro_liquido_usd_milhoes": 6559,
        "periodo": "ano",
        "pagina": "F-4",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year",
        "trecho": "Net income for the year 6,559  [2003]  2,311 [2002]  3,491 [2001]",
    },
    {
        "ano": 2004,
        "tipo": "20-F",
        "data_protocolo": "2005-06-30",
        "accession": "0001193125-05-135283",
        "arquivo": "d20f.htm",
        "lucro_liquido_usd_milhoes": 6190,
        "periodo": "ano",
        "pagina": "F-7",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year",
        "trecho": "Net income for the year 6,190  [2004]  6,559 [2003]  2,311 [2002]",
    },
    {
        "ano": 2005,
        "tipo": "20-F",
        "data_protocolo": "2006-06-28",
        "accession": "0000950123-06-008263",
        "arquivo": "y22597e20vf.htm",
        "lucro_liquido_usd_milhoes": 10344,
        "periodo": "ano",
        "pagina": "F-5",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year",
        "trecho": "Net income for the year 10,344  [2005]  6,190 [2004]  6,559 [2003]",
    },
    {
        "ano": 2006,
        "tipo": "20-F",
        "data_protocolo": "2007-06-26",
        "accession": "0000950123-07-009192",
        "arquivo": "y36368e20vf.htm",
        "lucro_liquido_usd_milhoes": 12826,
        "periodo": "ano",
        "pagina": "F-5",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year",
        "trecho": "Net income for the year 12,826  [2006]  10,344 [2005]  6,190 [2004]",
    },
    {
        "ano": 2007,
        "tipo": "20-F",
        "data_protocolo": "2008-05-19",
        "accession": "0001362310-08-002879",
        "arquivo": "c73239e20vf.htm",
        "lucro_liquido_usd_milhoes": 13138,
        "periodo": "ano",
        "pagina": "F-10",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year",
        "trecho": "Net income for the year 13,138  [2007]  12,826 [2006]  10,344 [2005]",
    },
    {
        "ano": 2008,
        "tipo": "20-F",
        "data_protocolo": "2009-05-22",
        "accession": "0000950123-09-009383",
        "arquivo": "y76586e20vf.htm",
        "lucro_liquido_usd_milhoes": 18879,
        "periodo": "ano",
        "pagina": "F-6",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year",
        "trecho": "Net income for the year 18,879  [2008]  13,138 [2007]  12,826 [2006]",
    },
    {
        "ano": 2009,
        "tipo": "20-F",
        "data_protocolo": "2010-05-20",
        "accession": "0001292814-10-001665",
        "arquivo": "pbraform20f2009.htm",
        "lucro_liquido_usd_milhoes": 15504,
        "periodo": "ano",
        "pagina": "F-7",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year attributable to Petrobras",
        "trecho": (
            "Net income for the year 16,823; attributable to Petrobras 15,504  "
            "[2009]  18,879 [2008]  13,138 [2007]"
        ),
    },
    {
        "ano": 2010,
        "tipo": "20-F",
        "data_protocolo": "2011-05-26",
        "accession": "0001292814-11-001552",
        "arquivo": "pbraform20f2010.htm",
        "lucro_liquido_usd_milhoes": 19184,
        "periodo": "ano",
        "pagina": "F-7",
        "norma": "US GAAP",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income for the year attributable to Petrobras",
        "trecho": (
            "Net income for the year attributable to Petrobras 19,184  [2010]  "
            "15,504 [2009]  18,879 [2008]. O 20-F de 2011 (1º IFRS) reapresenta "
            "2010 como 20,055"
        ),
    },
    {
        "ano": 2011,
        "tipo": "20-F",
        "data_protocolo": "2012-04-02",
        "accession": "0001292814-12-000786",
        "arquivo": "pbraform20f_2011.htm",
        "lucro_liquido_usd_milhoes": 20121,
        "periodo": "ano",
        "pagina": "F-7",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income attributable to the shareholders of Petrobras",
        "trecho": (
            "Net income attributable to the shareholders of Petrobras 20,121  "
            "[2011]  20,055 [2010 IFRS]  15,308 [2009 IFRS]"
        ),
    },
    {
        "ano": 2012,
        "tipo": "20-F",
        "data_protocolo": "2013-04-29",
        "accession": "0001292814-13-000928",
        "arquivo": "pbraform20f_2012.htm",
        "lucro_liquido_usd_milhoes": 11034,
        "periodo": "ano",
        "pagina": "F-7",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income attributable to the shareholders of Petrobras",
        "trecho": (
            "Net income attributable to the shareholders of Petrobras 11,034  "
            "[2012]  20,121 [2011]  20,055 [2010]"
        ),
    },
    {
        "ano": 2013,
        "tipo": "20-F",
        "data_protocolo": "2014-04-30",
        "accession": "0001292814-14-001060",
        "arquivo": "pbraform20f_2013.htm",
        "lucro_liquido_usd_milhoes": 11094,
        "periodo": "ano",
        "pagina": "F-7",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income attributable to the shareholders of Petrobras",
        "trecho": (
            "Net income attributable to the shareholders of Petrobras 11,094  "
            "[2013]  11,034 [2012]  20,121 [2011]"
        ),
    },
    {
        "ano": 2014,
        "tipo": "20-F",
        "data_protocolo": "2015-05-15",
        "accession": "0001292814-15-001242",
        "arquivo": "pbraform20f_2014.htm",
        "lucro_liquido_usd_milhoes": -7367,
        "periodo": "ano",
        "pagina": "F-5",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income (loss) attributable to the shareholders of Petrobras",
        "trecho": (
            "Net income (loss) attributable to the shareholders of Petrobras "
            "(7,367)  [2014]  11,094 [2013]  11,034 [2012]"
        ),
    },
    {
        "ano": 2015,
        "tipo": "20-F",
        "data_protocolo": "2016-04-28",
        "accession": "0001292814-16-004364",
        "arquivo": "pbraform20f_2015.htm",
        "lucro_liquido_usd_milhoes": -8450,
        "periodo": "ano",
        "pagina": "F-5",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income (loss) attributable to the shareholders of Petrobras",
        "trecho": (
            "Net income (loss) attributable to the shareholders of Petrobras "
            "(8,450)  [2015]  (7,367) [2014]  11,094 [2013]"
        ),
    },
    {
        "ano": 2016,
        "tipo": "20-F",
        "data_protocolo": "2017-04-27",
        "accession": "0001193125-17-140235",
        "arquivo": "d375139d20f.htm",
        "lucro_liquido_usd_milhoes": -4838,
        "periodo": "ano",
        "pagina": "F-6",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income (loss) attributable to the shareholders of Petrobras",
        "trecho": (
            "Net income (loss) attributable to the shareholders of Petrobras "
            "(4,838)  [2016]  (8,450) [2015]  (7,367) [2014]"
        ),
    },
    {
        "ano": 2017,
        "tipo": "20-F",
        "data_protocolo": "2018-04-18",
        "accession": "0001193125-18-120259",
        "arquivo": "d521855d20f.htm",
        "lucro_liquido_usd_milhoes": -91,
        "periodo": "ano",
        "pagina": "F-9",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income (loss) attributable to the shareholders of Petrobras",
        "trecho": (
            "Net income (loss) attributable to the shareholders of Petrobras "
            "(91)  [2017]  (4,838) [2016]  (8,450) [2015]"
        ),
    },
    {
        "ano": 2018,
        "tipo": "20-F",
        "data_protocolo": "2019-04-01",
        "accession": "0001193125-19-093231",
        "arquivo": "d692671d20f.htm",
        "lucro_liquido_usd_milhoes": 7173,
        "periodo": "ano",
        "pagina": "F-9",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income attributable to shareholders of Petrobras",
        "trecho": (
            "Net income attributable to shareholders of Petrobras 7,173  "
            "[2018]  (91) [2017]  (4,838) [2016]"
        ),
    },
    {
        "ano": 2019,
        "tipo": "20-F",
        "data_protocolo": "2020-03-23",
        "accession": "0001193125-20-080953",
        "arquivo": "d883642d20f.htm",
        "lucro_liquido_usd_milhoes": 10151,
        "periodo": "ano",
        "pagina": "F-11",
        "norma": "IFRS",
        "secao": "Consolidated Statement of Income",
        "metrica": "Net income (loss) attributable to shareholders of Petrobras",
        "trecho": (
            "Net income (loss) attributable to shareholders of Petrobras 10,151  "
            "[2019]  7,173 [2018]  (91) [2017]. Inclui operações descontinuadas "
            "2,491 (continuadas 7,660)"
        ),
    },
    {
        "ano": 2020,
        "tipo": "20-F",
        "data_protocolo": "2021-03-25",
        "accession": "0001292814-21-001152",
        "arquivo": "pbraform20f_2020.htm",
        "lucro_liquido_usd_milhoes": 1141,
        "periodo": "ano",
        "pagina": "F-10",
        "norma": "IFRS",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income attributable to shareholders of Petrobras",
        "trecho": (
            "Net income attributable to shareholders of Petrobras 1,141  "
            "[2020]  10,151 [2019]  7,173 [2018]"
        ),
    },
    {
        "ano": 2021,
        "tipo": "20-F",
        "data_protocolo": "2022-03-30",
        "accession": "0001292814-22-001285",
        "arquivo": "pbraform20f_2021.htm",
        "lucro_liquido_usd_milhoes": 19875,
        "periodo": "ano",
        "pagina": "F-10",
        "norma": "IFRS",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income attributable to shareholders of Petrobras",
        "trecho": (
            "Net income attributable to shareholders of Petrobras 19,875  "
            "[2021]  1,141 [2020]  10,151 [2019]"
        ),
    },
    {
        "ano": 2022,
        "tipo": "20-F",
        "data_protocolo": "2023-03-29",
        "accession": "0001292814-23-001253",
        "arquivo": "pbrform20f_2022.htm",
        "lucro_liquido_usd_milhoes": 36623,
        "periodo": "ano",
        "pagina": "F-4",
        "norma": "IFRS",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income attributable to shareholders of Petrobras",
        "trecho": (
            "Net income attributable to shareholders of Petrobras 36,623  "
            "[2022]  19,875 [2021]  1,141 [2020]"
        ),
    },
    {
        "ano": 2023,
        "tipo": "20-F",
        "data_protocolo": "2024-04-12",
        "accession": "0001292814-24-001340",
        "arquivo": "pbrform20f_2023.htm",
        "lucro_liquido_usd_milhoes": 24884,
        "periodo": "ano",
        "pagina": "F-4",
        "norma": "IFRS",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income attributable to shareholders of Petrobras",
        "trecho": (
            "Net income attributable to shareholders of Petrobras 24,884  "
            "[2023]  36,623 [2022]  19,875 [2021]"
        ),
    },
    {
        "ano": 2024,
        "tipo": "20-F",
        "data_protocolo": "2025-04-03",
        "accession": "0001292814-25-001352",
        "arquivo": "pbrform20f_2024.htm",
        "lucro_liquido_usd_milhoes": 7528,
        "periodo": "ano",
        "pagina": "F-4",
        "norma": "IFRS",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income (loss) attributable to shareholders of Petrobras",
        "trecho": (
            "Net income (loss) attributable to shareholders of Petrobras 7,528  "
            "[2024]  24,884 [2023]  36,623 [2022]"
        ),
    },
    {
        "ano": 2025,
        "tipo": "20-F",
        "data_protocolo": "2026-04-09",
        "accession": "0001292814-26-002168",
        "arquivo": "pbrform20f_2025.htm",
        "lucro_liquido_usd_milhoes": 19634,
        "periodo": "ano",
        "pagina": "F-4",
        "norma": "IFRS",
        "secao": "Consolidated Statements of Income",
        "metrica": "Net income attributable to shareholders of Petrobras",
        "trecho": (
            "Net income attributable to shareholders of Petrobras 19,634  "
            "[2025]  7,528 [2024]  24,884 [2023]"
        ),
    },
    {
        "ano": 2026,
        "tipo": "6-K",
        "data_protocolo": "2026-08-07",
        "accession": "0001292814-26-004133",
        "arquivo": "pbrfs2q26usd_6k.htm",
        "lucro_liquido_usd_milhoes": 16627,
        "periodo": "1S (jan–jun)",
        "pagina": "4",
        "norma": "IFRS",
        "secao": "Unaudited Condensed Consolidated Statements of Income (p. 4)",
        "metrica": "Net income attributable to shareholders of Petrobras (1S2026)",
        "trecho": (
            "Net income attributable to shareholders of Petrobras 16,627  "
            "[Jan-Jun/2026]  10,708 [Jan-Jun/2025]. Ano de 2026 incompleto — "
            "não há 20-F."
        ),
    },
]


def _variacao_pct(valor: int, prev: int | None, anual: bool) -> float | None:
    if prev is None or not anual or prev == 0:
        return None
    if (prev < 0) != (valor < 0):
        return None
    return (valor - prev) / abs(prev)


def montar_dataframe(linhas: list[dict] | None = None) -> pd.DataFrame:
    rows = []
    prev = None
    for item in linhas or LINHAS:
        valor = item["lucro_liquido_usd_milhoes"]
        anual = item["periodo"] == "ano"
        var_abs = None if (prev is None or not anual) else valor - prev
        var_pct = _variacao_pct(valor, prev, anual)
        rows.append(
            {
                "ano": item["ano"],
                "periodo": item["periodo"],
                "tipo": item["tipo"],
                "data_protocolo": item["data_protocolo"],
                "lucro_liquido_usd_milhoes": valor,
                "variacao_usd_milhoes": var_abs,
                "variacao_pct": None if var_pct is None else round(var_pct, 4),
                "pagina": item["pagina"],
                "norma": item["norma"],
                "secao": item["secao"],
                "metrica": item["metrica"],
                "trecho": item["trecho"],
                "url_documento": edgar_url(item["accession"], item["arquivo"]),
                "url_indice": index_url(item["accession"]),
                "accession": item["accession"],
            }
        )
        if anual:
            prev = valor
    return pd.DataFrame(rows)


def _fmt_mi(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    n = int(valor)
    sinal = "-" if n < 0 else ""
    return sinal + f"{abs(n):,}".replace(",", ".")


def _fmt_pct(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    return f"{valor * 100:+.1f}%"


def escrever_markdown(df: pd.DataFrame, gerado: str) -> str:
    linhas = [
        "# Discriminativo — Lucro líquido da Petrobras (20-F 2002–2025 e 6-K 1S2026)",
        "",
        f"**Gerado em:** {gerado}",
        "",
        "Valores em **US$ milhões**, **lucro (prejuízo) líquido atribuível aos "
        "acionistas da Petrobras** (não inclui participação de não controladores). "
        "Fonte: Form 20-F original de Petróleo Brasileiro S.A. — Petrobras "
        "(CIK 0001119639), demonstração do resultado consolidada do próprio "
        "exercício (não a reapresentação em 20-F posteriores).",
        "",
        "Lista EDGAR: [20-F](https://www.sec.gov/cgi-bin/browse-edgar?"
        "action=getcompany&CIK=0001119639&type=20-F&dateb=&owner=exclude&count=100).",
        "",
        "## Como ler a série",
        "",
        "- **2002–2008 (US GAAP):** *Net income for the year* na DRE "
        "(já líquido da participação minoritária).",
        "- **2009–2010 (US GAAP):** *Net income for the year attributable to "
        "Petrobras* — o total do grupo é maior (ex.: 2009 16.823 vs. 15.504).",
        "- **2011–2025 (IFRS):** *Net income (loss) attributable to the "
        "shareholders of Petrobras*. O 20-F de 2011 é o primeiro em IFRS e "
        "reapresenta 2010 como US$ 20.055 milhões; a série usa o número "
        "US GAAP do 20-F de 2010 (19.184).",
        "- **2019:** o lucro de 10.151 inclui operações descontinuadas "
        "(BR Distribuidora) de 2.491; o lucro das continuadas foi 7.660.",
        "- **2026:** ainda **não existe 20-F**. O valor é o 6-K das "
        "demonstrações interinas de **30/06/2026** (janeiro–junho): "
        "US$ 16.627 milhões vs. US$ 10.708 milhões no 1S2025.",
        "",
        "A variação percentual usa |ano anterior| no denominador e fica em "
        "branco quando o lucro muda de sinal (prejuízo ↔ lucro) ou quando o "
        "período não é um ano completo.",
        "",
        "A coluna **Página** é o folio da DRE (F-N) ou a página 4 do 6-K 2T26.",
        "",
        "## Evolução",
        "",
        "| Ano | Período | Protocolo | Lucro líquido (US$ mi) | Δ US$ mi | Δ % | Página | Norma | Documento |",
        "|----:|---------|-----------|-----------------------:|---------:|----:|--------|-------|-----------|",
    ]
    for r in df.itertuples(index=False):
        linhas.append(
            f"| {r.ano} | {r.periodo} | {r.data_protocolo} | "
            f"{_fmt_mi(r.lucro_liquido_usd_milhoes)} | "
            f"{_fmt_mi(r.variacao_usd_milhoes)} | {_fmt_pct(r.variacao_pct)} | "
            f"{r.pagina} | {r.norma} | [{r.tipo}]({r.url_documento}) |"
        )
    anuais = df[df["periodo"] == "ano"]
    parcial = df[df["periodo"] != "ano"]
    total_anos = int(anuais["lucro_liquido_usd_milhoes"].sum())
    total_com_1s = total_anos + int(parcial["lucro_liquido_usd_milhoes"].sum())
    linhas.append(
        f"| **Total 2002–2025** | 24 anos | — | **{_fmt_mi(total_anos)}** | — | — | — | "
        f"soma US GAAP+IFRS | — |"
    )
    linhas.append(
        f"| **Total + 1S2026** | 24 anos + 1S | — | **{_fmt_mi(total_com_1s)}** | — | — | — | "
        f"inclui 6-K incompleto | — |"
    )
    pico = anuais.loc[anuais["lucro_liquido_usd_milhoes"].idxmax()]
    vale = anuais.loc[anuais["lucro_liquido_usd_milhoes"].idxmin()]
    linhas.extend(
        [
            "",
            f"**Total 2002–2025 (anos completos):** US$ {_fmt_mi(total_anos)} milhões. "
            f"**Total incluindo 1S2026:** US$ {_fmt_mi(total_com_1s)} milhões.",
            f"Pico (anos completos): **US$ {_fmt_mi(pico.lucro_liquido_usd_milhoes)} milhões** "
            f"em {int(pico.ano)} (página {pico.pagina}).",
            f"Mínimo (anos completos): **US$ {_fmt_mi(vale.lucro_liquido_usd_milhoes)} milhões** "
            f"em {int(vale.ano)} (página {vale.pagina}).",
            "",
            "## Localização no formulário (página e trecho)",
            "",
        ]
    )
    for r in df.itertuples(index=False):
        linhas.extend(
            [
                f"### {r.ano} ({r.periodo}) — US$ {_fmt_mi(r.lucro_liquido_usd_milhoes)} milhões",
                "",
                f"- **Página:** {r.pagina}",
                f"- **Seção:** {r.secao}",
                f"- **Norma:** {r.norma}",
                f"- **Documento:** [HTML do {r.tipo}]({r.url_documento})",
                f"- **Índice EDGAR:** [accession {r.accession}]({r.url_indice})",
                f"- **Trecho:** {r.trecho}",
                "",
            ]
        )
    linhas.extend(
        [
            "## Fonte",
            "",
            "SEC EDGAR. 2002–2025: Form 20-F anual, demonstração do resultado "
            "consolidada. 2026: Form 6-K de 07/08/2026 (demonstrações em US$ "
            "do 2º trimestre).",
            "",
        ]
    )
    return "\n".join(linhas)


def gerar_grafico(df: pd.DataFrame, destino: Path) -> Path:
    """Barras anuais 2002–2025; 2026 hachurado (1º semestre, 6-K)."""
    cores = []
    hatches = []
    for row in df.itertuples(index=False):
        if row.periodo != "ano":
            cores.append("#7a9bb8")
            hatches.append("//")
        elif row.lucro_liquido_usd_milhoes < 0:
            cores.append("#e45756")
            hatches.append("")
        elif row.ano <= 2010:
            cores.append("#4c78a8")
            hatches.append("")
        else:
            cores.append("#54a24b")
            hatches.append("")

    fig, ax = plt.subplots(figsize=(14, 6.4))
    bars = ax.bar(
        df["ano"].astype(str),
        df["lucro_liquido_usd_milhoes"],
        color=cores,
        edgecolor="#1f2a37",
        linewidth=0.4,
    )
    for bar, hatch in zip(bars, hatches, strict=True):
        bar.set_hatch(hatch)

    ax.axhline(0, color="#333333", linewidth=0.8)
    ax.axvline(x=8.5, color="#888888", linestyle=":", linewidth=0.8)
    ax.set_title("Petrobras — lucro líquido atribuível aos acionistas, 2002–2025 e 1S2026")
    ax.set_xlabel("Exercício")
    ax.set_ylabel("US$ milhões")
    y_max = ax.get_ylim()[1]
    ax.text(4, y_max * 0.90, "US GAAP", ha="center", fontsize=8, color="#555555")
    ax.text(16.5, y_max * 0.90, "IFRS", ha="center", fontsize=8, color="#555555")
    ax.annotate(
        "1S 2026\n(6-K)",
        xy=(24, 16627),
        xytext=(21.6, 28000),
        arrowprops={"arrowstyle": "->", "color": "#1f2a37"},
        fontsize=8,
        ha="center",
    )
    ax.tick_params(axis="x", labelrotation=45)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    destino.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(destino, dpi=140)
    plt.close(fig)
    return destino


def escrever_saidas(df: pd.DataFrame, saida_dir: Path) -> dict[str, Path]:
    saida_dir.mkdir(parents=True, exist_ok=True)
    gerado = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    paths = {
        "csv": saida_dir / f"{STEM}.csv",
        "xlsx": saida_dir / f"{STEM}.xlsx",
        "md": saida_dir / f"{STEM}.md",
        "png": saida_dir / f"{STEM}.png",
    }
    df.to_csv(paths["csv"], index=False)
    df.to_excel(paths["xlsx"], index=False)
    paths["md"].write_text(escrever_markdown(df, gerado), encoding="utf-8")
    gerar_grafico(df, paths["png"])
    return paths


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--saida-dir", type=Path, default=ROOT / "output")
    args = p.parse_args()
    df = montar_dataframe()
    paths = escrever_saidas(df, args.saida_dir)
    print(f"Linhas: {len(df)}")
    print(df[["ano", "periodo", "lucro_liquido_usd_milhoes", "pagina"]].to_string(index=False))
    for k, path in paths.items():
        print(f"{k}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
