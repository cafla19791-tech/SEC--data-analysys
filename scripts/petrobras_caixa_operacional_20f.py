#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Discriminativo da geração operacional de caixa da Petrobras
(Forms 20-F 2002–2025 + 6-K 1S2026).

Série anual em US$ milhões: *Net cash provided by operating activities*
do fluxo de caixa consolidado do 20-F original (CIK 0001119639).
2026 ainda não tem 20-F; usa o 6-K das demonstrações de 30/06/2026 (jan–jun).

Uso::

  python scripts/petrobras_caixa_operacional_20f.py
  python scripts/petrobras_caixa_operacional_20f.py --saida-dir output
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
STEM = "petrobras_caixa_operacional_20f_2002_2026"


def edgar_url(accession: str, arquivo: str) -> str:
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{acc}/{arquivo}"


def index_url(accession: str) -> str:
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{acc}/{accession}-index.htm"


# Geração operacional de caixa do próprio 20-F do exercício, salvo 2026 (6-K 1S).
# Métrica: Net cash provided by operating activities (CFS consolidado).
LINHAS: list[dict] = [
    {
        "ano": 2002, "tipo": "20-F", "data_protocolo": "2003-06-19",
        "accession": "0000950123-03-007204", "arquivo": "y87469e20vf.htm",
        "caixa_operacional_usd_milhoes": 6287, "periodo": "ano", "pagina": "F-7",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": (
            "Net cash provided by operating activities was U.S.$6,287 million for 2002, "
            "as compared to U.S.$8,743 million for 2001; CFS F-7: 6,287 [2002]"
        ),
    },
    {
        "ano": 2003, "tipo": "20-F", "data_protocolo": "2004-06-30",
        "accession": "0001193125-04-112315", "arquivo": "d20f.htm",
        "caixa_operacional_usd_milhoes": 8569, "periodo": "ano", "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": "Operating activities provided net cash flows of U.S.$8,569 million in 2003; CFS: 8,569 [2003]",
    },
    {
        "ano": 2004, "tipo": "20-F", "data_protocolo": "2005-06-30",
        "accession": "0001193125-05-135283", "arquivo": "d20f.htm",
        "caixa_operacional_usd_milhoes": 8833, "periodo": "ano", "pagina": "F-9",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": (
            "Operating activities provided net cash flows of U.S.$ 8,833 million in 2004; "
            "CFS: 8,833 [2004]. O 20-F de 2005 reapresenta 2004 como 8,155 — usa-se o próprio ano."
        ),
    },
    {
        "ano": 2005, "tipo": "20-F", "data_protocolo": "2006-06-28",
        "accession": "0000950123-06-008263", "arquivo": "y22597e20vf.htm",
        "caixa_operacional_usd_milhoes": 15115, "periodo": "ano", "pagina": "F-7",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": "Net cash provided by operating activities 15,115 [2005]  8,155 [2004]  8,569 [2003]",
    },
    {
        "ano": 2006, "tipo": "20-F", "data_protocolo": "2007-06-26",
        "accession": "0000950123-07-009192", "arquivo": "y36368e20vf.htm",
        "caixa_operacional_usd_milhoes": 21077, "periodo": "ano", "pagina": "F-7",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": "Net cash provided by operating activities 21,077 [2006]  15,115 [2005]  8,155 [2004]",
    },
    {
        "ano": 2007, "tipo": "20-F", "data_protocolo": "2008-05-19",
        "accession": "0001362310-08-002879", "arquivo": "c73239e20vf.htm",
        "caixa_operacional_usd_milhoes": 22664, "periodo": "ano", "pagina": "F-12",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": "Net cash provided by operating activities 22,664 [2007]  21,077 [2006]  15,115 [2005]",
    },
    {
        "ano": 2008, "tipo": "20-F", "data_protocolo": "2009-05-22",
        "accession": "0000950123-09-009383", "arquivo": "y76586e20vf.htm",
        "caixa_operacional_usd_milhoes": 28220, "periodo": "ano", "pagina": "F-8",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": "Net cash provided by operating activities 28,220 [2008]  22,664 [2007]  21,077 [2006]",
    },
    {
        "ano": 2009, "tipo": "20-F", "data_protocolo": "2010-05-20",
        "accession": "0001292814-10-001665", "arquivo": "pbraform20f2009.htm",
        "caixa_operacional_usd_milhoes": 24920, "periodo": "ano", "pagina": "F-9",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": "Net cash provided by operating activities 24,920 [2009]  28,220 [2008]  22,664 [2007]",
    },
    {
        "ano": 2010, "tipo": "20-F", "data_protocolo": "2011-05-26",
        "accession": "0001292814-11-001552", "arquivo": "pbraform20f2010.htm",
        "caixa_operacional_usd_milhoes": 28495, "periodo": "ano", "pagina": "F-9",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (US GAAP)",
        "norma": "US GAAP",
        "trecho": (
            "Net cash provided by operating activities 28,495 [2010]  24,920 [2009]. "
            "O 20-F de 2011 (IFRS) reapresenta 2010 como 30,110 — usa-se o US GAAP do próprio ano."
        ),
    },
    {
        "ano": 2011, "tipo": "20-F", "data_protocolo": "2012-04-02",
        "accession": "0001292814-12-000786", "arquivo": "pbraform20f_2011.htm",
        "caixa_operacional_usd_milhoes": 33698, "periodo": "ano", "pagina": "F-10",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities 33,698 [2011]; MD&A compara a 30,110 (2010 IFRS)",
    },
    {
        "ano": 2012, "tipo": "20-F", "data_protocolo": "2013-04-29",
        "accession": "0001292814-13-000928", "arquivo": "pbraform20f_2012.htm",
        "caixa_operacional_usd_milhoes": 27888, "periodo": "ano", "pagina": "F-10",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Operating activities provided net cash flows of U.S.$27,888 million for 2012 compared to U.S.$33,698 million for 2011",
    },
    {
        "ano": 2013, "tipo": "20-F", "data_protocolo": "2014-04-30",
        "accession": "0001292814-14-001060", "arquivo": "pbraform20f_2013.htm",
        "caixa_operacional_usd_milhoes": 26289, "periodo": "ano", "pagina": "F-9",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities 26,289 [2013]",
    },
    {
        "ano": 2014, "tipo": "20-F", "data_protocolo": "2015-05-15",
        "accession": "0001292814-15-001242", "arquivo": "pbraform20f_2014.htm",
        "caixa_operacional_usd_milhoes": 26632, "periodo": "ano", "pagina": "F-7",
        "secao": "Item 5 — Liquidity; Consolidated Statements of Cash Flows",
        "metrica": "Cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": (
            "cash provided by operating activities (amounting to U.S.$26,632 million); "
            "cash flow from operations (U.S.$26,632 million)"
        ),
    },
    {
        "ano": 2015, "tipo": "20-F", "data_protocolo": "2016-04-28",
        "accession": "0001292814-16-004364", "arquivo": "pbraform20f_2015.htm",
        "caixa_operacional_usd_milhoes": 25913, "periodo": "ano", "pagina": "F-7",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities 25,913 [2015]",
    },
    {
        "ano": 2016, "tipo": "20-F", "data_protocolo": "2017-04-27",
        "accession": "0001193125-17-140235", "arquivo": "d375139d20f.htm",
        "caixa_operacional_usd_milhoes": 26114, "periodo": "ano", "pagina": "F-8",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities 26,114 [2016]",
    },
    {
        "ano": 2017, "tipo": "20-F", "data_protocolo": "2018-04-18",
        "accession": "0001193125-18-120259", "arquivo": "d521855d20f.htm",
        "caixa_operacional_usd_milhoes": 27112, "periodo": "ano", "pagina": "F-11",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities increased by 4% to US$27,112 million from US$26,114; CFS: 27,112 [2017]",
    },
    {
        "ano": 2018, "tipo": "20-F", "data_protocolo": "2019-04-01",
        "accession": "0001193125-19-093231", "arquivo": "d692671d20f.htm",
        "caixa_operacional_usd_milhoes": 26353, "periodo": "ano", "pagina": "F-11",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities decreased by 3% to US$26,353 million from US$27,112 million",
    },
    {
        "ano": 2019, "tipo": "20-F", "data_protocolo": "2020-03-23",
        "accession": "0001193125-20-080953", "arquivo": "d883642d20f.htm",
        "caixa_operacional_usd_milhoes": 25600, "periodo": "ano", "pagina": "F-13",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities decreased 2.9% to US$25,600 million in 2019, from US$26,353 million in 2018",
    },
    {
        "ano": 2020, "tipo": "20-F", "data_protocolo": "2021-03-25",
        "accession": "0001292814-21-001152", "arquivo": "pbraform20f_2020.htm",
        "caixa_operacional_usd_milhoes": 28890, "periodo": "ano", "pagina": "F-12",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities was US$28,890 million in 2020, an increase of 13% from US$25,600 million in 2019",
    },
    {
        "ano": 2021, "tipo": "20-F", "data_protocolo": "2022-03-30",
        "accession": "0001292814-22-001285", "arquivo": "pbraform20f_2021.htm",
        "caixa_operacional_usd_milhoes": 37791, "periodo": "ano", "pagina": "F-12",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities was US$37,791 million in 2021, an increase of 31% from US$28,890 million in 2020",
    },
    {
        "ano": 2022, "tipo": "20-F", "data_protocolo": "2023-03-29",
        "accession": "0001292814-23-001253", "arquivo": "pbrform20f_2022.htm",
        "caixa_operacional_usd_milhoes": 49717, "periodo": "ano", "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "cash from operating activities of US$49,717 million; CFS: 49,717 [2022]  37,791 [2021]",
    },
    {
        "ano": 2023, "tipo": "20-F", "data_protocolo": "2024-04-12",
        "accession": "0001292814-24-001340", "arquivo": "pbrform20f_2023.htm",
        "caixa_operacional_usd_milhoes": 43212, "periodo": "ano", "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities was US$43,212 million in 2023, a decrease of 13.1% from US$49,717 million in 2022",
    },
    {
        "ano": 2024, "tipo": "20-F", "data_protocolo": "2025-04-03",
        "accession": "0001292814-25-001352", "arquivo": "pbrform20f_2024.htm",
        "caixa_operacional_usd_milhoes": 37984, "periodo": "ano", "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "cash provided by operating activities was US$37,984 million in 2024, a decrease of 12% from US$43,212 million in 2023",
    },
    {
        "ano": 2025, "tipo": "20-F", "data_protocolo": "2026-04-09",
        "accession": "0001292814-26-002168", "arquivo": "pbrform20f_2025.htm",
        "caixa_operacional_usd_milhoes": 36047, "periodo": "ano", "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (IFRS)",
        "norma": "IFRS",
        "trecho": "Net cash provided by operating activities was US$36,047 million in 2025, a decrease of 5.1% from US$ 37,984 million in 2024",
    },
    {
        "ano": 2026, "tipo": "6-K", "data_protocolo": "2026-08-07",
        "accession": "0001292814-26-004133", "arquivo": "pbrfs2q26usd_6k.htm",
        "caixa_operacional_usd_milhoes": 20649, "periodo": "1S (jan–jun)", "pagina": "6",
        "secao": "Unaudited Condensed Consolidated Statements of Cash Flows",
        "metrica": "Net cash provided by operating activities (1S2026)",
        "norma": "IFRS",
        "trecho": (
            "Net cash provided by operating activities  20,649 [Jan-Jun/2026]  "
            "16,029 [Jan-Jun/2025]. Ano de 2026 incompleto — não há 20-F."
        ),
    },
]


def montar_dataframe(linhas: list[dict] | None = None) -> pd.DataFrame:
    rows = []
    prev = None
    for item in linhas or LINHAS:
        valor = item["caixa_operacional_usd_milhoes"]
        anual = item["periodo"] == "ano"
        var_abs = None if (prev is None or not anual) else valor - prev
        var_pct = None if prev in (None, 0) or not anual else (valor / prev - 1.0)
        rows.append(
            {
                "ano": item["ano"],
                "periodo": item["periodo"],
                "tipo": item["tipo"],
                "data_protocolo": item["data_protocolo"],
                "caixa_operacional_usd_milhoes": valor,
                "variacao_usd_milhoes": var_abs,
                "variacao_pct": None if var_pct is None else round(var_pct, 4),
                "pagina": item["pagina"],
                "secao": item["secao"],
                "metrica": item["metrica"],
                "norma": item["norma"],
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
    return f"{int(valor):,}".replace(",", ".")


def _fmt_pct(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    return f"{valor * 100:+.1f}%"


def escrever_markdown(df: pd.DataFrame, gerado: str) -> str:
    linhas = [
        "# Discriminativo — Geração operacional de caixa da Petrobras (20-F 2002–2025 e 6-K 1S2026)",
        "",
        f"**Gerado em:** {gerado}",
        "",
        "Valores em **US$ milhões**, linha **Net cash provided by operating activities** "
        "da demonstração dos fluxos de caixa consolidada do Form 20-F original de "
        "Petróleo Brasileiro S.A. — Petrobras (CIK 0001119639). Não é EBITDA nem "
        "lucro líquido: é caixa gerado pelas operações no exercício.",
        "",
        "Lista EDGAR: [20-F](https://www.sec.gov/cgi-bin/browse-edgar?"
        "action=getcompany&CIK=0001119639&type=20-F&dateb=&owner=exclude&count=100).",
        "",
        "## Como ler a série",
        "",
        "- **2002–2010 (US GAAP):** *Net cash provided by operating activities* no CFS "
        "do próprio 20-F. O 20-F de 2005 reapresenta 2004 como 8.155; a série usa "
        "8.833 do 20-F de 2004. O 20-F de 2011 (IFRS) reapresenta 2010 como 30.110; "
        "a série usa 28.495 do 20-F de 2010.",
        "- **2011–2025 (IFRS):** a mesma linha no CFS IFRS. Pico em 2022 "
        "(US$ 49.717 milhões).",
        "- **2026:** ainda **não existe 20-F**. O valor é o 6-K das demonstrações "
        "interinas de **30/06/2026** (janeiro–junho): US$ 20.649 milhões vs. "
        "US$ 16.029 milhões no 1S2025.",
        "",
        "A coluna **Página** é o folio do CFS (F-N) ou a página 6 do 6-K 2T26.",
        "",
        "## Evolução",
        "",
        "| Ano | Período | Protocolo | Caixa operacional (US$ mi) | Δ US$ mi | Δ % | Página | Norma | Documento |",
        "|----:|---------|-----------|---------------------------:|---------:|----:|--------|-------|-----------|",
    ]
    for r in df.itertuples(index=False):
        linhas.append(
            f"| {r.ano} | {r.periodo} | {r.data_protocolo} | "
            f"{_fmt_mi(r.caixa_operacional_usd_milhoes)} | "
            f"{_fmt_mi(r.variacao_usd_milhoes)} | {_fmt_pct(r.variacao_pct)} | "
            f"{r.pagina} | {r.norma} | [{r.tipo}]({r.url_documento}) |"
        )
    anuais = df[df["periodo"] == "ano"]
    parcial = df[df["periodo"] != "ano"]
    total_anos = int(anuais["caixa_operacional_usd_milhoes"].sum())
    total_com_1s = total_anos + int(parcial["caixa_operacional_usd_milhoes"].sum())
    linhas.append(
        f"| **Total 2002–2025** | 24 anos | — | **{_fmt_mi(total_anos)}** | — | — | — | "
        f"soma US GAAP+IFRS | — |"
    )
    linhas.append(
        f"| **Total + 1S2026** | 24 anos + 1S | — | **{_fmt_mi(total_com_1s)}** | — | — | — | "
        f"inclui 6-K incompleto | — |"
    )
    pico = anuais.loc[anuais["caixa_operacional_usd_milhoes"].idxmax()]
    vale = anuais.loc[anuais["caixa_operacional_usd_milhoes"].idxmin()]
    linhas.extend(
        [
            "",
            f"**Total 2002–2025 (anos completos):** US$ {_fmt_mi(total_anos)} milhões. "
            f"**Total incluindo 1S2026:** US$ {_fmt_mi(total_com_1s)} milhões.",
            f"Pico (anos completos): **US$ {_fmt_mi(pico.caixa_operacional_usd_milhoes)} milhões** "
            f"em {int(pico.ano)} (página {pico.pagina}).",
            f"Mínimo (anos completos): **US$ {_fmt_mi(vale.caixa_operacional_usd_milhoes)} milhões** "
            f"em {int(vale.ano)} (página {vale.pagina}).",
            "",
            "## Localização no formulário (página e trecho)",
            "",
        ]
    )
    for r in df.itertuples(index=False):
        linhas.extend(
            [
                f"### {r.ano} ({r.periodo}) — US$ {_fmt_mi(r.caixa_operacional_usd_milhoes)} milhões",
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
            "SEC EDGAR. 2002–2025: Form 20-F anual, fluxo de caixa consolidado. "
            "2026: Form 6-K de 07/08/2026 (demonstrações em US$ do 2º trimestre).",
            "",
        ]
    )
    return "\n".join(linhas)


def gerar_grafico(df: pd.DataFrame, destino: Path) -> Path:
    cores, hatches = [], []
    for row in df.itertuples(index=False):
        if row.periodo != "ano":
            cores.append("#7a9bb8")
            hatches.append("//")
        elif row.ano <= 2010:
            cores.append("#4c78a8")
            hatches.append("")
        else:
            cores.append("#54a24b")
            hatches.append("")

    fig, ax = plt.subplots(figsize=(14, 6.2))
    bars = ax.bar(
        df["ano"].astype(str),
        df["caixa_operacional_usd_milhoes"],
        color=cores,
        edgecolor="#1f2a37",
        linewidth=0.4,
    )
    for bar, hatch in zip(bars, hatches, strict=True):
        bar.set_hatch(hatch)

    ax.set_title("Petrobras — geração operacional de caixa, 2002–2025 e 1S2026")
    ax.set_xlabel("Exercício")
    ax.set_ylabel("US$ milhões")
    ax.axvline(x=8.5, color="#888888", linestyle=":", linewidth=0.8)
    y_txt = ax.get_ylim()[1] * 0.92
    ax.text(4, y_txt, "US GAAP", ha="center", fontsize=8, color="#555555")
    ax.text(16, y_txt, "IFRS", ha="center", fontsize=8, color="#555555")
    ax.annotate(
        "1S 2026\n(6-K)",
        xy=(24, 20649),
        xytext=(21.5, 38000),
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
    print(df[["ano", "periodo", "caixa_operacional_usd_milhoes", "pagina"]].to_string(index=False))
    for k, path in paths.items():
        print(f"{k}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
