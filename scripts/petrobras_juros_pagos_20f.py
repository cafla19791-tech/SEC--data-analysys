#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Discriminativo dos juros pagos pela Petrobras (Forms 20-F 2002–2025 + 6-K 1S2026).

Série anual em US$ milhões, caixa do exercício, extraída do 20-F original
(CIK 0001119639). 2026 ainda não tem 20-F; usa o 6-K das demonstrações
interinas de 30/06/2026 (jan–jun).

Uso::

  python scripts/petrobras_juros_pagos_20f.py
  python scripts/petrobras_juros_pagos_20f.py --saida-dir output
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
STEM = "petrobras_juros_pagos_20f_2002_2026"


def edgar_url(accession: str, arquivo: str) -> str:
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{acc}/{arquivo}"


def index_url(accession: str) -> str:
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{acc}/{accession}-index.htm"


# Juros pagos (caixa) do próprio 20-F do exercício, salvo 2026 (6-K 1S).
LINHAS: list[dict] = [
    {
        "ano": 2002,
        "tipo": "20-F",
        "data_protocolo": "2003-06-19",
        "accession": "0000950123-03-007204",
        "arquivo": "y87469e20vf.htm",
        "juros_pagos_usd_milhoes": 200,
        "periodo": "ano",
        "pagina": "F-8",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid during the period for Interest",
        "trecho": "Cash paid during the period for Interest 200  [2002]  393 [2001]  622 [2000]",
    },
    {
        "ano": 2003,
        "tipo": "20-F",
        "data_protocolo": "2004-06-30",
        "accession": "0001193125-04-112315",
        "arquivo": "d20f.htm",
        "juros_pagos_usd_milhoes": 622,
        "periodo": "ano",
        "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid during the year for Interest",
        "trecho": "Supplemental cash flow information: Cash paid during the year for Interest 622  [2003]  200 [2002]",
    },
    {
        "ano": 2004,
        "tipo": "20-F",
        "data_protocolo": "2005-06-30",
        "accession": "0001193125-05-135283",
        "arquivo": "d20f.htm",
        "juros_pagos_usd_milhoes": 995,
        "periodo": "ano",
        "pagina": "F-9",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid for Interest, net of amount capitalized",
        "trecho": "Cash paid during the year for Interest, net of amount capitalized 995  [2004]  622 [2003]",
    },
    {
        "ano": 2005,
        "tipo": "20-F",
        "data_protocolo": "2006-06-28",
        "accession": "0000950123-06-008263",
        "arquivo": "y22597e20vf.htm",
        "juros_pagos_usd_milhoes": 1083,
        "periodo": "ano",
        "pagina": "F-7",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid for Interest, net of amount capitalized",
        "trecho": "Cash paid during the year for Interest, net of amount capitalized 1,083  [2005]  995 [2004]",
    },
    {
        "ano": 2006,
        "tipo": "20-F",
        "data_protocolo": "2007-06-26",
        "accession": "0000950123-07-009192",
        "arquivo": "y36368e20vf.htm",
        "juros_pagos_usd_milhoes": 877,
        "periodo": "ano",
        "pagina": "F-7",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid for Interest, net of amount capitalized",
        "trecho": "Interest, net of amount capitalized 877  [2006]  1,083 [2005]",
    },
    {
        "ano": 2007,
        "tipo": "20-F",
        "data_protocolo": "2008-05-19",
        "accession": "0001362310-08-002879",
        "arquivo": "c73239e20vf.htm",
        "juros_pagos_usd_milhoes": 1684,
        "periodo": "ano",
        "pagina": "F-12",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid for Interest, net of amount capitalized",
        "trecho": "Interest, net of amount capitalized 1,684  [2007]  877 [2006]",
    },
    {
        "ano": 2008,
        "tipo": "20-F",
        "data_protocolo": "2009-05-22",
        "accession": "0000950123-09-009383",
        "arquivo": "y76586e20vf.htm",
        "juros_pagos_usd_milhoes": 1515,
        "periodo": "ano",
        "pagina": "F-8",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid for Interest, net of amount capitalized",
        "trecho": (
            "Interest, net of amount capitalized 1,515  [2008]  1,684 [2007]. "
            "O 20-F de 2009 reapresenta 2008 como 2,304"
        ),
    },
    {
        "ano": 2009,
        "tipo": "20-F",
        "data_protocolo": "2010-05-20",
        "accession": "0001292814-10-001665",
        "arquivo": "pbraform20f2009.htm",
        "juros_pagos_usd_milhoes": 3059,
        "periodo": "ano",
        "pagina": "F-9",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid for Interest, net of amount capitalized",
        "trecho": "Interest, net of amount capitalized 3,059  [2009]  2,304 [2008 restated]",
    },
    {
        "ano": 2010,
        "tipo": "20-F",
        "data_protocolo": "2011-05-26",
        "accession": "0001292814-11-001552",
        "arquivo": "pbraform20f2010.htm",
        "juros_pagos_usd_milhoes": 3700,
        "periodo": "ano",
        "pagina": "F-9",
        "secao": "Consolidated Statements of Cash Flows — supplemental",
        "metrica": "Cash paid for Interest, net of amount capitalized",
        "trecho": "Interest, net of amount capitalized 3,700  [2010]  3,059 [2009]",
    },
    {
        "ano": 2011,
        "tipo": "20-F",
        "data_protocolo": "2012-04-02",
        "accession": "0001292814-12-000786",
        "arquivo": "pbraform20f_2011.htm",
        "juros_pagos_usd_milhoes": 4574,
        "periodo": "ano",
        "pagina": "F-10",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of interest",
        "trecho": "Repayment of interest (4,574)  [2011]  (3,659) [2010]  (1,693) [2009]",
    },
    {
        "ano": 2012,
        "tipo": "20-F",
        "data_protocolo": "2013-04-29",
        "accession": "0001292814-13-000928",
        "arquivo": "pbraform20f_2012.htm",
        "juros_pagos_usd_milhoes": 4772,
        "periodo": "ano",
        "pagina": "F-10",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of interest",
        "trecho": "Repayment of interest (4,772)  [2012]  (4,574) [2011]",
    },
    {
        "ano": 2013,
        "tipo": "20-F",
        "data_protocolo": "2014-04-30",
        "accession": "0001292814-14-001060",
        "arquivo": "pbraform20f_2013.htm",
        "juros_pagos_usd_milhoes": 5066,
        "periodo": "ano",
        "pagina": "F-9",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of interest",
        "trecho": "Repayment of interest (5,066)  [2013]  (4,772) [2012]",
    },
    {
        "ano": 2014,
        "tipo": "20-F",
        "data_protocolo": "2015-05-15",
        "accession": "0001292814-15-001242",
        "arquivo": "pbraform20f_2014.htm",
        "juros_pagos_usd_milhoes": 5995,
        "periodo": "ano",
        "pagina": "F-7",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of interest",
        "trecho": "Repayment of interest (5,995)  [2014]  (5,066) [2013]",
    },
    {
        "ano": 2015,
        "tipo": "20-F",
        "data_protocolo": "2016-04-28",
        "accession": "0001292814-16-004364",
        "arquivo": "pbraform20f_2015.htm",
        "juros_pagos_usd_milhoes": 6305,
        "periodo": "ano",
        "pagina": "F-7",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of interest",
        "trecho": "Repayment of interest (6,305)  [2015]  (5,995) [2014]",
    },
    {
        "ano": 2016,
        "tipo": "20-F",
        "data_protocolo": "2017-04-27",
        "accession": "0001193125-17-140235",
        "arquivo": "d375139d20f.htm",
        "juros_pagos_usd_milhoes": 7308,
        "periodo": "ano",
        "pagina": "F-8",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of interest",
        "trecho": "Repayment of interest (7,308)  [2016]  (6,305) [2015]",
    },
    {
        "ano": 2017,
        "tipo": "20-F",
        "data_protocolo": "2018-04-18",
        "accession": "0001193125-18-120259",
        "arquivo": "d521855d20f.htm",
        "juros_pagos_usd_milhoes": 6981,
        "periodo": "ano",
        "pagina": "F-11",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of interest",
        "trecho": "Repayment of interest (6,981)  [2017]  (7,308) [2016]",
    },
    {
        "ano": 2018,
        "tipo": "20-F",
        "data_protocolo": "2019-04-01",
        "accession": "0001193125-19-093231",
        "arquivo": "d692671d20f.htm",
        "juros_pagos_usd_milhoes": 5791,
        "periodo": "ano",
        "pagina": "F-11",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of interest",
        "trecho": "Repayment of interest (5,791)  [2018]  (6,981) [2017]",
    },
    {
        "ano": 2019,
        "tipo": "20-F",
        "data_protocolo": "2020-03-23",
        "accession": "0001193125-20-080953",
        "arquivo": "d883642d20f.htm",
        "juros_pagos_usd_milhoes": 4501,
        "periodo": "ano",
        "pagina": "F-13",
        "secao": "Consolidated Statement of Cash Flows — financing",
        "metrica": "Repayment of finance debt - interest",
        "trecho": "Repayment of finance debt - interest (4,501)  [2019]  (5,703) [2018]  (6,500) [2017]",
    },
    {
        "ano": 2020,
        "tipo": "20-F",
        "data_protocolo": "2021-03-25",
        "accession": "0001292814-21-001152",
        "arquivo": "pbraform20f_2020.htm",
        "juros_pagos_usd_milhoes": 3157,
        "periodo": "ano",
        "pagina": "F-12",
        "secao": "Consolidated Statements of Cash Flows — financing",
        "metrica": "Repayment of interest - finance debt",
        "trecho": "Repayment of interest - finance debt (3,157)  [2020]  (4,501) [2019]",
    },
    {
        "ano": 2021,
        "tipo": "20-F",
        "data_protocolo": "2022-03-30",
        "accession": "0001292814-22-001285",
        "arquivo": "pbraform20f_2021.htm",
        "juros_pagos_usd_milhoes": 2229,
        "periodo": "ano",
        "pagina": "F-12",
        "secao": "Consolidated Statements of Cash Flows — financing (nota 32.2)",
        "metrica": "Repayment of interest - finance debt",
        "trecho": "Repayment of interest - finance debt 32.2 (2,229)  [2021]  (3,157) [2020]",
    },
    {
        "ano": 2022,
        "tipo": "20-F",
        "data_protocolo": "2023-03-29",
        "accession": "0001292814-23-001253",
        "arquivo": "pbrform20f_2022.htm",
        "juros_pagos_usd_milhoes": 1850,
        "periodo": "ano",
        "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows — financing (nota 31.3)",
        "metrica": "Repayment of interest - finance debt",
        "trecho": "Repayment of interest - finance debt 31.3 (1,850)  [2022]  (2,229) [2021]",
    },
    {
        "ano": 2023,
        "tipo": "20-F",
        "data_protocolo": "2024-04-12",
        "accession": "0001292814-24-001340",
        "arquivo": "pbrform20f_2023.htm",
        "juros_pagos_usd_milhoes": 1978,
        "periodo": "ano",
        "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows — financing (nota 32)",
        "metrica": "Repayment of interest - finance debt",
        "trecho": "Repayment of interest - finance debt 32 (1,978)  [2023]  (1,850) [2022]",
    },
    {
        "ano": 2024,
        "tipo": "20-F",
        "data_protocolo": "2025-04-03",
        "accession": "0001292814-25-001352",
        "arquivo": "pbrform20f_2024.htm",
        "juros_pagos_usd_milhoes": 1918,
        "periodo": "ano",
        "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows — financing (nota 30)",
        "metrica": "Repayment of interest - finance debt",
        "trecho": "Repayment of interest - finance debt 30 (1,918)  [2024]  (1,978) [2023]",
    },
    {
        "ano": 2025,
        "tipo": "20-F",
        "data_protocolo": "2026-04-09",
        "accession": "0001292814-26-002168",
        "arquivo": "pbrform20f_2025.htm",
        "juros_pagos_usd_milhoes": 1836,
        "periodo": "ano",
        "pagina": "F-6",
        "secao": "Consolidated Statements of Cash Flows — financing (nota 30)",
        "metrica": "Repayment of interest - finance debt",
        "trecho": "Repayment of interest - finance debt 30 (1,836)  [2025]  (1,918) [2024]",
    },
    {
        "ano": 2026,
        "tipo": "6-K",
        "data_protocolo": "2026-08-07",
        "accession": "0001292814-26-004133",
        "arquivo": "pbrfs2q26usd_6k.htm",
        "juros_pagos_usd_milhoes": 1070,
        "periodo": "1S (jan–jun)",
        "pagina": "6",
        "secao": "Unaudited Condensed Consolidated Statements of Cash Flows (p. 6); nota 24",
        "metrica": "Repayment of interest - finance debt (1S2026)",
        "trecho": (
            "Repayment of interest - finance debt 24  (1,070)  [Jan-Jun/2026]  "
            "(856) [Jan-Jun/2025]. Ano de 2026 incompleto — não há 20-F."
        ),
    },
]


def montar_dataframe(linhas: list[dict] | None = None) -> pd.DataFrame:
    rows = []
    prev = None
    for item in linhas or LINHAS:
        valor = item["juros_pagos_usd_milhoes"]
        anual = item["periodo"] == "ano"
        var_abs = None if (prev is None or not anual) else valor - prev
        var_pct = None if prev in (None, 0) or not anual else (valor / prev - 1.0)
        rows.append(
            {
                "ano": item["ano"],
                "periodo": item["periodo"],
                "tipo": item["tipo"],
                "data_protocolo": item["data_protocolo"],
                "juros_pagos_usd_milhoes": valor,
                "variacao_usd_milhoes": var_abs,
                "variacao_pct": None if var_pct is None else round(var_pct, 4),
                "pagina": item["pagina"],
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
    return f"{int(valor):,}".replace(",", ".")


def _fmt_pct(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    return f"{valor * 100:+.1f}%"


def escrever_markdown(df: pd.DataFrame, gerado: str) -> str:
    linhas = [
        "# Discriminativo — Juros pagos pela Petrobras (20-F 2002–2025 e 6-K 1S2026)",
        "",
        f"**Gerado em:** {gerado}",
        "",
        "Valores em **US$ milhões**, **caixa** do exercício (não despesa de juros "
        "pelo regime de competência). Fonte: Form 20-F original de Petróleo "
        "Brasileiro S.A. — Petrobras (CIK 0001119639), demonstração dos fluxos "
        "de caixa consolidada.",
        "",
        "Lista EDGAR: [20-F](https://www.sec.gov/cgi-bin/browse-edgar?"
        "action=getcompany&CIK=0001119639&type=20-F&dateb=&owner=exclude&count=100).",
        "",
        "## Como ler a série",
        "",
        "- **2002–2003:** *Cash paid during the year/period for Interest* "
        "(nota suplementar do fluxo de caixa).",
        "- **2004–2010:** a mesma linha, rotulada *Interest, net of amount "
        "capitalized* — caixa de juros **líquido** dos custos de empréstimo "
        "capitalizados. Não é estritamente comparável ao bloco seguinte.",
        "- **2011–2025:** *Repayment of interest* / *Repayment of interest - "
        "finance debt* na seção de financiamento do fluxo de caixa — saída "
        "bruta de caixa para juros da dívida financeira.",
        "- **2026:** ainda **não existe 20-F**. O valor é o 6-K das "
        "demonstrações interinas de **30/06/2026** (janeiro–junho): "
        "US$ 1.070 milhões vs. US$ 856 milhões no 1S2025.",
        "",
        "A coluna **Página** é o folio das demonstrações financeiras (F-N) "
        "ou a página 6 do 6-K 2T26.",
        "",
        "## Evolução",
        "",
        "| Ano | Período | Protocolo | Juros pagos (US$ mi) | Δ US$ mi | Δ % | Página | Métrica | Documento |",
        "|----:|---------|-----------|---------------------:|---------:|----:|--------|---------|-----------|",
    ]
    for r in df.itertuples(index=False):
        linhas.append(
            f"| {r.ano} | {r.periodo} | {r.data_protocolo} | "
            f"{_fmt_mi(r.juros_pagos_usd_milhoes)} | "
            f"{_fmt_mi(r.variacao_usd_milhoes)} | {_fmt_pct(r.variacao_pct)} | "
            f"{r.pagina} | {r.metrica} | [{r.tipo}]({r.url_documento}) |"
        )
    anuais = df[df["periodo"] == "ano"]
    parcial = df[df["periodo"] != "ano"]
    total_anos = int(anuais["juros_pagos_usd_milhoes"].sum())
    total_com_1s = total_anos + int(parcial["juros_pagos_usd_milhoes"].sum())
    linhas.append(
        f"| **Total 2002–2025** | 24 anos | — | **{_fmt_mi(total_anos)}** | — | — | — | "
        f"soma dos anos completos | — |"
    )
    linhas.append(
        f"| **Total + 1S2026** | 24 anos + 1S | — | **{_fmt_mi(total_com_1s)}** | — | — | — | "
        f"inclui 6-K incompleto | — |"
    )
    pico = anuais.loc[anuais["juros_pagos_usd_milhoes"].idxmax()]
    vale = anuais.loc[anuais["juros_pagos_usd_milhoes"].idxmin()]
    linhas.extend(
        [
            "",
            f"**Total 2002–2025 (anos completos):** US$ {_fmt_mi(total_anos)} milhões. "
            f"**Total incluindo 1S2026:** US$ {_fmt_mi(total_com_1s)} milhões.",
            f"Pico (anos completos): **US$ {_fmt_mi(pico.juros_pagos_usd_milhoes)} milhões** "
            f"em {int(pico.ano)} (página {pico.pagina}).",
            f"Mínimo (anos completos): **US$ {_fmt_mi(vale.juros_pagos_usd_milhoes)} milhões** "
            f"em {int(vale.ano)} (página {vale.pagina}).",
            "",
            "## Localização no formulário (página e trecho)",
            "",
        ]
    )
    for r in df.itertuples(index=False):
        linhas.extend(
            [
                f"### {r.ano} ({r.periodo}) — US$ {_fmt_mi(r.juros_pagos_usd_milhoes)} milhões",
                "",
                f"- **Página:** {r.pagina}",
                f"- **Seção:** {r.secao}",
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
    """Barras anuais 2002–2025; 2026 hachurado (1º semestre, 6-K)."""
    cores = []
    hatches = []
    for row in df.itertuples(index=False):
        if row.periodo != "ano":
            cores.append("#7a9bb8")
            hatches.append("//")
        elif row.ano <= 2003:
            cores.append("#4c78a8")
            hatches.append("")
        elif row.ano <= 2010:
            cores.append("#f58518")
            hatches.append("")
        else:
            cores.append("#54a24b")
            hatches.append("")

    fig, ax = plt.subplots(figsize=(14, 6.2))
    bars = ax.bar(
        df["ano"].astype(str),
        df["juros_pagos_usd_milhoes"],
        color=cores,
        edgecolor="#1f2a37",
        linewidth=0.4,
    )
    for bar, hatch in zip(bars, hatches, strict=True):
        bar.set_hatch(hatch)

    ax.set_title("Petrobras — juros pagos (caixa), 2002–2025 e 1S2026")
    ax.set_xlabel("Exercício")
    ax.set_ylabel("US$ milhões")
    ax.axvline(x=8.5, color="#888888", linestyle=":", linewidth=0.8)
    y_txt = ax.get_ylim()[1] * 0.92
    ax.text(4, y_txt, "suplementar CFS\n(2004–10: líquido)", ha="center", fontsize=8, color="#555555")
    ax.text(16, y_txt, "Repayment of interest (financiamento)", ha="center", fontsize=8, color="#555555")
    ax.annotate(
        "1S 2026\n(6-K)",
        xy=(24, 1070),
        xytext=(22.2, 2800),
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
    print(df[["ano", "periodo", "juros_pagos_usd_milhoes", "pagina"]].to_string(index=False))
    for k, path in paths.items():
        print(f"{k}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
