#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Discriminativo da dívida bruta da Petrobras nos Forms 20-F (2002–2025).

Série anual em US$ milhões, posição em 31 de dezembro, extraída do 20-F
original (não da emenda 20-F/A) de Petróleo Brasileiro S.A. — Petrobras,
CIK 0001119639, File Number 001-15106.

Uso::

  python scripts/petrobras_divida_bruta_20f.py
  python scripts/petrobras_divida_bruta_20f.py --saida-dir output
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CIK = "1119639"
STEM = "petrobras_divida_bruta_20f_2002_2025"


def edgar_url(accession: str, arquivo: str) -> str:
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{acc}/{arquivo}"


def index_url(accession: str) -> str:
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{acc}/{accession}-index.htm"


# Cada linha é o número de dívida bruta do próprio 20-F daquele exercício.
# Página = número impresso no HTML (rodapé / “N Table of Contents” / nota F-N).
LINHAS: list[dict] = [
    {
        "ano": 2002,
        "tipo": "20-F",
        "data_protocolo": "2003-06-19",
        "accession": "0000950123-03-007204",
        "arquivo": "y87469e20vf.htm",
        "divida_bruta_usd_milhoes": 14680,
        "pagina": "115",
        "secao": "Item 5 — Total Indebtedness",
        "metrica": "Total debt (ST + LT + project finance + sale and leaseback)",
        "norma": "US GAAP",
        "trecho": (
            "Our total debt (including short-term debt, long-term debt, "
            "project financing and sale and leaseback, including current "
            "portions) increased to U.S.$14,680 million as of December 31, 2002"
        ),
    },
    {
        "ano": 2003,
        "tipo": "20-F",
        "data_protocolo": "2004-06-30",
        "accession": "0001193125-04-112315",
        "arquivo": "d20f.htm",
        "divida_bruta_usd_milhoes": 21890,
        "pagina": "75",
        "secao": "Item 5 — Short-Term / Long-Term Debt; Selected BS",
        "metrica": "Soma BS: ST + parcelas circulantes + LT + project finance + capital leases",
        "norma": "US GAAP",
        "trecho": (
            "At December 31, 2003, our short-term debt (excluding current "
            "portions of long-term obligations) increased to U.S.$1,329 million; "
            "soma com LTD 11,888 + project financings 5,066 + capital leases "
            "1,242 + parcelas circulantes = U.S.$21,890 million"
        ),
    },
    {
        "ano": 2004,
        "tipo": "20-F",
        "data_protocolo": "2005-06-30",
        "accession": "0001193125-05-135283",
        "arquivo": "d20f.htm",
        "divida_bruta_usd_milhoes": 20938,
        "pagina": "4 / 104",
        "secao": "Selected Financial Data (p. 4) e Item 5 (p. 104)",
        "metrica": "Soma BS: ST + parcelas circulantes + LT + project finance + capital leases",
        "norma": "US GAAP",
        "trecho": (
            "Outstanding long-term debt, plus the current portion of our "
            "long-term debt, totaled U.S.$13,344 million at December 31, 2004; "
            "soma com project financings, capital leases e short-term = "
            "U.S.$20,938 million"
        ),
    },
    {
        "ano": 2005,
        "tipo": "20-F",
        "data_protocolo": "2006-06-28",
        "accession": "0000950123-06-008263",
        "arquivo": "y22597e20vf.htm",
        "divida_bruta_usd_milhoes": 21177,
        "pagina": "113",
        "secao": "Item 5 — Long-Term Debt; Selected BS",
        "metrica": "Soma BS: ST + parcelas circulantes + LT + project finance + capital leases",
        "norma": "US GAAP",
        "trecho": (
            "Outstanding long-term debt, plus the current portion of our "
            "long-term debt, totaled U.S.$12,931 million at December 31, 2005; "
            "soma com project financings 3,629 + capital leases 1,015 + "
            "short-term e parcelas = U.S.$21,177 million"
        ),
    },
    {
        "ano": 2006,
        "tipo": "20-F",
        "data_protocolo": "2007-06-26",
        "accession": "0000950123-07-009192",
        "arquivo": "y36368e20vf.htm",
        "divida_bruta_usd_milhoes": 21338,
        "pagina": "6 / 108",
        "secao": "Selected Financial Data (p. 6) e Item 5 (p. 108)",
        "metrica": "Soma BS: ST + parcelas circulantes + LT + project finance + capital leases",
        "norma": "US GAAP",
        "trecho": (
            "On December 31, 2006, our short-term debt (excluding current "
            "portions of long-term obligations) amounted to U.S.$1,293 million; "
            "soma com LTD 10,510 + project financings 4,192 + capital leases "
            "824 + parcelas circulantes = U.S.$21,338 million"
        ),
    },
    {
        "ano": 2007,
        "tipo": "20-F",
        "data_protocolo": "2008-05-19",
        "accession": "0001362310-08-002879",
        "arquivo": "c73239e20vf.htm",
        "divida_bruta_usd_milhoes": 21895,
        "pagina": "8 / F-8",
        "secao": "Selected Financial Data (p. 8) e Consolidated BS (F-8)",
        "metrica": "Soma BS: ST + parcelas circulantes + LT + project finance + capital leases",
        "norma": "US GAAP",
        "trecho": (
            "Short-term debt 1,458 + current LTD 1,273 + current project "
            "financings 1,692 + current capital leases 227 + LTD 12,148 + "
            "project financings 4,586 + capital leases 511 = U.S.$21,895 million"
        ),
    },
    {
        "ano": 2008,
        "tipo": "20-F",
        "data_protocolo": "2009-05-22",
        "accession": "0000950123-09-009383",
        "arquivo": "y76586e20vf.htm",
        "divida_bruta_usd_milhoes": 27351,
        "pagina": "84 / F-4",
        "secao": "Item 5 (p. 84) e Consolidated BS (F-4)",
        "metrica": "Soma BS: ST + parcelas circulantes + LT + project finance + capital leases",
        "norma": "US GAAP",
        "trecho": (
            "On December 31, 2008, our short-term debt (excluding current "
            "portions of long-term debt) amounted to U.S.$2,399 million; "
            "soma com LTD 16,031 + project financings 5,015 + capital leases "
            "344 + parcelas circulantes = U.S.$27,351 million"
        ),
    },
    {
        "ano": 2009,
        "tipo": "20-F",
        "data_protocolo": "2010-05-20",
        "accession": "0001292814-10-001665",
        "arquivo": "pbraform20f2009.htm",
        "divida_bruta_usd_milhoes": 56702,
        "pagina": "92–93",
        "secao": "Item 5 — Short-Term Debt / Long-Term Debt",
        "metrica": "Total short-term debt + total long-term debt",
        "norma": "US GAAP",
        "trecho": (
            "Including the current portion of long-term debt, total short-term "
            "debt was U.S.$8,553 million as of December 31, 2009. Our total "
            "long-term debt amounted to U.S.$48,149 million"
        ),
    },
    {
        "ano": 2010,
        "tipo": "20-F",
        "data_protocolo": "2011-05-26",
        "accession": "0001292814-11-001552",
        "arquivo": "pbraform20f2010.htm",
        "divida_bruta_usd_milhoes": 69431,
        "pagina": "Selected Data / Item 5",
        "secao": "Item 5 — Short-Term Debt / Long-Term Debt",
        "metrica": "Total short-term debt + total long-term debt",
        "norma": "US GAAP / IFRS comparatives",
        "trecho": (
            "On December 31, 2010, our total short-term debt amounted to "
            "U.S.$8,960 million. Our total long-term debt amounted to "
            "U.S.$60,471 million"
        ),
    },
    {
        "ano": 2011,
        "tipo": "20-F",
        "data_protocolo": "2012-04-02",
        "accession": "0001292814-12-000786",
        "arquivo": "pbraform20f_2011.htm",
        "divida_bruta_usd_milhoes": 82927,
        "pagina": "105",
        "secao": "Item 5 — Contractual Obligations",
        "metrica": "Current and non-current debt + finance lease obligations",
        "norma": "IFRS",
        "trecho": (
            "Current and non-current debt obligations 82,785 + Capital "
            "(finance) lease obligations 142 = Total balance sheet items "
            "82,927. Confere com ST U.S.$10,111 million + LT U.S.$72,816 million"
        ),
    },
    {
        "ano": 2012,
        "tipo": "20-F",
        "data_protocolo": "2013-04-29",
        "accession": "0001292814-13-000928",
        "arquivo": "pbraform20f_2012.htm",
        "divida_bruta_usd_milhoes": 95963,
        "pagina": "12 / F-5",
        "secao": "Selected Financial Data (p. 12) e BS (F-5)",
        "metrica": "Current debt + Long-term debt (excl. current portion)",
        "norma": "IFRS",
        "trecho": (
            "Current debt 7,479 + Long-term debt 88,484 = U.S.$95,963 million. "
            "On December 31, 2012, our total short-term debt amounted to "
            "U.S.$7,497 million (inclui arrendamentos financeiros circulantes)"
        ),
    },
    {
        "ano": 2013,
        "tipo": "20-F",
        "data_protocolo": "2014-04-30",
        "accession": "0001292814-14-001060",
        "arquivo": "pbraform20f_2013.htm",
        "divida_bruta_usd_milhoes": 114325,
        "pagina": "Nota — Capital management",
        "secao": "Notas — Net debt / Total debt (current and noncurrent)",
        "metrica": "Total debt (current and noncurrent)",
        "norma": "IFRS",
        "trecho": "Total debt (current and noncurrent) 114,325  [2013]  96,067 [2012]",
    },
    {
        "ano": 2014,
        "tipo": "20-F",
        "data_protocolo": "2015-05-15",
        "accession": "0001292814-15-001242",
        "arquivo": "pbraform20f_2014.htm",
        "divida_bruta_usd_milhoes": 132158,
        "pagina": "Nota — Capital management / Item 3",
        "secao": "Notas — Total debt; Item 3 cita U.S.$132,086 million (c/ juros)",
        "metrica": "Total debt (current and noncurrent)",
        "norma": "IFRS",
        "trecho": (
            "Total debt (current and noncurrent) 132,158. Item 3: Our total "
            "debt (including accrued interest) increased by 16% to U.S.$132,086 "
            "million as of December 31, 2014"
        ),
    },
    {
        "ano": 2015,
        "tipo": "20-F",
        "data_protocolo": "2016-04-28",
        "accession": "0001292814-16-004364",
        "arquivo": "pbraform20f_2015.htm",
        "divida_bruta_usd_milhoes": 126216,
        "pagina": "16 / Nota capital management",
        "secao": "Item 3 (p. 16) e Notas — Total debt",
        "metrica": "Total debt (current and noncurrent)",
        "norma": "IFRS",
        "trecho": (
            "Total debt (current and noncurrent) 126,216. Item 3: Our total "
            "debt (including accrued interest) decreased by 4% to US$126,165 "
            "million as of December 31, 2015"
        ),
    },
    {
        "ano": 2016,
        "tipo": "20-F",
        "data_protocolo": "2017-04-27",
        "accession": "0001193125-17-140235",
        "arquivo": "d375139d20f.htm",
        "divida_bruta_usd_milhoes": 118370,
        "pagina": "F-116",
        "secao": "Notas — Capital management (F-116)",
        "metrica": "Total debt (current and noncurrent)",
        "norma": "IFRS",
        "trecho": "Total debt (current and noncurrent) 118,370  [2016]  126,262 [2015]",
    },
    {
        "ano": 2017,
        "tipo": "20-F",
        "data_protocolo": "2018-04-18",
        "accession": "0001193125-18-120259",
        "arquivo": "d521855d20f.htm",
        "divida_bruta_usd_milhoes": 109275,
        "pagina": "114",
        "secao": "Item 5 — Adjusted EBITDA and Net Debt / Gross Debt",
        "metrica": "Current and non-current debt — Gross Debt",
        "norma": "IFRS (pré-IFRS 16)",
        "trecho": "Current and non-current debt—Gross Debt  …  109,275  [2017]  118,370 [2016] (US$ million)",
    },
    {
        "ano": 2018,
        "tipo": "20-F",
        "data_protocolo": "2019-04-01",
        "accession": "0001193125-19-093231",
        "arquivo": "d692671d20f.htm",
        "divida_bruta_usd_milhoes": 84360,
        "pagina": "114",
        "secao": "Item 5 — Adjusted EBITDA and Net Debt / Gross Debt",
        "metrica": "Current and non-current debt — Gross Debt",
        "norma": "IFRS (pré-IFRS 16)",
        "trecho": "Current and non-current debt – Gross Debt  …  84,360  [2018]  109,275 [2017] (US$ million)",
    },
    {
        "ano": 2019,
        "tipo": "20-F",
        "data_protocolo": "2020-03-23",
        "accession": "0001193125-20-080953",
        "arquivo": "d883642d20f.htm",
        "divida_bruta_usd_milhoes": 87121,
        "pagina": "139",
        "secao": "Item 5 — Liquidity — Debt",
        "metrica": "Gross Debt (finance debt + lease liabilities, IFRS 16)",
        "norma": "IFRS 16",
        "trecho": (
            "Considering the effects of IFRS 16, our gross debt totaled "
            "US$87,121 million"
        ),
    },
    {
        "ano": 2020,
        "tipo": "20-F",
        "data_protocolo": "2021-03-25",
        "accession": "0001292814-21-001152",
        "arquivo": "pbraform20f_2020.htm",
        "divida_bruta_usd_milhoes": 75538,
        "pagina": "Liquidity — Debt / Nota 7",
        "secao": "Item 5 — Debt; Nota 7 Capital management",
        "metrica": "Gross Debt (finance debt + lease liabilities)",
        "norma": "IFRS 16",
        "trecho": (
            "Our Gross Debt totaled US$75,538 million. As of December 31, 2020, "
            "gross debt decreased to US$ 75,538, from US$ 87,121 as of "
            "December 31, 2019"
        ),
    },
    {
        "ano": 2021,
        "tipo": "20-F",
        "data_protocolo": "2022-03-30",
        "accession": "0001292814-22-001285",
        "arquivo": "pbraform20f_2021.htm",
        "divida_bruta_usd_milhoes": 58743,
        "pagina": "F-21",
        "secao": "Nota 7 — Capital management (F-21)",
        "metrica": "Gross Debt (finance debt + lease liabilities)",
        "norma": "IFRS 16",
        "trecho": (
            "As of December 31, 2021, gross debt decreased to US$ 58,743, "
            "from US$ 75,538 as of December 31, 2020"
        ),
    },
    {
        "ano": 2022,
        "tipo": "20-F",
        "data_protocolo": "2023-03-29",
        "accession": "0001292814-23-001253",
        "arquivo": "pbrform20f_2022.htm",
        "divida_bruta_usd_milhoes": 53799,
        "pagina": "204",
        "secao": "Item 5 — Liquidity — Debt (p. 204); Nota 7",
        "metrica": "Gross Debt (finance debt + lease liabilities)",
        "norma": "IFRS 16",
        "trecho": (
            "Our Gross Debt (which represents the sum of current and "
            "non-current finance debt and lease liabilities) totaled "
            "US$53,799 million"
        ),
    },
    {
        "ano": 2023,
        "tipo": "20-F",
        "data_protocolo": "2024-04-12",
        "accession": "0001292814-24-001340",
        "arquivo": "pbrform20f_2023.htm",
        "divida_bruta_usd_milhoes": 62600,
        "pagina": "F-21",
        "secao": "Item 5 — Debt; Nota 7 Capital management (F-21)",
        "metrica": "Gross Debt (finance debt + lease liabilities)",
        "norma": "IFRS 16",
        "trecho": (
            "Our Gross Debt (…) totaled US$62,600 million. As of December 31, "
            "2023, gross debt increased to US$ 62,600, from US$ 53,799 as of "
            "December 31, 2022"
        ),
    },
    {
        "ano": 2024,
        "tipo": "20-F",
        "data_protocolo": "2025-04-03",
        "accession": "0001292814-25-001352",
        "arquivo": "pbrform20f_2024.htm",
        "divida_bruta_usd_milhoes": 60311,
        "pagina": "Nota 7 / Item 5 — Debt",
        "secao": "Item 5 — Debt; Nota 7 Capital management",
        "metrica": "Gross Debt (finance debt + lease liabilities)",
        "norma": "IFRS 16",
        "trecho": (
            "Our Gross Debt (…) totaled US$60,311 million. As of December 31, "
            "2024, gross debt decreased to US$ 60,311, from US$ 62,600 as of "
            "December 31, 2023"
        ),
    },
    {
        "ano": 2025,
        "tipo": "20-F",
        "data_protocolo": "2026-04-09",
        "accession": "0001292814-26-002168",
        "arquivo": "pbrform20f_2025.htm",
        "divida_bruta_usd_milhoes": 69793,
        "pagina": "180 / F-22",
        "secao": "Item 5 — Debt (p. 180); Nota 7 (F-22)",
        "metrica": "Gross Debt (finance debt + lease liabilities)",
        "norma": "IFRS 16",
        "trecho": (
            "Our Gross Debt (which represents the sum of current and "
            "non-current finance debt and lease liabilities) totaled "
            "US$69,793 million"
        ),
    },
]


def montar_dataframe(linhas: list[dict] | None = None) -> pd.DataFrame:
    rows = []
    prev = None
    for item in linhas or LINHAS:
        valor = item["divida_bruta_usd_milhoes"]
        var_abs = None if prev is None else valor - prev
        var_pct = None if prev in (None, 0) else (valor / prev - 1.0)
        rows.append(
            {
                "ano": item["ano"],
                "tipo": item["tipo"],
                "data_protocolo": item["data_protocolo"],
                "divida_bruta_usd_milhoes": valor,
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
        prev = valor
    return pd.DataFrame(rows)


def _fmt_mi(valor) -> str:
    if pd.isna(valor):
        return "—"
    return f"{int(valor):,}".replace(",", ".")


def _fmt_pct(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    return f"{valor * 100:+.1f}%"


def escrever_markdown(df: pd.DataFrame, gerado: str) -> str:
    linhas = [
        "# Discriminativo — Dívida bruta da Petrobras (Forms 20-F, 2002–2025)",
        "",
        f"**Gerado em:** {gerado}",
        "",
        "Valores em **US$ milhões**, posição em **31 de dezembro** do exercício, "
        "extraídos do Form 20-F original de Petróleo Brasileiro S.A. — Petrobras "
        "(CIK 0001119639).",
        "",
        "Lista EDGAR: [todos os 20-F](https://www.sec.gov/cgi-bin/browse-edgar?"
        "action=getcompany&CIK=0001119639&type=20-F&dateb=&owner=exclude&count=100).",
        "",
        "## Como ler a série",
        "",
        "- **2002–2008 (US GAAP):** *total debt* = dívida de curto e longo prazo "
        "+ *project financings* + arrendamentos financeiros (*capital leases*), "
        "incluindo parcelas circulantes — a definição do 20-F de 2002.",
        "- **2009–2012:** soma do *total short-term debt* e do *total long-term "
        "debt* declarados no Item 5 (2011 inclui explicitamente *finance leases*).",
        "- **2013–2018 (IFRS):** *Total debt (current and noncurrent)* / "
        "*Gross Debt* da conciliação de Net Debt — **ainda sem** o recorte "
        "cheio de arrendamentos operacionais da IFRS 16.",
        "- **2019–2025 (IFRS 16):** *Gross Debt* oficial = *finance debt* "
        "circulante e não circulante + *lease liabilities*. O salto de 2018 "
        "para 2019 mistura desalavancagem com a 1ª aplicação da IFRS 16 "
        "(a companhia cita redução de US$ 24 bi vs. 2018 *pro forma* IFRS 16).",
        "",
        "A coluna **Página** é o número impresso no próprio 20-F (rodapé, "
        "“N Table of Contents” ou nota “F-N”). Quando o HTML da SEC não grava "
        "o folio (campo SEQ), usa-se a seção / nota.",
        "",
        "## Evolução",
        "",
        "| Ano | Protocolo | Dívida bruta (US$ mi) | Δ US$ mi | Δ % | Página | Métrica | Documento |",
        "|----:|-----------|----------------------:|---------:|----:|--------|---------|-----------|",
    ]
    for r in df.itertuples(index=False):
        linhas.append(
            f"| {r.ano} | {r.data_protocolo} | {_fmt_mi(r.divida_bruta_usd_milhoes)} | "
            f"{_fmt_mi(r.variacao_usd_milhoes)} | {_fmt_pct(r.variacao_pct)} | "
            f"{r.pagina} | {r.metrica} | [20-F]({r.url_documento}) |"
        )
    pico = df.loc[df["divida_bruta_usd_milhoes"].idxmax()]
    vale_pos2014 = df[df["ano"] >= 2014].loc[
        df[df["ano"] >= 2014]["divida_bruta_usd_milhoes"].idxmin()
    ]
    linhas.extend(
        [
            "",
            f"Pico da série: **US$ {_fmt_mi(pico.divida_bruta_usd_milhoes)} milhões** "
            f"em {int(pico.ano)} (página {pico.pagina}).",
            f"Mínimo após 2014: **US$ {_fmt_mi(vale_pos2014.divida_bruta_usd_milhoes)} milhões** "
            f"em {int(vale_pos2014.ano)} (página {vale_pos2014.pagina}).",
            "",
            "## Localização no 20-F (página e trecho)",
            "",
        ]
    )
    for r in df.itertuples(index=False):
        linhas.extend(
            [
                f"### {r.ano} — US$ {_fmt_mi(r.divida_bruta_usd_milhoes)} milhões",
                "",
                f"- **Página:** {r.pagina}",
                f"- **Seção:** {r.secao}",
                f"- **Norma:** {r.norma}",
                f"- **Documento:** [HTML do 20-F]({r.url_documento})",
                f"- **Índice EDGAR:** [accession {r.accession}]({r.url_indice})",
                f"- **Trecho:** {r.trecho}",
                "",
            ]
        )
    linhas.extend(
        [
            "## Fonte",
            "",
            "SEC EDGAR, Form 20-F anual, CIK 0001119639. Os links abrem o HTML "
            "principal do relatório (não a pasta de exhibits).",
            "",
        ]
    )
    return "\n".join(linhas)


def escrever_saidas(df: pd.DataFrame, saida_dir: Path) -> dict[str, Path]:
    saida_dir.mkdir(parents=True, exist_ok=True)
    gerado = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    paths = {
        "csv": saida_dir / f"{STEM}.csv",
        "xlsx": saida_dir / f"{STEM}.xlsx",
        "md": saida_dir / f"{STEM}.md",
    }
    df.to_csv(paths["csv"], index=False)
    df.to_excel(paths["xlsx"], index=False)
    paths["md"].write_text(escrever_markdown(df, gerado), encoding="utf-8")
    return paths


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--saida-dir", type=Path, default=ROOT / "output")
    args = p.parse_args()
    df = montar_dataframe()
    paths = escrever_saidas(df, args.saida_dir)
    print(f"Linhas: {len(df)}")
    print(df[["ano", "divida_bruta_usd_milhoes", "pagina"]].to_string(index=False))
    for k, path in paths.items():
        print(f"{k}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
