#!/usr/bin/env python3
"""Demonstrativo FNO 2000-2021: correntes (reversão IGP-DI) e IPCA a 30/06/2026.

Fonte da série a preços constantes de 2021 (IGP-DI):
  Banco da Amazônia, *O impacto do FNO em dados e ciência*, Figura 3 (p. 79).
  Texto: total 2000-2021 = R$ 85,31 bi; 2000 = R$ 871,18 mi; 2021 = R$ 13,32 bi.

Os valores anuais intermediários foram digitalizados das barras do gráfico e
recalibrados para respeitar exatamente esses três âncoras textuais.

Reversão do deflator (valores correntes do ano t):
  V_corrente_t = V_const_2021_t * (IGP-DI_médio_t / IGP-DI_médio_2021)
  IGP-DI: variação mensal BCB SGS 190 → índice acumulado → média anual.

Atualização a 30/06/2026:
  IPCA BCB SGS 433, índice médio do ano da contratação → índice de jun/2026.
"""

from __future__ import annotations

import json
import urllib.request
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT_XLSX = ROOT / "output" / "fno_contratacoes_2000_2021_ipca.xlsx"
OUT_MD = ROOT / "output" / "fno_contratacoes_resumo.md"
DATA_DIR = ROOT / "data" / "fno"

# Alturas das barras digitalizadas (OpenCV, página 79 do PDF), em pixels.
# 2000: componente isolado; 2001-2021: segmentos com baseline comum.
BAR_HEIGHTS = np.array(
    [
        36,
        24,
        26,
        54,
        66,
        50,
        46,
        48,
        94,
        126,
        108,
        88,
        204,
        230,
        268,
        182,
        108,
        146,
        210,
        402,
        457,
        555,
    ],
    dtype=float,
)
YEARS = np.arange(2000, 2022)

# Âncoras do texto da p. 79 (preços constantes de 2021).
ANCHOR_2000 = 871.18e6
ANCHOR_2021 = 13.32e9
ANCHOR_SUM = 85.31e9

# Totais oficiais publicados (R$), para aba de confronto — não substituem o gráfico.
OFFICIAL_CURRENT = {
    2003: 1_075.1e6,
    2004: 1_321.1e6,
    2005: 976.3e6,
    2006: 986.3e6,
    2007: 1_110.0e6,
    2008: 2_053.6e6,
    2009: 2_440.5e6,
    2010: 2_568.7e6,
    2011: 1_869.2e6,
    2012: 4_282.6e6,
    2013: 4_719.2e6,
    2014: 5_356.9e6,
    2019: 7_670.9e6,
    2020: 10_486.0e6,
    2021: 12_497.8e6,
}


def constant_2021_from_chart() -> pd.Series:
    """Digitalização das barras + âncoras textuais exatas (2000, 2021, soma)."""
    h = BAR_HEIGHTS.copy()
    remain = ANCHOR_SUM - ANCHOR_2000 - ANCHOR_2021
    mid = h[1:-1] / h[1:-1].sum() * remain
    vals = np.concatenate([[ANCHOR_2000], mid, [ANCHOR_2021]])
    s = pd.Series(vals, index=YEARS, name="Constante_2021_IGPDI_R$")
    assert abs(s.sum() - ANCHOR_SUM) < 1.0
    assert abs(s.loc[2000] - ANCHOR_2000) < 1.0
    assert abs(s.loc[2021] - ANCHOR_2021) < 1.0
    return s


def fetch_json(url: str, path: Path) -> list:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        return json.loads(path.read_text())
    with urllib.request.urlopen(url, timeout=90) as resp:
        data = json.loads(resp.read().decode())
    path.write_text(json.dumps(data), encoding="utf-8")
    return data


def annual_avg_index_from_monthly_var(
    rows: list, start_year: int, end_year: int
) -> pd.Series:
    df = pd.DataFrame(rows)
    df["date"] = pd.to_datetime(df["data"], dayfirst=True)
    df["var"] = df["valor"].astype(float)
    df = df.sort_values("date").reset_index(drop=True)
    df["factor"] = 1.0 + df["var"] / 100.0
    df["index"] = 100.0 * df["factor"].cumprod() / df["factor"].iloc[0]
    df["year"] = df["date"].dt.year
    ann = df.groupby("year")["index"].mean()
    missing = [y for y in range(start_year, end_year + 1) if y not in ann.index]
    if missing:
        raise RuntimeError(f"Índice anual incompleto: {missing}")
    return ann.loc[start_year:end_year]


def load_igpdi_annual() -> pd.Series:
    rows = fetch_json(
        "https://api.bcb.gov.br/dados/serie/bcdata.sgs.190/dados"
        "?formato=json&dataInicial=01/01/1995&dataFinal=01/12/2021",
        DATA_DIR / "bcb_190_igpdi.json",
    )
    return annual_avg_index_from_monthly_var(rows, 2000, 2021)


def load_ipca_target() -> tuple[pd.Series, float, str]:
    rows = fetch_json(
        "https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados"
        "?formato=json&dataInicial=01/01/1999&dataFinal=01/06/2026",
        ROOT / "data" / "raw" / "bcb_series" / "433_ipca.json",
    )
    df = pd.DataFrame(rows)
    df["date"] = pd.to_datetime(df["data"], dayfirst=True)
    df["var"] = df["valor"].astype(float)
    df = df.sort_values("date").reset_index(drop=True)
    df["factor"] = 1.0 + df["var"] / 100.0
    df["index"] = 100.0 * df["factor"].cumprod() / df["factor"].iloc[0]
    df["year"] = df["date"].dt.year
    ann = df.groupby("year")["index"].mean()
    target = df[(df["date"].dt.year == 2026) & (df["date"].dt.month == 6)]
    if target.empty:
        raise RuntimeError("IPCA jun/2026 não encontrado.")
    return ann, float(target["index"].iloc[0]), "06/2026"


def build() -> dict[str, pd.DataFrame]:
    const = constant_2021_from_chart()
    igpdi = load_igpdi_annual()
    ipca_ann, ipca_jun2026, meta_label = load_ipca_target()

    fator_rev = igpdi / igpdi.loc[2021]
    corrente_rev = const * fator_rev.loc[const.index]
    fator_ipca = ipca_jun2026 / ipca_ann.loc[const.index]
    atual_rev = corrente_rev * fator_ipca

    # Alternativa: níveis do gráfico tratados como correntes (alinham-se melhor
    # aos totais oficiais BASA/SUDAM do que a reversão literal do IGP-DI).
    corrente_alt = const.copy()
    atual_alt = corrente_alt * fator_ipca

    demo = pd.DataFrame(
        {
            "Ano": const.index.astype(int),
            "Valor lido no gráfico (rótulo: const. 2021 IGP-DI) (R$)": const.values,
            "Fator reversão IGP-DI (média ano / média 2021)": fator_rev.loc[
                const.index
            ].values,
            "Contratações correntes via reversão IGP-DI (R$)": corrente_rev.values,
            "Fator IPCA (média do ano → jun/2026)": fator_ipca.values,
            "Atualizado IPCA 30/06/2026 (sobre reversão) (R$)": atual_rev.values,
            "Fonte": (
                "Figura 3 — O impacto do FNO em dados e ciência (BASA), p. 79; "
                "barras digitalizadas + âncoras do texto"
            ),
            "Atualização": f"IPCA BCB SGS 433 até {meta_label}",
        }
    )

    alt = pd.DataFrame(
        {
            "Ano": const.index.astype(int),
            "Contratações correntes (nível do gráfico) (R$)": corrente_alt.values,
            "Fator IPCA (média do ano → jun/2026)": fator_ipca.values,
            "Atualizado IPCA 30/06/2026 (R$)": atual_alt.values,
            "Justificativa": (
                "Os níveis da Figura 3 acompanham de perto os totais correntes "
                "publicados (razão ~1,05–1,27), ao passo que a reversão literal "
                "do IGP-DI gera correntes muito abaixo do oficial nos anos 2000s."
            ),
        }
    )

    ofic_years = sorted(OFFICIAL_CURRENT)
    ofic_vals = pd.Series({y: OFFICIAL_CURRENT[y] for y in ofic_years})
    ofic_ipca = ofic_vals * (ipca_jun2026 / ipca_ann.loc[ofic_vals.index])
    ofic_df = pd.DataFrame(
        {
            "Ano": ofic_vals.index.astype(int),
            "Contratações correntes oficiais (R$)": ofic_vals.values,
            "Atualizado IPCA 30/06/2026 (R$)": ofic_ipca.values,
            "Fonte": "Relatórios SUDAM/MDR / BASA (série parcial de confronto)",
        }
    )

    confront = []
    for ano, ofic in OFFICIAL_CURRENT.items():
        confront.append(
            {
                "Ano": ano,
                "Nível do gráfico (R$)": float(const.loc[ano]),
                "Corrente via reversão IGP-DI (R$)": float(corrente_rev.loc[ano]),
                "Corrente oficial publicada (R$)": ofic,
                "Razão gráfico/oficial": float(const.loc[ano] / ofic),
                "Razão reversão/oficial": float(corrente_rev.loc[ano] / ofic),
            }
        )
    confront_df = pd.DataFrame(confront)

    metodologia = pd.DataFrame(
        [
            {
                "Item": "Documento",
                "Valor": (
                    "https://d1rb2uej4kk1a4.cloudfront.net/bancoamazonia/"
                    "Livro_O_impacto_do_FNO_em_dados_e_ciencia_59d1247343.pdf"
                ),
            },
            {
                "Item": "Figura",
                "Valor": "Figura 3 – Valores contratados do FNO (2000-2021), p. 79",
            },
            {
                "Item": "Rótulo do gráfico",
                "Valor": "R$ a preços constantes de 2021 (deflator IGP-DI)",
            },
            {
                "Item": "Âncoras textuais",
                "Valor": "2000 = R$ 871,18 mi; 2021 = R$ 13,32 bi; soma = R$ 85,31 bi",
            },
            {
                "Item": "Digitalização",
                "Valor": (
                    "Alturas das barras (OpenCV); anos intermediários rateados "
                    "pela altura relativa para fechar exatamente as âncoras"
                ),
            },
            {
                "Item": "Aba Demonstrativo",
                "Valor": (
                    "Reversão literal do IGP-DI (BCB SGS 190, média anual) e "
                    "atualização pelo IPCA até 30/06/2026 — pedido do usuário"
                ),
            },
            {
                "Item": "Aba Alternativa_grafico_como_corrente",
                "Valor": (
                    "Usa os níveis digitalizados do gráfico como correntes "
                    "(sem reverter IGP-DI) e atualiza pelo IPCA; recomendada "
                    "para coerência com totais oficiais publicados"
                ),
            },
            {
                "Item": "Aba Oficial_parcial_IPCA",
                "Valor": (
                    "Totais oficiais publicados em relatórios FNO (anos "
                    "disponíveis) atualizados pelo IPCA até 30/06/2026"
                ),
            },
            {
                "Item": "Soma correntes (reversão IGP-DI)",
                "Valor": float(corrente_rev.sum()),
            },
            {
                "Item": "Soma IPCA 30/06/2026 (reversão)",
                "Valor": float(atual_rev.sum()),
            },
            {
                "Item": "Soma correntes (nível do gráfico)",
                "Valor": float(corrente_alt.sum()),
            },
            {
                "Item": "Soma IPCA 30/06/2026 (nível do gráfico)",
                "Valor": float(atual_alt.sum()),
            },
        ]
    )
    return {
        "demo": demo,
        "alt": alt,
        "oficial": ofic_df,
        "confront": confront_df,
        "metodologia": metodologia,
    }


def _bi_frame(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    out = df[["Ano"] + cols].copy()
    for c in cols:
        out[c.replace(" (R$)", " (R$ bi)")] = out[c] / 1e9
    keep = ["Ano"] + [c.replace(" (R$)", " (R$ bi)") for c in cols]
    return out[keep]


def write_outputs(tables: dict[str, pd.DataFrame]) -> None:
    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    demo, alt = tables["demo"], tables["alt"]

    bi_rev = _bi_frame(
        demo,
        [
            "Valor lido no gráfico (rótulo: const. 2021 IGP-DI) (R$)",
            "Contratações correntes via reversão IGP-DI (R$)",
            "Atualizado IPCA 30/06/2026 (sobre reversão) (R$)",
        ],
    )
    bi_alt = _bi_frame(
        alt,
        [
            "Contratações correntes (nível do gráfico) (R$)",
            "Atualizado IPCA 30/06/2026 (R$)",
        ],
    )

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        demo.to_excel(writer, sheet_name="Demonstrativo_reversao_IGPDI", index=False)
        bi_rev.to_excel(writer, sheet_name="Reversao_R$_bi", index=False)
        alt.to_excel(writer, sheet_name="Alternativa_grafico_como_corrente", index=False)
        bi_alt.to_excel(writer, sheet_name="Alternativa_R$_bi", index=False)
        tables["oficial"].to_excel(writer, sheet_name="Oficial_parcial_IPCA", index=False)
        tables["confront"].to_excel(writer, sheet_name="Confronto_oficial", index=False)
        tables["metodologia"].to_excel(writer, sheet_name="Metodologia", index=False)

    lines = [
        "# Contratações do FNO (2000–2021)",
        "",
        "Fonte: Figura 3 do livro *O impacto do FNO em dados e ciência* (BASA), p. 79.",
        "",
        f"Arquivo: `{OUT_XLSX.relative_to(ROOT)}`",
        "",
        "## 1) Pedido: reverter IGP-DI e atualizar pelo IPCA (30/06/2026)",
        "",
        "| Ano | Lido no gráfico (R$ bi) | Corrente após reversão IGP-DI (R$ bi) | IPCA 30/06/2026 (R$ bi) |",
        "|----:|------------------------:|--------------------------------------:|------------------------:|",
    ]
    for _, r in demo.iterrows():
        lines.append(
            f"| {int(r['Ano'])} | "
            f"{r['Valor lido no gráfico (rótulo: const. 2021 IGP-DI) (R$)'] / 1e9:.3f} | "
            f"{r['Contratações correntes via reversão IGP-DI (R$)'] / 1e9:.3f} | "
            f"{r['Atualizado IPCA 30/06/2026 (sobre reversão) (R$)'] / 1e9:.3f} |"
        )
    lines += [
        "",
        f"Soma correntes (reversão): R$ {demo['Contratações correntes via reversão IGP-DI (R$)'].sum() / 1e9:.2f} bi",
        f"Soma IPCA 30/06/2026 (reversão): R$ {demo['Atualizado IPCA 30/06/2026 (sobre reversão) (R$)'].sum() / 1e9:.2f} bi",
        "",
        "## 2) Alternativa recomendada: nível do gráfico como corrente + IPCA",
        "",
        "Os níveis da Figura 3 estão muito mais próximos dos totais correntes "
        "oficiais do FNO do que a série obtida ao reverter o IGP-DI.",
        "",
        "| Ano | Corrente = nível do gráfico (R$ bi) | IPCA 30/06/2026 (R$ bi) |",
        "|----:|------------------------------------:|------------------------:|",
    ]
    for _, r in alt.iterrows():
        lines.append(
            f"| {int(r['Ano'])} | "
            f"{r['Contratações correntes (nível do gráfico) (R$)'] / 1e9:.3f} | "
            f"{r['Atualizado IPCA 30/06/2026 (R$)'] / 1e9:.3f} |"
        )
    lines += [
        "",
        f"Soma correntes (nível gráfico): R$ {alt['Contratações correntes (nível do gráfico) (R$)'].sum() / 1e9:.2f} bi",
        f"Soma IPCA 30/06/2026 (nível gráfico): R$ {alt['Atualizado IPCA 30/06/2026 (R$)'].sum() / 1e9:.2f} bi",
    ]
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {OUT_XLSX}")
    print(f"Wrote {OUT_MD}")


if __name__ == "__main__":
    write_outputs(build())
