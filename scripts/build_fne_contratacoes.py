#!/usr/bin/env python3
"""Demonstrativo das contratações anuais do FNE (2000-2025).

Extrai/consolida o valor total de contratações do Fundo Constitucional de
Financiamento do Nordeste a partir dos Relatórios de Gestão do FNE (BNB/ETENE)
e, quando o PDF de Gestão não está disponível no portal, dos Relatórios de
Resultados e Impactos / Atividades (RFNE) e Demonstrações Financeiras do BNB.

Atualiza os valores correntes pelo IPCA (BCB SGS 433) até 30/06/2026, usando o
índice médio do ano da contratação (mesma convenção do demonstrativo BNDES).
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
IPCA_PATH = ROOT / "data" / "raw" / "bcb_series" / "433_ipca.json"
OUT_XLSX = ROOT / "output" / "fne_contratacoes_2000_2025_ipca.xlsx"
OUT_MD = ROOT / "output" / "fne_contratacoes_resumo.md"

# Valor total de contratações do FNE no exercício (R$ correntes).
# Preferência: Relatório de Gestão; lacunas preenchidas com RFNE/DF oficiais.
ROWS = [
    {
        "Ano": 2000,
        "Contratações correntes (R$)": 569_257_400.0,
        "Fonte": "Relatório de Gestão do FNE 2000 (BNB/ETENE)",
        "Documento": "gestao_2000",
        "Observações": "Tabela de contratações: R$ 569.257,4 mil.",
    },
    {
        "Ano": 2001,
        "Contratações correntes (R$)": 302_545_900.0,
        "Fonte": "Relatório de Gestão do FNE 2001 (BNB/ETENE)",
        "Documento": "gestao_2001",
        "Observações": "Tabela de contratações: R$ 302.545,9 mil.",
    },
    {
        "Ano": 2002,
        "Contratações correntes (R$)": 254_400_000.0,
        "Fonte": "Relatório de Gestão do FNE 2002 (BNB/ETENE)",
        "Documento": "gestao_2002",
        "Observações": "Texto do relatório: R$ 254,4 milhões.",
    },
    {
        "Ano": 2003,
        "Contratações correntes (R$)": 1_019_173_000.0,
        "Fonte": "Relatório de Gestão do FNE 2003 (BNB/ETENE)",
        "Documento": "gestao_2003",
        "Observações": "Tabela: R$ 1.019.173 mil.",
    },
    {
        "Ano": 2004,
        "Contratações correntes (R$)": 3_208_940_000.0,
        "Fonte": "Relatório de Gestão do FNE 2004 (BNB/ETENE); detalhe RFNE 2004",
        "Documento": "gestao_2004",
        "Observações": (
            "Gestão cita R$ 3,2 bilhões; tabela RFNE de Atividades/Resultados: "
            "R$ 3.208.940 mil."
        ),
    },
    {
        "Ano": 2005,
        "Contratações correntes (R$)": 4_173_934_000.0,
        "Fonte": "Relatório de Gestão do FNE 2005 / RFNE 2005 (BNB)",
        "Documento": "gestao_2005",
        "Observações": "Tabela: R$ 4.173.934 mil (texto arredonda para R$ 4,2 bi).",
    },
    {
        "Ano": 2006,
        "Contratações correntes (R$)": 4_588_182_000.0,
        "Fonte": "RFNE – Relatório de Atividades e Resultados 2006 (DSpace BNB)",
        "Documento": "rfne_2006",
        "Observações": (
            "PDF de Gestão 2006 não listado no content-set ETENE; "
            "RFNE: R$ 4.588.182 mil (texto: R$ 4,6 bi)."
        ),
    },
    {
        "Ano": 2007,
        "Contratações correntes (R$)": 4_246_501_000.0,
        "Fonte": "RFNE 2007 (DSpace BNB) / DF BNB 2008",
        "Documento": "rfne_2007",
        "Observações": "Tabela: R$ 4.246.501 mil.",
    },
    {
        "Ano": 2008,
        "Contratações correntes (R$)": 7_668_595_000.0,
        "Fonte": "RFNE 2008 (DSpace BNB) / Demonstrações Financeiras BNB 2008",
        "Documento": "rfne_2008",
        "Observações": "Tabela: R$ 7.668.595 mil.",
    },
    {
        "Ano": 2009,
        "Contratações correntes (R$)": 9_134_100_000.0,
        "Fonte": "Demonstrações Financeiras BNB 2010 / RFNE 2010",
        "Documento": "df_2010_rfne_2010",
        "Observações": (
            "Consolidado usado nas comparações YoY de 2010: R$ 9.134,1 milhões. "
            "A tabela setorial do RFNE 2009 publica R$ 8.838.768 mil "
            "(diferença de metodologia/revisão posterior)."
        ),
    },
    {
        "Ano": 2010,
        "Contratações correntes (R$)": 10_755_163_000.0,
        "Fonte": "RFNE 2010 (DSpace BNB) / Demonstrações Financeiras BNB 2011",
        "Documento": "rfne_2010",
        "Observações": "Tabela: R$ 10.755.163 mil.",
    },
    {
        "Ano": 2011,
        "Contratações correntes (R$)": 11_090_654_000.0,
        "Fonte": "RFNE 2011 (DSpace BNB) / Demonstrações Financeiras BNB 2011",
        "Documento": "rfne_2011",
        "Observações": "Tabela: R$ 11.090.654 mil.",
    },
    {
        "Ano": 2012,
        "Contratações correntes (R$)": 11_970_187_000.0,
        "Fonte": "RFNE 2012 (DSpace BNB)",
        "Documento": "rfne_2012",
        "Observações": "Tabela: R$ 11.970.187 mil (texto: R$ 11,97 bi).",
    },
    {
        "Ano": 2013,
        "Contratações correntes (R$)": 12_727_523_000.0,
        "Fonte": "RFNE / Relatório de Resultados e Impactos 2013 (BNB/ETENE)",
        "Documento": "resultados_2013",
        "Observações": "Tabela: R$ 12.727.523 mil.",
    },
    {
        "Ano": 2014,
        "Contratações correntes (R$)": 13_453_709_000.0,
        "Fonte": "Relatório de Gestão / Resultados e Impactos FNE 2014",
        "Documento": "gestao_resultados_2014",
        "Observações": "Tabela: R$ 13.453.709 mil (texto: R$ 13,5 bi).",
    },
    {
        "Ano": 2015,
        "Contratações correntes (R$)": 11_495_227_000.0,
        "Fonte": "Relatório de Gestão / Resultados e Impactos FNE 2015",
        "Documento": "gestao_resultados_2015",
        "Observações": "Tabela: R$ 11.495.227 mil (texto: R$ 11,5 bi).",
    },
    {
        "Ano": 2016,
        "Contratações correntes (R$)": 11_240_506_000.0,
        "Fonte": "Relatório de Gestão / Resultados e Impactos FNE 2016",
        "Documento": "gestao_resultados_2016",
        "Observações": "Tabela: R$ 11.240.506 mil (texto: R$ 11,2 bi).",
    },
    {
        "Ano": 2017,
        "Contratações correntes (R$)": 15_970_900_000.0,
        "Fonte": "Relatório de Gestão do FNE 2017 (BNB/ETENE)",
        "Documento": "gestao_2017",
        "Observações": (
            "Total geral R$ 15,9709 bi (Programação Padrão R$ 12,3 bi + "
            "Projetos de Grande Porte de Infraestrutura). Não usar só a "
            "Programação Padrão como total do Fundo."
        ),
    },
    {
        "Ano": 2018,
        "Contratações correntes (R$)": 32_653_300_000.0,
        "Fonte": "Relatório de Resultados e Impactos FNE 2018 (BNB)",
        "Documento": "resultados_2018",
        "Observações": (
            "Gestão 2018 não disponível como PDF único no ETENE (há anexos); "
            "texto: R$ 32,6 bi; tabela setorial: R$ 32.653,3 milhões."
        ),
    },
    {
        "Ano": 2019,
        "Contratações correntes (R$)": 29_558_092_900.44,
        "Fonte": "Relatório de Resultados e Impactos FNE 2019 (BNB)",
        "Documento": "resultados_2019",
        "Observações": "Total Geral das contratações: R$ 29.558.092.900,44.",
    },
    {
        "Ano": 2020,
        "Contratações correntes (R$)": 25_842_698_168.34,
        "Fonte": "Relatório de Resultados e Impactos FNE 2020 (BNB)",
        "Documento": "resultados_2020",
        "Observações": "Total Geral das contratações: R$ 25.842.698.168,34.",
    },
    {
        "Ano": 2021,
        "Contratações correntes (R$)": 25_882_267_471.04,
        "Fonte": "Relatório de Gestão do FNE 2021 (BNB/ETENE)",
        "Documento": "gestao_2021",
        "Observações": "Total geral: R$ 25.882.267.471,04 (texto: R$ 25,9 bi).",
    },
    {
        "Ano": 2022,
        "Contratações correntes (R$)": 32_254_507_000.0,
        "Fonte": "Relatório de Gestão do FNE 2022 (BNB/ETENE)",
        "Documento": "gestao_2022",
        "Observações": "Tabela Total Geral: R$ 32.254.507 mil.",
    },
    {
        "Ano": 2023,
        "Contratações correntes (R$)": 43_673_105_300.0,
        "Fonte": "Relatório de Gestão do FNE 2023 (BNB/ETENE)",
        "Documento": "gestao_2023",
        "Observações": "Tabela Total Geral: R$ 43.673.105,3 mil.",
    },
    {
        "Ano": 2024,
        "Contratações correntes (R$)": 44_805_493_000.0,
        "Fonte": "Relatório de Gestão do FNE 2024 (BNB/ETENE)",
        "Documento": "gestao_2024",
        "Observações": "Tabela/comparativo: R$ 44.805.493 mil (texto: R$ 44,8 bi).",
    },
    {
        "Ano": 2025,
        "Contratações correntes (R$)": 50_199_932_000.0,
        "Fonte": "Relatório de Gestão do FNE 2025 (BNB/ETENE)",
        "Documento": "gestao_2025",
        "Observações": "Tabela Total Geral: R$ 50.199.932 mil (texto: R$ 50,2 bi).",
    },
]


def ensure_ipca(path: Path) -> list:
    """Load local IPCA JSON or download BCB SGS 433 (1999–jun/2026)."""
    need_fetch = True
    raw: list = []
    if path.exists():
        raw = json.loads(path.read_text())
        years = {int(x["data"].split("/")[-1]) for x in raw}
        if 2000 in years and 2026 in years:
            need_fetch = False
    if need_fetch:
        import urllib.request

        url = (
            "https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados"
            "?formato=json&dataInicial=01/01/1999&dataFinal=01/06/2026"
        )
        with urllib.request.urlopen(url, timeout=60) as resp:
            raw = json.loads(resp.read().decode())
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(raw, ensure_ascii=False), encoding="utf-8")
    return raw


def load_ipca_index(path: Path) -> tuple[pd.DataFrame, float, str]:
    raw = ensure_ipca(path)
    ipca = pd.DataFrame(raw)
    ipca["date"] = pd.to_datetime(ipca["data"], dayfirst=True)
    ipca["value"] = ipca["valor"].astype(float)
    ipca = ipca.sort_values("date").reset_index(drop=True)
    ipca["factor"] = 1.0 + ipca["value"] / 100.0
    ipca["index"] = 100.0 * ipca["factor"].cumprod() / ipca["factor"].iloc[0]

    target = ipca[(ipca["date"].dt.year == 2026) & (ipca["date"].dt.month == 6)]
    if target.empty:
        raise RuntimeError("IPCA de junho/2026 não encontrado na série BCB 433.")
    target_idx = float(target["index"].iloc[0])
    target_label = "06/2026"
    return ipca, target_idx, target_label


def build() -> pd.DataFrame:
    ipca, target_idx, target_label = load_ipca_index(IPCA_PATH)
    ipca["Ano"] = ipca["date"].dt.year
    idx_ano = ipca.groupby("Ano", as_index=False)["index"].mean()

    df = pd.DataFrame(ROWS)
    df = df.merge(idx_ano, on="Ano", how="left")
    if df["index"].isna().any():
        missing = df.loc[df["index"].isna(), "Ano"].tolist()
        raise RuntimeError(f"Sem IPCA médio para os anos: {missing}")

    df["Contratações atualizadas IPCA 30/06/2026 (R$)"] = (
        df["Contratações correntes (R$)"] * target_idx / df["index"]
    )
    df["Fator IPCA (média do ano → jun/2026)"] = target_idx / df["index"]
    df["Índice IPCA médio do ano"] = df["index"]
    df["Meta IPCA"] = target_label
    df["Série IPCA"] = "BCB SGS 433"

    cols = [
        "Ano",
        "Contratações correntes (R$)",
        "Contratações atualizadas IPCA 30/06/2026 (R$)",
        "Fator IPCA (média do ano → jun/2026)",
        "Fonte",
        "Documento",
        "Observações",
        "Índice IPCA médio do ano",
        "Meta IPCA",
        "Série IPCA",
    ]
    return df[cols]


def write_outputs(df: pd.DataFrame) -> None:
    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

    resumo = pd.DataFrame(
        [
            {
                "Item": "Período",
                "Valor": "2000-2025",
            },
            {
                "Item": "Soma contratações correntes (R$)",
                "Valor": float(df["Contratações correntes (R$)"].sum()),
            },
            {
                "Item": "Soma contratações atualizadas IPCA 30/06/2026 (R$)",
                "Valor": float(
                    df["Contratações atualizadas IPCA 30/06/2026 (R$)"].sum()
                ),
            },
            {
                "Item": "Atualização",
                "Valor": (
                    "IPCA BCB SGS 433; índice médio do ano da contratação "
                    "atualizado para o índice de jun/2026 (posição 30/06/2026)."
                ),
            },
            {
                "Item": "Fontes primárias",
                "Valor": (
                    "Relatórios de Gestão do FNE (BNB/ETENE); RFNE/Resultados "
                    "(DSpace BNB e portal ETENE); Demonstrações Financeiras BNB "
                    "para conciliação 2007-2011."
                ),
            },
            {
                "Item": "Portal de referência",
                "Valor": "https://bnb.gov.br/web/guest/etene/relatorios-fne",
            },
        ]
    )

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Demonstrativo", index=False)
        resumo.to_excel(writer, sheet_name="Metodologia", index=False)
        # Aba amigável em R$ bilhões
        bi = df[
            [
                "Ano",
                "Contratações correntes (R$)",
                "Contratações atualizadas IPCA 30/06/2026 (R$)",
                "Fonte",
            ]
        ].copy()
        bi["Contratações correntes (R$ bi)"] = (
            bi["Contratações correntes (R$)"] / 1e9
        )
        bi["Contratações atualizadas IPCA 30/06/2026 (R$ bi)"] = (
            bi["Contratações atualizadas IPCA 30/06/2026 (R$)"] / 1e9
        )
        bi = bi[
            [
                "Ano",
                "Contratações correntes (R$ bi)",
                "Contratações atualizadas IPCA 30/06/2026 (R$ bi)",
                "Fonte",
            ]
        ]
        bi.to_excel(writer, sheet_name="Demonstrativo_R$_bi", index=False)

    lines = [
        "# Contratações do FNE (2000–2025)",
        "",
        "Valores totais de contratações do Fundo Constitucional de Financiamento "
        "do Nordeste, em valores correntes e atualizados pelo IPCA até 30/06/2026.",
        "",
        f"Arquivo: `{OUT_XLSX.relative_to(ROOT)}`",
        "",
        "| Ano | Correntes (R$ bi) | Atualizado IPCA 30/06/2026 (R$ bi) |",
        "|----:|------------------:|----------------------------------:|",
    ]
    for _, row in df.iterrows():
        lines.append(
            f"| {int(row['Ano'])} | "
            f"{row['Contratações correntes (R$)'] / 1e9:.3f} | "
            f"{row['Contratações atualizadas IPCA 30/06/2026 (R$)'] / 1e9:.3f} |"
        )
    lines += [
        "",
        f"**Soma correntes:** R$ {df['Contratações correntes (R$)'].sum() / 1e9:.1f} bi",
        f"**Soma atualizada (IPCA 30/06/2026):** "
        f"R$ {df['Contratações atualizadas IPCA 30/06/2026 (R$)'].sum() / 1e9:.1f} bi",
        "",
        "## Notas",
        "",
        "- Preferência pelo total do Relatório de Gestão do FNE; anos sem PDF de "
        "Gestão no portal ETENE usam RFNE/Resultados ou Demonstrações Financeiras do BNB.",
        "- Em 2017, o total do Gestão (R$ 15,97 bi) inclui Programação Padrão e "
        "Grande Porte de Infraestrutura; relatórios de Resultados frequentemente "
        "destacam só a Programação Padrão (R$ 12,3 bi).",
        "- Atualização monetária: IPCA médio do ano → índice de junho/2026 (BCB SGS 433).",
    ]
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {OUT_XLSX}")
    print(f"Wrote {OUT_MD}")


if __name__ == "__main__":
    write_outputs(build())
