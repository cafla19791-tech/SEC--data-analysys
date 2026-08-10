#!/usr/bin/env python3
"""Contratações do FNO 2011–2026 a partir dos Relatórios da Administração (BASA RI).

Fonte dos PDFs: Central de Resultados do Banco da Amazônia
  https://ri.bancoamazonia.com.br/informacoes-financeiras/central-de-resultados/

Até 2020 o Relatório da Administração vem embutido no pacote de Demonstrações
Contábeis 4T (categoria MZIQ `central_de_resultados_demonstracoes_contabeis`).
A partir de 2021 há categoria própria
`central_de_resultados_relatorios_da_administracao`.
Para 2026 só há RA do 1T26 (ano incompleto).

Atualização monetária: IPCA BCB SGS 433, índice médio do ano da contratação
→ índice de jun/2026 (mesma convenção do demonstrativo FNE/FNO anterior).
"""

from __future__ import annotations

import json
import urllib.request
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT_XLSX = ROOT / "output" / "fno_contratacoes_2011_2026_basa_ra.xlsx"
OUT_MD = ROOT / "output" / "fno_contratacoes_2011_2026_basa_ra.md"
RI_URL = (
    "https://ri.bancoamazonia.com.br/informacoes-financeiras/central-de-resultados/"
)

# Valores em R$ (correntes do ano). Curadoria manual a partir dos textos extraídos.
# Preferência: total de contratações/aplicações FNO do próprio exercício no RA/DF 4T.
RECORDS: list[dict] = [
    {
        "Ano": 2011,
        "Valor_corrente_R$": 1_870.2e6,
        "Unidade_no_texto": "R$ 1.870,2 milhões",
        "Conceito": "contratado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T11",
        "Trecho": (
            "No exercício de 2011, foram contratados, com recursos do FNO, "
            "32.064 operações de crédito no valor total de R$1.870,2 milhões."
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "e00ed16c-ea32-4eeb-9323-1654b77cc2ce?origin=2"
        ),
        "Observacao": "",
    },
    {
        "Ano": 2012,
        "Valor_corrente_R$": 4_300.1e6,
        "Unidade_no_texto": "R$ 4.300,1 milhões",
        "Conceito": "contratado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T12",
        "Trecho": (
            "Com essa fonte de recursos, o Banco contratou, em 2012, o montante "
            "de R$4.300,1 milhões"
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "4605ff49-91b8-4bf7-bdf3-bd75ea2df812?origin=2"
        ),
        "Observacao": "",
    },
    {
        "Ano": 2013,
        "Valor_corrente_R$": 4_721.7e6,
        "Unidade_no_texto": "R$ 4.721,7 milhões",
        "Conceito": "aplicado/contratado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T13",
        "Trecho": (
            "total de recursos de fomento aplicado na Amazônia Legal... "
            "R$5.384,7 milhões, sendo R$4.721,7 milhões do FNO"
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "7c4677f8-dd31-47c6-8f20-ad98adc2a383?origin=2"
        ),
        "Observacao": "Confirmado no RA 2014 como volume aplicado em 2013.",
    },
    {
        "Ano": 2014,
        "Valor_corrente_R$": 5_366.5e6,
        "Unidade_no_texto": "R$ 5.366,5 milhões",
        "Conceito": "aplicado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T14",
        "Trecho": (
            "Fundo Constitucional de Financiamento do Norte (FNO), cujo volume "
            "atingiu R$5.366,5 milhões, ultrapassando em R$644,8 milhões o volume "
            "aplicado em 2013 (R$4.721,7 milhões)."
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "9a93c796-8ddf-4a3e-9fea-0d06a7c74a1d?origin=2"
        ),
        "Observacao": "",
    },
    {
        "Ano": 2015,
        "Valor_corrente_R$": 3_964.9e6,
        "Unidade_no_texto": "R$ 3.964,9 milhões",
        "Conceito": "contratado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T16 (tabela YoY)",
        "Trecho": (
            "Tabela FNO - Contratações por Porte/Setor: TOTAL Exercício 2015 = "
            "R$ 3.964,9 milhões (vs R$ 2.333,9 milhões em 2016)."
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "adc1e284-d13f-4016-a1f3-73f2dbe839a5?origin=2"
        ),
        "Observacao": (
            "O DF/RA 4T15 cita R$ 5.068,4 milhões de FNO, mas compara com "
            "volume liberado de 2014 — conceito liberado, não usado aqui."
        ),
    },
    {
        "Ano": 2016,
        "Valor_corrente_R$": 2_333.9e6,
        "Unidade_no_texto": "R$ 2.333,9 milhões",
        "Conceito": "contratado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T16",
        "Trecho": "Tabela FNO - Contratações: TOTAL Exercício 2016 = R$ 2.333,9 milhões.",
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "adc1e284-d13f-4016-a1f3-73f2dbe839a5?origin=2"
        ),
        "Observacao": "",
    },
    {
        "Ano": 2017,
        "Valor_corrente_R$": 2_905.9e6,
        "Unidade_no_texto": "R$ 2.905,9 milhões",
        "Conceito": "contratado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T17",
        "Trecho": "Tabela FNO - Contratações: TOTAL Exercício 2017 = R$ 2.905,9 milhões.",
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "b6038035-48b9-4d92-a23f-0dc134c450c9?origin=2"
        ),
        "Observacao": "Confirmado no RA 2018 como aplicados em 2017.",
    },
    {
        "Ano": 2018,
        "Valor_corrente_R$": 4_636.0e6,
        "Unidade_no_texto": "R$ 4.636,0 milhões",
        "Conceito": "contratado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T18",
        "Trecho": (
            "O volume de crédito de FNO contratado no ano de 2018 representou "
            "um valor acumulado em R$ 4.636,0 milhões"
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "a0fdf454-f32f-4633-855f-15bcd032f1f6?origin=2"
        ),
        "Observacao": (
            "RA 2019 menciona R$ 4.610,0 milhões aplicados em 2018 "
            "(pequena diferença vs contratado)."
        ),
    },
    {
        "Ano": 2019,
        "Valor_corrente_R$": 7_670.8e6,
        "Unidade_no_texto": "R$ 7.670,8 milhões",
        "Conceito": "contratado",
        "Cobertura": "ano completo",
        "Documento": "DF/RA 4T19",
        "Trecho": (
            "No exercício de 2019, as contratações do FNO totalizaram "
            "R$ 7.670,8 milhões"
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "99562963-ac0f-45e3-b934-25cc05422120?origin=2"
        ),
        "Observacao": "",
    },
    {
        "Ano": 2020,
        "Valor_corrente_R$": 10_500.0e6,
        "Unidade_no_texto": "R$ 10,5 bilhões",
        "Conceito": "contratado",
        "Cobertura": "ano completo",
        "Documento": "RA 4T21 (comparativo YoY)",
        "Trecho": (
            "aplicados, R$ 12,5 bilhões, o que representa aumento de 19,2%, "
            "em comparação a 2020 que foi R$ 10,5 bilhões"
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "ec2a1b9a-cade-4c93-a8e2-b9ef795c2b5d?origin=2"
        ),
        "Observacao": (
            "Pacote DF 4T20 no RI não traz narrativa de contratações FNO "
            "(arquivo Word/notas). Valor arredondado no RA 2021; série MDR/"
            "SUDAM publica R$ 10.486,0 milhões."
        ),
    },
    {
        "Ano": 2021,
        "Valor_corrente_R$": 12_500.0e6,
        "Unidade_no_texto": "R$ 12,5 bilhões",
        "Conceito": "contratado/aplicado",
        "Cobertura": "ano completo",
        "Documento": "RA 4T21",
        "Trecho": (
            "Melhor resultado nos últimos anos nas contratações do FNO "
            "R$ 12,5 bilhões."
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "ec2a1b9a-cade-4c93-a8e2-b9ef795c2b5d?origin=2"
        ),
        "Observacao": (
            "RA 2022 cita R$ 13,3 bilhões aplicados em 2021 — divergência "
            "em relação ao RA do próprio exercício (12,5). Mantido o valor "
            "do RA 2021. MDR/SUDAM: R$ 12.497,8 milhões."
        ),
    },
    {
        "Ano": 2022,
        "Valor_corrente_R$": 12_000.0e6,
        "Unidade_no_texto": "R$ 12,0 bilhões",
        "Conceito": "contratado/aplicado",
        "Cobertura": "ano completo",
        "Documento": "RA 4T22",
        "Trecho": (
            "foram contratados no exercício de 2022... no valor de R$ 12,0 bilhões"
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "623f7680-7ca5-4126-9bdf-90725a24f147?origin=2"
        ),
        "Observacao": "",
    },
    {
        "Ano": 2023,
        "Valor_corrente_R$": 11_300.0e6,
        "Unidade_no_texto": "R$ 11,3 bilhões",
        "Conceito": "aplicado/disponibilizado",
        "Cobertura": "ano completo",
        "Documento": "RA 4T23",
        "Trecho": (
            "No ano de 2023, disponibilizamos R$ 11,3 bilhões em financiamentos "
            "com recursos do FNO."
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "0b153c03-f5f0-40ae-b672-7e8a5df45899?origin=2"
        ),
        "Observacao": "RA 2024 confirma: em 2023 foi contratado R$ 11,3 bilhões.",
    },
    {
        "Ano": 2024,
        "Valor_corrente_R$": 13_600.0e6,
        "Unidade_no_texto": "R$ 13,6 bilhões",
        "Conceito": "aplicado/contratado",
        "Cobertura": "ano completo",
        "Documento": "RA 4T24",
        "Trecho": (
            "Em 2024, aplicamos R$ 13,6 bilhões em financiamentos com recursos "
            "do FNO... em relação a 2023, quando foi contratado R$ 11,3 bilhões."
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "e34d4222-6ba1-d0f9-8032-f198c2a0021e?origin=2"
        ),
        "Observacao": "",
    },
    {
        "Ano": 2025,
        "Valor_corrente_R$": 17_800.0e6,
        "Unidade_no_texto": "R$ 17,8 bilhões",
        "Conceito": "aplicado/contratado",
        "Cobertura": "ano completo",
        "Documento": "RA 4T25",
        "Trecho": (
            "Aplicamos R$ 17,8 bilhões em financiamentos com recursos do FNO... "
            "em relação ao realizado em 2024, quando foram contratados "
            "R$ 13,6 bilhões."
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "57ed37d4-9660-acac-c823-c7118c36f6b9?origin=2"
        ),
        "Observacao": "",
    },
    {
        "Ano": 2026,
        "Valor_corrente_R$": 2_600.0e6,
        "Unidade_no_texto": "R$ 2,6 bilhões",
        "Conceito": "contratado (FNO no fomento)",
        "Cobertura": "1T26 (parcial)",
        "Documento": "RA 1T26",
        "Trecho": (
            "Gráfico Fomento Contratado: FNO 1T26 = R$ 2,6 bi "
            "(fomento total contratado R$ 3,2 bi)."
        ),
        "URL": (
            "https://api.mziq.com/mzfilemanager/v2/d/"
            "348fb9a6-7a7d-45a4-8beb-a6941accbe1d/"
            "d8c71701-a8ca-2e67-ef0c-2aa3769e3cc7?origin=2"
        ),
        "Observacao": (
            "Ano incompleto — apenas 1º trimestre de 2026 disponível na "
            "Central de Resultados. Não somar com anos cheios sem ressalva."
        ),
    },
]


def fetch_json(url: str, path: Path) -> list:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        return json.loads(path.read_text())
    with urllib.request.urlopen(url, timeout=90) as resp:
        data = json.loads(resp.read().decode())
    path.write_text(json.dumps(data), encoding="utf-8")
    return data


def load_ipca() -> tuple[pd.Series, float, str]:
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
    ipca_ann, ipca_jun, meta = load_ipca()
    base = pd.DataFrame(RECORDS)
    years = base["Ano"].astype(int)
    missing = [y for y in years if y not in ipca_ann.index]
    if missing:
        raise RuntimeError(f"IPCA anual ausente para: {missing}")

    fator = ipca_jun / ipca_ann.loc[years].values
    base["Fator_IPCA_media_ano_para_jun2026"] = fator
    base["Atualizado_IPCA_30_06_2026_R$"] = base["Valor_corrente_R$"] * fator
    base["Fonte_portal"] = RI_URL
    base["Atualizacao"] = f"IPCA BCB SGS 433 até {meta}"

    demo = base[
        [
            "Ano",
            "Valor_corrente_R$",
            "Unidade_no_texto",
            "Conceito",
            "Cobertura",
            "Fator_IPCA_media_ano_para_jun2026",
            "Atualizado_IPCA_30_06_2026_R$",
            "Documento",
            "Trecho",
            "URL",
            "Observacao",
            "Atualizacao",
        ]
    ].copy()

    bi = pd.DataFrame(
        {
            "Ano": demo["Ano"],
            "Contratações correntes (R$ bi)": demo["Valor_corrente_R$"] / 1e9,
            "Atualizado IPCA 30/06/2026 (R$ bi)": demo["Atualizado_IPCA_30_06_2026_R$"]
            / 1e9,
            "Cobertura": demo["Cobertura"],
            "Conceito": demo["Conceito"],
            "Documento": demo["Documento"],
        }
    )

    full = demo[demo["Cobertura"] == "ano completo"]
    metodologia = pd.DataFrame(
        [
            {"Item": "Portal RI", "Valor": RI_URL},
            {
                "Item": "Categorias MZIQ",
                "Valor": (
                    "2011-2020: central_de_resultados_demonstracoes_contabeis (4T); "
                    "2021-2025: central_de_resultados_relatorios_da_administracao (4T); "
                    "2026: RA 1T26"
                ),
            },
            {
                "Item": "Conceito preferido",
                "Valor": (
                    "Contratações/aplicações FNO do exercício, conforme texto "
                    "ou tabela do RA. Quando o RA do ano fala em liberado, "
                    "usa-se a tabela de contratações do RA seguinte."
                ),
            },
            {
                "Item": "Atualização",
                "Valor": (
                    f"IPCA BCB SGS 433; fator = índice jun/2026 / média anual "
                    f"do índice no ano da contratação (meta {meta})."
                ),
            },
            {
                "Item": "Soma correntes 2011-2025 (anos cheios)",
                "Valor": float(full["Valor_corrente_R$"].sum()),
            },
            {
                "Item": "Soma IPCA 30/06/2026 2011-2025",
                "Valor": float(full["Atualizado_IPCA_30_06_2026_R$"].sum()),
            },
            {
                "Item": "2026",
                "Valor": "Parcial (1T26): FNO ≈ R$ 2,6 bi; fomento total R$ 3,2 bi.",
            },
            {
                "Item": "Série complementar",
                "Valor": (
                    "output/fno_contratacoes_2000_2021_ipca.xlsx — digitalização "
                    "da Figura 3 do livro BASA (2000-2021)."
                ),
            },
        ]
    )

    fontes = demo[["Ano", "Documento", "URL", "Trecho", "Observacao"]].copy()
    return {"demo": demo, "bi": bi, "metodologia": metodologia, "fontes": fontes}


def write_outputs(tables: dict[str, pd.DataFrame]) -> None:
    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    demo, bi = tables["demo"], tables["bi"]

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        demo.to_excel(writer, sheet_name="Serie_BASA_RA", index=False)
        bi.to_excel(writer, sheet_name="Serie_R$_bi", index=False)
        tables["fontes"].to_excel(writer, sheet_name="Fontes_e_trechos", index=False)
        tables["metodologia"].to_excel(writer, sheet_name="Metodologia", index=False)

    full = bi[bi["Cobertura"] == "ano completo"]
    lines = [
        "# Contratações do FNO (2011–2026) — Relatórios da Administração BASA",
        "",
        f"Portal: {RI_URL}",
        "",
        f"Arquivo: `{OUT_XLSX.relative_to(ROOT)}`",
        "",
        "## Série (valores correntes e IPCA 30/06/2026)",
        "",
        "| Ano | Corrente (R$ bi) | IPCA 30/06/2026 (R$ bi) | Cobertura | Conceito | Documento |",
        "|----:|-----------------:|------------------------:|:----------|:---------|:----------|",
    ]
    for _, r in bi.iterrows():
        lines.append(
            f"| {int(r['Ano'])} | {r['Contratações correntes (R$ bi)']:.3f} | "
            f"{r['Atualizado IPCA 30/06/2026 (R$ bi)']:.3f} | {r['Cobertura']} | "
            f"{r['Conceito']} | {r['Documento']} |"
        )
    lines += [
        "",
        f"Soma 2011–2025 (anos cheios), correntes: "
        f"R$ {full['Contratações correntes (R$ bi)'].sum():.2f} bi",
        f"Soma 2011–2025 (anos cheios), IPCA 30/06/2026: "
        f"R$ {full['Atualizado IPCA 30/06/2026 (R$ bi)'].sum():.2f} bi",
        "",
        "## Notas",
        "",
        "- 2015: usa TOTAL de **contratações** da tabela do RA 2016 "
        "(R$ 3.964,9 mi); o RA 2015 cita R$ 5.068,4 mi em linguagem de "
        "**liberado**.",
        "- 2020: o pacote DF 4T20 do RI não traz o total FNO; valor R$ 10,5 bi "
        "vem do comparativo do RA 2021 (MDR/SUDAM: R$ 10.486,0 mi).",
        "- 2021: RA do exercício = R$ 12,5 bi; RA 2022 cita R$ 13,3 bi "
        "aplicados em 2021 (divergência documentada).",
        "- 2026: apenas 1T26 — FNO ≈ R$ 2,6 bi no gráfico de fomento contratado.",
        "",
        "Regenerar:",
        "",
        "```bash",
        "python3 scripts/build_fno_basa_ra.py",
        "```",
    ]
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {OUT_XLSX}")
    print(f"Wrote {OUT_MD}")


if __name__ == "__main__":
    write_outputs(build())
