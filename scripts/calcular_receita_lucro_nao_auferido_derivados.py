#!/usr/bin/env python3
"""Calcula receita líquida não obtida e lucro não auferido por derivado/mês.

Produtos: ÓLEO DIESEL, GASOLINA A, GLP, QUEROSENE DE AVIAÇÃO (2010–2026).

Fórmulas (por mês e derivado):

  receita_liquida_nao_obtida =
      (custo_medio_ponderado - 25) × volume_producao_Brasil

  lucro_nao_auferido =
      receita_liquida_nao_obtida + dispêndio_importacao

O volume usado é a produção nacional do derivado no mês (barris), pois o
referencial 25 é o custo unitário atribuído à produção doméstica.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from scripts.calcular_custo_medio_ponderado_derivados import (
    ANO_FIM,
    ANO_INICIO,
    CUSTO_PRODUCAO_NACIONAL,
    DEFAULT_IMP_XLSX,
    DEFAULT_OUT_DIR,
    DEFAULT_PROD_XLSX,
    MES_ORDEM,
    montar_base,
)
from scripts.extrair_volumes_importacao_derivados import PRODUTOS_FOCO


def calcular_receita_e_lucro_nao_auferido(df_custo: pd.DataFrame) -> pd.DataFrame:
    out = df_custo.copy()
    custo = out["custo_medio_ponderado_usd_por_barril"]
    vol = out["volume_producao_barris"]
    disp = out["dispendio_importacao_usd_fob"]

    # Só calcula quando há custo definido (há volume total > 0)
    receita = (custo - CUSTO_PRODUCAO_NACIONAL) * vol
    lucro = receita + disp

    out["receita_liquida_nao_obtida_usd"] = receita
    out["lucro_nao_auferido_usd"] = lucro
    out.loc[custo.isna(), ["receita_liquida_nao_obtida_usd", "lucro_nao_auferido_usd"]] = np.nan
    out["volume_derivado_usado_barris"] = vol
    out["volume_derivado_base"] = "producao_brasil"
    out["unidade_valor"] = "US$"
    return out


def gerar_saidas(df: pd.DataFrame, out_dir: Path) -> dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    cols = [
        "ano",
        "mes",
        "mes_nome",
        "mes_ordem",
        "produto",
        "volume_importacao_barris",
        "dispendio_importacao_usd_fob",
        "volume_producao_barris",
        "volume_derivado_usado_barris",
        "volume_derivado_base",
        "custo_medio_ponderado_usd_por_barril",
        "receita_liquida_nao_obtida_usd",
        "lucro_nao_auferido_usd",
        "unidade_valor",
    ]
    longo = df[cols]

    csv_longo = out_dir / "receita_lucro_nao_auferido_diesel_gasolina_glp_qav_2010_2026.csv"
    longo.to_csv(csv_longo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    resumo = (
        longo.groupby(["ano", "produto"], as_index=False)
        .agg(
            volume_importacao_barris=("volume_importacao_barris", "sum"),
            dispendio_importacao_usd_fob=("dispendio_importacao_usd_fob", "sum"),
            volume_producao_barris=("volume_producao_barris", "sum"),
            receita_liquida_nao_obtida_usd=("receita_liquida_nao_obtida_usd", "sum"),
            lucro_nao_auferido_usd=("lucro_nao_auferido_usd", "sum"),
        )
        .sort_values(["ano", "produto"])
    )
    csv_resumo = out_dir / "receita_lucro_nao_auferido_diesel_gasolina_glp_qav_resumo_anual.csv"
    resumo.to_csv(csv_resumo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    # Totais por ano (soma dos 4 derivados)
    totais_ano = (
        resumo.groupby("ano", as_index=False)[
            ["receita_liquida_nao_obtida_usd", "lucro_nao_auferido_usd"]
        ]
        .sum()
        .sort_values("ano")
    )
    csv_totais = out_dir / "receita_lucro_nao_auferido_totais_anuais.csv"
    totais_ano.to_csv(csv_totais, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    xlsx_path = out_dir / "receita_lucro_nao_auferido_diesel_gasolina_glp_qav_2010_2026.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        longo.to_excel(writer, sheet_name="Longo", index=False)
        resumo.to_excel(writer, sheet_name="Resumo anual", index=False)
        totais_ano.to_excel(writer, sheet_name="Totais anuais", index=False)

        for metric, sheet in [
            ("receita_liquida_nao_obtida_usd", "Matriz receita nao obtida"),
            ("lucro_nao_auferido_usd", "Matriz lucro nao auferido"),
        ]:
            matriz = longo.copy()
            matriz["ano_mes"] = matriz["ano"].astype(str) + "-" + matriz["mes"]
            wide = matriz.pivot_table(
                index="produto", columns="ano_mes", values=metric, aggfunc="sum"
            )
            ordered = sorted(
                wide.columns,
                key=lambda c: (int(c.split("-")[0]), MES_ORDEM.index(c.split("-")[1])),
            )
            wide = wide.reindex(columns=ordered).reindex(index=list(PRODUTOS_FOCO))
            wide.to_excel(writer, sheet_name=sheet)

        for produto in PRODUTOS_FOCO:
            sub = longo[longo["produto"] == produto]
            pivot = (
                sub.pivot_table(
                    index="mes",
                    columns="ano",
                    values="lucro_nao_auferido_usd",
                    aggfunc="sum",
                )
                .reindex(MES_ORDEM)
            )
            mes_nome = {
                "JAN": "Janeiro",
                "FEV": "Fevereiro",
                "MAR": "Março",
                "ABR": "Abril",
                "MAI": "Maio",
                "JUN": "Junho",
                "JUL": "Julho",
                "AGO": "Agosto",
                "SET": "Setembro",
                "OUT": "Outubro",
                "NOV": "Novembro",
                "DEZ": "Dezembro",
            }
            pivot.index = [mes_nome.get(m, m) for m in pivot.index]
            pivot.index.name = "Mês"
            name = f"Lucro {produto}"[:31]
            for ch in "[]:*?/\\":
                name = name.replace(ch, "-")
            pivot.reset_index().to_excel(writer, sheet_name=name, index=False)

    return {
        "csv_longo": csv_longo,
        "csv_resumo": csv_resumo,
        "csv_totais": csv_totais,
        "xlsx": xlsx_path,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--imp-source", type=Path, default=DEFAULT_IMP_XLSX)
    parser.add_argument("--prod-source", type=Path, default=DEFAULT_PROD_XLSX)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    parser.add_argument("--ano-inicio", type=int, default=ANO_INICIO)
    parser.add_argument("--ano-fim", type=int, default=ANO_FIM)
    args = parser.parse_args()

    for path in (args.imp_source, args.prod_source):
        if not path.exists():
            raise SystemExit(f"Arquivo fonte não encontrado: {path}")

    base = montar_base(args.imp_source, args.prod_source, args.ano_inicio, args.ano_fim)
    df = calcular_receita_e_lucro_nao_auferido(base)
    paths = gerar_saidas(df, args.out_dir)

    print(f"Linhas: {len(df)}")
    print(f"Produtos: {sorted(df['produto'].unique())}")
    print(f"Anos: {df['ano'].min()}–{df['ano'].max()}")
    print(
        f"Receita líquida não obtida (total): "
        f"{df['receita_liquida_nao_obtida_usd'].sum():,.2f} US$"
    )
    print(f"Lucro não auferido (total): {df['lucro_nao_auferido_usd'].sum():,.2f} US$")
    for label, path in paths.items():
        print(f"{label}: {path}")


if __name__ == "__main__":
    main()
