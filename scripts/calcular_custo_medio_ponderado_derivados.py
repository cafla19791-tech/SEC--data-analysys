#!/usr/bin/env python3
"""Calcula o custo médio ponderado mensal (US$/barril) por derivado.

Produtos: ÓLEO DIESEL, GASOLINA A, GLP, QUEROSENE DE AVIAÇÃO.
Período: cada mês de 2010 a 2026.

Fórmula (custo médio ponderado no mês):

  P_imp = dispêndio_importação / volume_importação
  w_imp = volume_importação / (volume_importação + volume_produção_Brasil)
  w_prod = volume_produção_Brasil / (volume_importação + volume_produção_Brasil)

  custo = P_imp * w_imp + 25 * w_prod

onde 25 é o custo unitário atribuído à produção nacional (US$/barril).

Casos-limite:
  - sem importação e com produção → custo = 25
  - sem produção e com importação → custo = P_imp
  - ambos zero → custo indefinido (vazio)
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from scripts.extrair_volumes_importacao_derivados import (
    PRODUTOS_FOCO,
    carregar_volumes_e_dispendios,
)
from scripts.extrair_volumes_producao_derivados import carregar_producao_brasil

DEFAULT_IMP_XLSX = Path("data/anp/anp_importacoes_exportacoes_barris.xlsx")
DEFAULT_PROD_XLSX = Path("data/anp/anp_producao_nacional_derivados_barris.xlsx")
DEFAULT_OUT_DIR = Path("output/anp")

CUSTO_PRODUCAO_NACIONAL = 25.0  # US$/barril
ANO_INICIO = 2010
ANO_FIM = 2026
MES_ORDEM = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]


def calcular_custo_medio_ponderado(
    volume_importacao: float,
    dispendio_usd_fob: float,
    volume_producao: float,
    custo_producao: float = CUSTO_PRODUCAO_NACIONAL,
) -> dict[str, float]:
    vi = float(volume_importacao or 0.0)
    vp = float(volume_producao or 0.0)
    disp = float(dispendio_usd_fob or 0.0)
    total = vi + vp

    if total <= 0:
        return {
            "preco_importacao_usd_por_barril": np.nan,
            "peso_importacao": np.nan,
            "peso_producao": np.nan,
            "custo_medio_ponderado_usd_por_barril": np.nan,
        }

    w_imp = vi / total
    w_prod = vp / total

    if vi > 0:
        p_imp = disp / vi
        custo = p_imp * w_imp + custo_producao * w_prod
    else:
        p_imp = np.nan
        custo = custo_producao  # só produção nacional

    return {
        "preco_importacao_usd_por_barril": p_imp,
        "peso_importacao": w_imp,
        "peso_producao": w_prod,
        "custo_medio_ponderado_usd_por_barril": custo,
    }


def montar_base(
    imp_xlsx: Path,
    prod_xlsx: Path,
    ano_inicio: int = ANO_INICIO,
    ano_fim: int = ANO_FIM,
) -> pd.DataFrame:
    imp = carregar_volumes_e_dispendios(
        imp_xlsx, produtos=PRODUTOS_FOCO, ano_inicio=ano_inicio, ano_fim=ano_fim
    ).rename(
        columns={
            "volume_barris": "volume_importacao_barris",
            "dispendio_usd_fob": "dispendio_importacao_usd_fob",
        }
    )
    prod = carregar_producao_brasil(
        prod_xlsx, produtos=PRODUTOS_FOCO, ano_inicio=ano_inicio, ano_fim=ano_fim
    ).rename(columns={"volume_barris": "volume_producao_barris"})

    keys = ["ano", "mes", "mes_nome", "mes_ordem", "produto"]
    base = imp[keys + ["volume_importacao_barris", "dispendio_importacao_usd_fob"]].merge(
        prod[keys + ["volume_producao_barris"]],
        on=keys,
        how="outer",
    )
    for col in [
        "volume_importacao_barris",
        "dispendio_importacao_usd_fob",
        "volume_producao_barris",
    ]:
        base[col] = base[col].fillna(0.0)

    calc = base.apply(
        lambda r: calcular_custo_medio_ponderado(
            r["volume_importacao_barris"],
            r["dispendio_importacao_usd_fob"],
            r["volume_producao_barris"],
        ),
        axis=1,
        result_type="expand",
    )
    out = pd.concat([base, calc], axis=1)
    out["custo_producao_nacional_usd_por_barril"] = CUSTO_PRODUCAO_NACIONAL
    out["unidade_custo"] = "US$/barril"
    return out.sort_values(["produto", "ano", "mes_ordem"]).reset_index(drop=True)


def pivot_custo(df: pd.DataFrame) -> pd.DataFrame:
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
    pivot = (
        df.pivot_table(
            index="mes",
            columns="ano",
            values="custo_medio_ponderado_usd_por_barril",
            aggfunc="mean",
        )
        .reindex(MES_ORDEM)
    )
    pivot.index = [mes_nome.get(m, m) for m in pivot.index]
    pivot.index.name = "Mês"
    return pivot.reset_index()


def sheet_name(produto: str) -> str:
    name = f"Custo {produto}"[:31]
    for ch in "[]:*?/\\":
        name = name.replace(ch, "-")
    return name


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
        "preco_importacao_usd_por_barril",
        "peso_importacao",
        "peso_producao",
        "custo_producao_nacional_usd_por_barril",
        "custo_medio_ponderado_usd_por_barril",
        "unidade_custo",
    ]
    longo = df[cols]

    csv_longo = out_dir / "custo_medio_ponderado_diesel_gasolina_glp_qav_2010_2026.csv"
    longo.to_csv(csv_longo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    resumo = (
        longo.groupby(["ano", "produto"], as_index=False)
        .agg(
            volume_importacao_barris=("volume_importacao_barris", "sum"),
            dispendio_importacao_usd_fob=("dispendio_importacao_usd_fob", "sum"),
            volume_producao_barris=("volume_producao_barris", "sum"),
            custo_medio_ponderado_media_simples=(
                "custo_medio_ponderado_usd_por_barril",
                "mean",
            ),
        )
        .sort_values(["ano", "produto"])
    )
    # média anual ponderada pelo volume total (imp+prod) do mês
    rows = []
    for (ano, produto), g in longo.groupby(["ano", "produto"]):
        pesos = g["volume_importacao_barris"] + g["volume_producao_barris"]
        custos = g["custo_medio_ponderado_usd_por_barril"]
        mask = (pesos > 0) & custos.notna()
        if mask.any() and float(pesos[mask].sum()) > 0:
            media = float(np.average(custos[mask], weights=pesos[mask]))
        else:
            media = np.nan
        rows.append({"ano": ano, "produto": produto, "custo_medio_ponderado_anual": media})
    resumo = resumo.merge(pd.DataFrame(rows), on=["ano", "produto"], how="left")

    csv_resumo = out_dir / "custo_medio_ponderado_diesel_gasolina_glp_qav_resumo_anual.csv"
    resumo.to_csv(csv_resumo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    xlsx_path = out_dir / "custo_medio_ponderado_diesel_gasolina_glp_qav_2010_2026.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        longo.to_excel(writer, sheet_name="Longo", index=False)
        resumo.to_excel(writer, sheet_name="Resumo anual", index=False)

        matriz = longo.copy()
        matriz["ano_mes"] = matriz["ano"].astype(str) + "-" + matriz["mes"]
        wide = matriz.pivot_table(
            index="produto",
            columns="ano_mes",
            values="custo_medio_ponderado_usd_por_barril",
            aggfunc="mean",
        )
        ordered = sorted(
            wide.columns,
            key=lambda c: (int(c.split("-")[0]), MES_ORDEM.index(c.split("-")[1])),
        )
        wide = wide.reindex(columns=ordered).reindex(index=list(PRODUTOS_FOCO))
        wide.to_excel(writer, sheet_name="Matriz custo USD-barril")

        for produto in PRODUTOS_FOCO:
            sub = longo[longo["produto"] == produto]
            pivot_custo(sub).to_excel(writer, sheet_name=sheet_name(produto), index=False)

    return {"csv_longo": csv_longo, "csv_resumo": csv_resumo, "xlsx": xlsx_path}


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

    df = montar_base(args.imp_source, args.prod_source, args.ano_inicio, args.ano_fim)
    paths = gerar_saidas(df, args.out_dir)

    print(f"Linhas: {len(df)}")
    print(f"Produtos: {sorted(df['produto'].unique())}")
    print(f"Anos: {df['ano'].min()}–{df['ano'].max()}")
    valid = df["custo_medio_ponderado_usd_por_barril"].dropna()
    print(f"Custo médio (média simples da série): {valid.mean():,.2f} US$/barril")
    for label, path in paths.items():
        print(f"{label}: {path}")


if __name__ == "__main__":
    main()
