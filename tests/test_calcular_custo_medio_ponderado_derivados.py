from __future__ import annotations

import math
from pathlib import Path

import pandas as pd

from scripts.calcular_custo_medio_ponderado_derivados import (
    calcular_custo_medio_ponderado,
    gerar_saidas,
)


def test_formula_ponderada_basica() -> None:
    # 100 barris importados a 100 USD/b + 300 produzidos a 25 →
    # w_imp=0.25, w_prod=0.75 → custo = 100*0.25 + 25*0.75 = 43.75
    r = calcular_custo_medio_ponderado(
        volume_importacao=100,
        dispendio_usd_fob=10_000,
        volume_producao=300,
    )
    assert math.isclose(r["preco_importacao_usd_por_barril"], 100.0)
    assert math.isclose(r["peso_importacao"], 0.25)
    assert math.isclose(r["peso_producao"], 0.75)
    assert math.isclose(r["custo_medio_ponderado_usd_por_barril"], 43.75)


def test_somente_producao() -> None:
    r = calcular_custo_medio_ponderado(0, 0, 500)
    assert math.isclose(r["custo_medio_ponderado_usd_por_barril"], 25.0)
    assert math.isnan(r["preco_importacao_usd_por_barril"])


def test_somente_importacao() -> None:
    r = calcular_custo_medio_ponderado(50, 5000, 0)
    assert math.isclose(r["custo_medio_ponderado_usd_por_barril"], 100.0)
    assert math.isclose(r["peso_importacao"], 1.0)


def test_ambos_zero() -> None:
    r = calcular_custo_medio_ponderado(0, 0, 0)
    assert math.isnan(r["custo_medio_ponderado_usd_por_barril"])


def test_gerar_saidas_synthetic(tmp_path: Path) -> None:
    rows = []
    for mes_i, mes in enumerate(["JAN", "FEV"], start=1):
        for produto in ["ÓLEO DIESEL", "GASOLINA A", "GLP", "QUEROSENE DE AVIAÇÃO"]:
            calc = calcular_custo_medio_ponderado(100, 8000, 300)
            rows.append(
                {
                    "ano": 2020,
                    "mes": mes,
                    "mes_nome": "Janeiro" if mes == "JAN" else "Fevereiro",
                    "mes_ordem": mes_i,
                    "produto": produto,
                    "volume_importacao_barris": 100.0,
                    "dispendio_importacao_usd_fob": 8000.0,
                    "volume_producao_barris": 300.0,
                    **calc,
                    "custo_producao_nacional_usd_por_barril": 25.0,
                    "unidade_custo": "US$/barril",
                }
            )
    df = pd.DataFrame(rows)
    paths = gerar_saidas(df, tmp_path)
    assert paths["csv_longo"].exists()
    longo = pd.read_csv(paths["csv_longo"], sep=";", decimal=",", encoding="utf-8-sig")
    assert len(longo) == 8
    assert math.isclose(float(longo["custo_medio_ponderado_usd_por_barril"].iloc[0]), 38.75)
