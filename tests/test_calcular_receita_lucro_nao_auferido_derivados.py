from __future__ import annotations

import math
from pathlib import Path

import pandas as pd

from scripts.calcular_receita_lucro_nao_auferido_derivados import (
    calcular_receita_e_lucro_nao_auferido,
    gerar_saidas,
)


def test_formula_receita_e_lucro() -> None:
    # custo=43.75, vol_prod=300, disp=10000
    # receita=(43.75-25)*300=5625; lucro=5625+10000=15625
    df = pd.DataFrame(
        [
            {
                "ano": 2020,
                "mes": "JAN",
                "mes_nome": "Janeiro",
                "mes_ordem": 1,
                "produto": "ÓLEO DIESEL",
                "volume_importacao_barris": 100.0,
                "dispendio_importacao_usd_fob": 10_000.0,
                "volume_producao_barris": 300.0,
                "custo_medio_ponderado_usd_por_barril": 43.75,
            }
        ]
    )
    out = calcular_receita_e_lucro_nao_auferido(df)
    assert math.isclose(float(out["receita_liquida_nao_obtida_usd"].iloc[0]), 5625.0)
    assert math.isclose(float(out["lucro_nao_auferido_usd"].iloc[0]), 15625.0)


def test_custo_indefinido_gera_nan() -> None:
    df = pd.DataFrame(
        [
            {
                "ano": 2020,
                "mes": "JAN",
                "mes_nome": "Janeiro",
                "mes_ordem": 1,
                "produto": "GLP",
                "volume_importacao_barris": 0.0,
                "dispendio_importacao_usd_fob": 0.0,
                "volume_producao_barris": 0.0,
                "custo_medio_ponderado_usd_por_barril": float("nan"),
            }
        ]
    )
    out = calcular_receita_e_lucro_nao_auferido(df)
    assert math.isnan(float(out["receita_liquida_nao_obtida_usd"].iloc[0]))
    assert math.isnan(float(out["lucro_nao_auferido_usd"].iloc[0]))


def test_gerar_saidas(tmp_path: Path) -> None:
    rows = []
    for produto in ["ÓLEO DIESEL", "GASOLINA A", "GLP", "QUEROSENE DE AVIAÇÃO"]:
        rows.append(
            {
                "ano": 2020,
                "mes": "JAN",
                "mes_nome": "Janeiro",
                "mes_ordem": 1,
                "produto": produto,
                "volume_importacao_barris": 100.0,
                "dispendio_importacao_usd_fob": 10_000.0,
                "volume_producao_barris": 300.0,
                "custo_medio_ponderado_usd_por_barril": 43.75,
            }
        )
    df = calcular_receita_e_lucro_nao_auferido(pd.DataFrame(rows))
    paths = gerar_saidas(df, tmp_path)
    assert paths["xlsx"].exists()
    longo = pd.read_csv(paths["csv_longo"], sep=";", decimal=",", encoding="utf-8-sig")
    assert len(longo) == 4
    assert math.isclose(float(longo["lucro_nao_auferido_usd"].sum()), 15625.0 * 4)
