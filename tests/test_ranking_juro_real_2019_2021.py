"""Testes do ranking de juro real acumulado."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.ranking_juro_real_acumulado import (
    inflacao_acumulada,
    juro_nominal_acumulado,
    juro_real_fisher,
    markdown_ranking,
    montar_ranking,
    periodo,
)


def test_fisher_e_capitalizacao():
    meses, _, _ = periodo(2019, 2021)
    taxas = pd.Series(12.0, index=meses)
    i_nom = juro_nominal_acumulado(taxas, meses)
    assert abs(i_nom - ((1.12**3) - 1)) < 1e-12
    assert abs(juro_real_fisher(0.10, 0.10)) < 1e-12
    assert juro_real_fisher(0.14, 0.20) < 0


def test_inflacao_indice():
    _, cpi_ini, cpi_fim = periodo(2019, 2021)
    cpi = pd.Series({cpi_ini: 100.0, cpi_fim: 121.0})
    assert abs(inflacao_acumulada(cpi, cpi_ini, cpi_fim) - 0.21) < 1e-12


def _painel(
    area: str, taxa: float, cpi0: float, cpi1: float, ano_fim: int = 2021
) -> tuple[pd.DataFrame, pd.DataFrame]:
    meses = pd.date_range("2018-12-01", f"{ano_fim}-12-01", freq="MS")
    cb = pd.DataFrame(
        {
            "REF_AREA": area,
            "TIME_PERIOD": [d.strftime("%Y-%m") for d in meses],
            "OBS_VALUE": taxa,
        }
    )
    cpi = pd.DataFrame(
        {
            "REF_AREA": [area, area],
            "UNIT_MEASURE": [628, 628],
            "TIME_PERIOD": ["2018-12", f"{ano_fim}-12"],
            "OBS_VALUE": [cpi0, cpi1],
        }
    )
    return cb, cpi


def test_ranking_ordena_pelo_real():
    cb_a, cpi_a = _painel("CN", 10.0, 100.0, 105.0)
    cb_b, cpi_b = _painel("BR", 6.0, 100.0, 125.0)
    cb = pd.concat([cb_a, cb_b], ignore_index=True)
    cpi = pd.concat([cpi_a, cpi_b], ignore_index=True)
    rank = montar_ranking(cb, cpi, 2019, 2021)
    assert list(rank["codigo"]) == ["CN", "BR"]
    assert rank.iloc[0]["juro_real_acumulado_%"] > 0
    assert rank.iloc[1]["juro_real_acumulado_%"] < 0
    md = markdown_ranking(rank, 2019, 2021)
    assert "Brasil" in md
    assert "China" in md


def test_periodo_2022_tem_48_meses():
    meses, cpi_ini, cpi_fim = periodo(2019, 2022)
    assert len(meses) == 48
    assert cpi_ini == pd.Timestamp("2018-12-01")
    assert cpi_fim == pd.Timestamp("2022-12-01")
    cb, cpi = _painel("BR", 8.0, 100.0, 150.0, ano_fim=2022)
    rank = montar_ranking(cb, cpi, 2019, 2022)
    assert len(rank) == 1
    assert rank.iloc[0]["juro_real_acumulado_%"] < 0
