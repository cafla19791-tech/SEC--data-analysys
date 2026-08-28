"""Testes do ranking de juro real acumulado 2019–2021."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.ranking_juro_real_2019_2021 import (
    inflacao_acumulada,
    juro_nominal_acumulado,
    juro_real_fisher,
    markdown_ranking,
    montar_ranking,
)


def test_fisher_e_capitalizacao():
    meses = pd.date_range("2019-01-01", "2021-12-01", freq="MS")
    taxas = pd.Series(12.0, index=meses)  # 12% a.a. o período inteiro
    i_nom = juro_nominal_acumulado(taxas)
    assert abs(i_nom - ((1.12**3) - 1)) < 1e-12
    assert abs(juro_real_fisher(0.10, 0.10)) < 1e-12
    assert juro_real_fisher(0.14, 0.20) < 0


def test_inflacao_indice():
    cpi = pd.Series(
        {pd.Timestamp("2018-12-01"): 100.0, pd.Timestamp("2021-12-01"): 121.0}
    )
    assert abs(inflacao_acumulada(cpi) - 0.21) < 1e-12


def _painel(area: str, taxa: float, cpi0: float, cpi1: float) -> tuple[pd.DataFrame, pd.DataFrame]:
    meses = pd.date_range("2018-12-01", "2021-12-01", freq="MS")
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
            "TIME_PERIOD": ["2018-12", "2021-12"],
            "OBS_VALUE": [cpi0, cpi1],
        }
    )
    return cb, cpi


def test_ranking_ordena_pelo_real(tmp_path: Path):
    cb_a, cpi_a = _painel("AA", 10.0, 100.0, 105.0)  # real positivo
    cb_b, cpi_b = _painel("BR", 6.0, 100.0, 125.0)  # real negativo
    # AA não está em NOMES — usa o código
    cb = pd.concat([cb_a, cb_b], ignore_index=True)
    cpi = pd.concat([cpi_a, cpi_b], ignore_index=True)
    # troca AA por CN (está no catálogo)
    cb.loc[cb["REF_AREA"] == "AA", "REF_AREA"] = "CN"
    cpi.loc[cpi["REF_AREA"] == "AA", "REF_AREA"] = "CN"
    rank = montar_ranking(cb, cpi)
    assert list(rank["codigo"]) == ["CN", "BR"]
    assert rank.iloc[0]["juro_real_acumulado_%"] > 0
    assert rank.iloc[1]["juro_real_acumulado_%"] < 0
    md = markdown_ranking(rank)
    assert "Brasil" in md
    assert "China" in md
