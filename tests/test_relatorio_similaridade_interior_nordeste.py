"""Testes das métricas de similaridade no interior do Nordeste."""

from __future__ import annotations

import numpy as np
import pandas as pd

from scripts.relatorio_similaridade_interior_nordeste import (
    cohen_kappa,
    faixa_aptos,
    icc_consistency,
    pearson_ponderado,
    stats_par,
)


def test_kappa_perfeito_e_discordancia():
    iguais = pd.Series(["PT", "PT", "OPP", "OPP"])
    assert cohen_kappa(iguais, iguais) == 1.0
    assert cohen_kappa(iguais, pd.Series(["OPP", "OPP", "PT", "PT"])) < 0


def test_icc_serie_constante_e_identica():
    igual = np.array([[70.0, 70.0, 70.0], [80.0, 80.0, 80.0], [60.0, 60.0, 60.0]])
    assert icc_consistency(igual) == 1.0


def test_pearson_ponderado_igual_ao_simples_com_peso_uniforme():
    x = pd.Series([10.0, 20.0, 30.0, 40.0])
    y = pd.Series([11.0, 19.0, 31.0, 39.0])
    w = pd.Series([1.0, 1.0, 1.0, 1.0])
    assert abs(pearson_ponderado(x, y, w) - float(x.corr(y))) < 1e-9


def test_stats_par_em_serie_quase_igual():
    df = pd.DataFrame(
        {
            "PCT_PT_VALIDOS_2018": [70.0, 80.0, 90.0, 60.0],
            "PCT_PT_VALIDOS_2022": [71.0, 79.0, 91.0, 61.0],
            "LADO_2018": ["PT", "PT", "PT", "PT"],
            "LADO_2022": ["PT", "PT", "PT", "PT"],
            "QT_VOTOS_VALIDOS_2022": [100, 100, 100, 100],
        }
    )
    s = stats_par(df, 2018, 2022)
    assert s["N"] == 4
    assert s["Pearson (r)"] > 0.99
    assert s["Persistência do vencedor (%)"] == 100.0
    assert s["Inversões"] == 0
    assert s["MAE (p.p.)"] == 1.0


def test_faixa_aptos():
    assert faixa_aptos(8000) == "Até 10 mil"
    assert faixa_aptos(15000) == "10 a 20 mil"
    assert faixa_aptos(25000) == "20 a 40 mil"
    assert faixa_aptos(80000) == "40 a 100 mil"
    assert faixa_aptos(150000) == "Mais de 100 mil"
