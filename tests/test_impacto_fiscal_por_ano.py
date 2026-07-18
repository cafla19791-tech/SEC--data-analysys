"""Testes do impacto fiscal agregado por ano de pagamento."""

from datetime import datetime

import numpy as np
import pandas as pd
import pytest

from scripts.gerar_fluxos import SelicSerie, calcular_impacto_fiscal_real
from scripts.impacto_fiscal_por_ano import (
    DATA_REFERENCIA,
    agregar_impacto_por_ano,
    calcular_meses_ate_2026,
)


def test_calcular_meses_ate_2026():
    assert calcular_meses_ate_2026(datetime(2026, 6, 30)) == 0
    assert calcular_meses_ate_2026(datetime(2025, 6, 30)) == 12
    assert calcular_meses_ate_2026(datetime(2026, 1, 15)) == 5
    assert calcular_meses_ate_2026(datetime(2009, 4, 15)) == (
        (2026 - 2009) * 12 + (6 - 4)
    )


def test_agregar_recalcular_formula_referencia():
    df = pd.DataFrame(
        {
            "data_fluxo": ["2009-04-15", "2010-04-15", "2010-05-15"],
            "subsidio": [100.0, 50.0, 50.0],
            "mes": [1, 2, 3],
        }
    )
    resumo = agregar_impacto_por_ano(df, modo="recalcular", taxa_selic_anual=0.145)

    assert list(resumo.columns) == [
        "Ano",
        "Soma Subsídio Nominal (R$)",
        "Impacto Fiscal 2026 (R$)",
        "Quantidade de Parcelas",
    ]
    assert list(resumo["Ano"]) == [2009, 2010]
    assert resumo.loc[resumo["Ano"] == 2009, "Soma Subsídio Nominal (R$)"].iloc[0] == 100.0
    assert resumo.loc[resumo["Ano"] == 2010, "Soma Subsídio Nominal (R$)"].iloc[0] == 100.0
    assert int(resumo["Quantidade de Parcelas"].sum()) == 3

    meses_2009 = calcular_meses_ate_2026(datetime(2009, 4, 15))
    esperado = round(100.0 * (1 + 0.145 / 12) ** meses_2009, 2)
    got = resumo.loc[resumo["Ano"] == 2009, "Impacto Fiscal 2026 (R$)"].iloc[0]
    assert got == pytest.approx(esperado, abs=0.01)


def test_agregar_modo_coluna():
    df = pd.DataFrame(
        {
            "data_fluxo": ["2009-04-15", "2010-04-15"],
            "subsidio": [100.0, 50.0],
            "impacto_fiscal": [1000.0, 400.0],
            "mes": [1, 2],
        }
    )
    resumo = agregar_impacto_por_ano(df, modo="coluna")
    assert resumo.loc[resumo["Ano"] == 2009, "Impacto Fiscal 2026 (R$)"].iloc[0] == 1000.0
    assert resumo.loc[resumo["Ano"] == 2010, "Impacto Fiscal 2026 (R$)"].iloc[0] == 400.0


def test_agregar_modo_coluna_alias_impacto():
    df = pd.DataFrame(
        {
            "data_fluxo": [DATA_REFERENCIA],
            "subsidio": [10.0],
            "impacto": [12.5],
        }
    )
    resumo = agregar_impacto_por_ano(df, modo="coluna")
    assert resumo["Impacto Fiscal 2026 (R$)"].iloc[0] == 12.5


def test_modo_coluna_sem_coluna_falha():
    df = pd.DataFrame({"data_fluxo": ["2009-04-15"], "subsidio": [1.0]})
    with pytest.raises(ValueError, match="impacto"):
        agregar_impacto_por_ano(df, modo="coluna")


def test_agregar_modo_contagil_mais_um_dia():
    """Paridade com calcular_impacto_fiscal_real (col E, capitaliza do dia seguinte)."""
    df = pd.DataFrame(
        {
            "data_fluxo": [datetime(2009, 2, 15), datetime(2010, 1, 15)],
            "subsidio": [100.0, 50.0],
            "mes": [1, 2],
        }
    )
    datas = np.array(
        [
            np.datetime64("2009-02-15"),
            np.datetime64("2009-02-16"),  # dia seguinte à parcela 15/02
            np.datetime64("2010-01-16"),
            np.datetime64("2026-06-30"),
        ],
        dtype="datetime64[ns]",
    )
    fatores = np.array([1.0, 2.0, 2.5, 4.0], dtype=float)
    serie = SelicSerie(datas, fatores, origem="test")

    resumo = agregar_impacto_por_ano(df, modo="contagil", selic_serie=serie)

    esperado_2009 = calcular_impacto_fiscal_real(
        100.0, datetime(2009, 2, 15), serie, DATA_REFERENCIA
    )
    esperado_2010 = calcular_impacto_fiscal_real(
        50.0, datetime(2010, 1, 15), serie, DATA_REFERENCIA
    )
    # 100 * 4/2 = 200 ; 50 * 4/2.5 = 80
    assert esperado_2009 == 200.0
    assert esperado_2010 == 80.0
    assert resumo.loc[resumo["Ano"] == 2009, "Impacto Fiscal 2026 (R$)"].iloc[0] == 200.0
    assert resumo.loc[resumo["Ano"] == 2010, "Impacto Fiscal 2026 (R$)"].iloc[0] == 80.0


def test_modo_contagil_exige_serie():
    df = pd.DataFrame({"data_fluxo": ["2009-04-15"], "subsidio": [1.0]})
    with pytest.raises(ValueError, match="série SELIC"):
        agregar_impacto_por_ano(df, modo="contagil", selic_serie=None)
