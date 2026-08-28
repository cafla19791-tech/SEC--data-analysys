"""Testes da relação de índices oficiais de inflação EUA/Brasil."""

from __future__ import annotations

import pandas as pd

from scripts.indices_inflacao_eua_brasil import (
    formatar_pct,
    markdown_tabela,
    montar_tabela,
    serie_anual,
    variacao_ano_indice,
    variacao_ano_percentual_12m,
    variacao_ano_percentual_mensal,
)


def _mensal_pct(ano: int, taxas: list[float]) -> pd.DataFrame:
    datas = pd.date_range(f"{ano}-01-01", f"{ano}-12-01", freq="MS")
    return pd.DataFrame({"data": datas, "valor": taxas})


def _indice(anos_dez: dict[int, float]) -> pd.DataFrame:
    rows = [
        {"data": pd.Timestamp(year=ano, month=12, day=1), "valor": valor}
        for ano, valor in anos_dez.items()
    ]
    return pd.DataFrame(rows)


def test_produto_mensal_doze_meses():
    df = _mensal_pct(2000, [1.0] * 12)
    got = variacao_ano_percentual_mensal(df, 2000)
    assert got is not None
    assert abs(got - ((1.01**12) - 1.0)) < 1e-12


def test_ano_incompleto_nao_entra():
    df = _mensal_pct(2000, [1.0] * 12).iloc[:11]
    assert variacao_ano_percentual_mensal(df, 2000) is None


def test_dezembro_sobre_dezembro():
    df = _indice({1999: 100.0, 2000: 103.0})
    assert abs(variacao_ano_indice(df, 2000) - 0.03) < 1e-12
    assert variacao_ano_indice(df, 1999) is None


def test_doze_meses_oficial_usa_dezembro():
    df = pd.DataFrame(
        {
            "data": pd.to_datetime(["2023-11-01", "2023-12-01", "2024-12-01"]),
            "valor": [10.00, 4.62, 4.83],
        }
    )
    assert abs(variacao_ano_percentual_12m(df, 2023) - 0.0462) < 1e-12
    assert abs(variacao_ano_percentual_12m(df, 2024) - 0.0483) < 1e-12
    assert variacao_ano_percentual_12m(df, 2022) is None


def test_montar_tabela_alinha_paises():
    series = {
        "ipca": pd.DataFrame(
            {"data": [pd.Timestamp("2024-12-01")], "valor": [6.17]}
        ),
        "inpc": _mensal_pct(2024, [0.4] * 12),
        "igp_m": _mensal_pct(2024, [0.2] * 12),
        "igp_di": _mensal_pct(2024, [0.1] * 12),
        "cpi_u": _indice({2023: 200.0, 2024: 206.0}),
        "pce": _indice({2023: 100.0, 2024: 102.5}),
    }
    tab = montar_tabela(series, [2024])
    assert len(tab) == 1
    assert abs(tab.iloc[0]["ipca"] - 0.0617) < 1e-12
    assert abs(tab.iloc[0]["cpi_u"] - 0.03) < 1e-12
    assert abs(tab.iloc[0]["pce"] - 0.025) < 1e-12
    md = markdown_tabela(tab, 2024, 2024)
    assert "IPCA" in md
    assert "CPI-U" in md
    assert formatar_pct(0.0483) == "4,83%"


def test_serie_anual_so_anos_completos():
    incompleto = _mensal_pct(2024, [0.4] * 12).iloc[:11]
    df = pd.concat([_mensal_pct(2023, [0.3] * 12), incompleto], ignore_index=True)
    anual = serie_anual(df, "percentual_mensal", [2023, 2024])
    assert 2023 in anual
    assert 2024 not in anual
