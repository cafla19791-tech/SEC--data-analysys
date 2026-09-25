"""Identidades do discriminativo mensal da DBGG (Tabela 18 do BCB)."""

from pathlib import Path

import pandas as pd

from scripts.discriminativo_fatores_dbgg import _campo

CSV = Path(__file__).resolve().parents[1] / "output" / "discriminativo_fatores_dbgg_2007_2026.csv"


def test_rotulos_da_tabela_18():
    assert _campo("Fatores condicionantes:1/") == "fatores"
    assert _campo("Dívida bruta do governo geral - fatores condicionantes") is None
    assert _campo("Tabela 18 – Dívida bruta do Governo Geral – Fatores condicionantes") is None
    assert _campo("Dívida mobiliária interna indexada ao câmbio") == "interna_cambio"
    assert _campo("Dívida interna indexada ao câmbio") == "interna_cambio"
    assert _campo("Efeito crescimento PIB – dívida3/") == "efeito_pib"


def test_serie_mensal_fecha_as_identidades():
    df = pd.read_csv(CSV, parse_dates=["Mês"])
    assert len(df) == 235
    assert df["Mês"].min() == pd.Timestamp("2007-01-31")
    assert df["Mês"].max() == pd.Timestamp("2026-07-31")
    assert df.isna().sum().sum() == 0

    nec = df["Emissões líquidas (R$ milhões)"] + df["Juros nominais (R$ milhões)"]
    ajuste = (
        df["Dívida interna indexada ao câmbio (R$ milhões)"]
        + df["Dívida externa — metodológico (R$ milhões)"]
    )
    fatores = (
        df["Necessidade de financiamento da DBGG (R$ milhões)"]
        + df["Ajuste cambial (R$ milhões)"]
        + df["Dívida externa — outros ajustes (R$ milhões)"]
        + df["Reconhecimento de dívidas (R$ milhões)"]
        + df["Privatizações (R$ milhões)"]
    )
    fecha = (
        df["Dívida bruta — abertura (R$ milhões)"]
        + df["Variação mensal da dívida bruta (R$ milhões)"]
    )
    var_pib = (
        df["Fatores condicionantes (% PIB)"]
        + df["Efeito do crescimento do PIB sobre a dívida (p.p. do PIB)"]
    )

    assert (nec - df["Necessidade de financiamento da DBGG (R$ milhões)"]).abs().max() < 1e-4
    assert (ajuste - df["Ajuste cambial (R$ milhões)"]).abs().max() < 1e-4
    assert (fatores - df["Fatores condicionantes (R$ milhões)"]).abs().max() < 1e-4
    assert (fecha - df["Dívida bruta — fechamento (R$ milhões)"]).abs().max() < 1e-3
    assert (var_pib - df["Variação mensal da dívida bruta (p.p. do PIB)"]).abs().max() < 1e-4

    dez24 = df.loc[df["Mês"] == "2024-12-31"].iloc[0]
    assert abs(dez24["Dívida bruta — fechamento (R$ milhões)"] - 8_984_236.594450) < 1
    # Razão da nota de mai/2026 (última que republica dez/2024). Revisões posteriores
    # do PIB nominal aparecem na Tabela 19 mais recente (76,27%) sem reescrever o mês.
    assert abs(dez24["Dívida bruta — fechamento (% PIB)"] - 76.496027) < 0.01

    jul26 = df.loc[df["Mês"] == "2026-07-31"].iloc[0]
    assert abs(jul26["Dívida bruta — fechamento (% PIB)"] - 82.509085) < 0.01
