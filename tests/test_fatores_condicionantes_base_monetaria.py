"""Testes dos saldos anuais dos fatores condicionantes da base monetária."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.fatores_condicionantes_base_monetaria import (
    SERIES,
    baixar_serie,
    formatar_milhoes,
    gravar_saidas,
    markdown_tabela,
    montar_tabelas,
    saldo_fim_de_ano,
)


def _serie(codigo: int, datas: list[str], valores: list[float]) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "codigo": codigo,
            "data": pd.to_datetime(datas),
            "valor": valores,
        }
    )


def test_saldo_fim_de_ano_usa_dezembro():
    df = _serie(
        1810,
        ["1995-11-30", "1995-12-31", "1996-12-31", "1996-06-30"],
        [10.0, 20.0, 40.0, 30.0],
    )
    out = saldo_fim_de_ano(df, [1995, 1996])
    assert list(out["ano"]) == [1995, 1996]
    assert list(out["valor_rs_mil"]) == [20.0, 40.0]
    assert set(out["fechamento"]) == {"fim_de_ano"}


def test_saldo_ano_corrente_usa_ultimo_mes():
    df = _serie(1788, ["2026-06-30", "2026-07-31"], [100.0, 110.0])
    out = saldo_fim_de_ano(df, [2026])
    assert len(out) == 1
    assert out.iloc[0]["valor_rs_mil"] == 110.0
    assert out.iloc[0]["fechamento"] == "ultimo_disponivel_07/2026"


def test_montar_tabelas_converte_para_milhoes():
    catalogo = [SERIES[0], SERIES[-1]]  # conta única + base
    series = {
        1810: _serie(1810, ["1995-12-31", "1996-12-31"], [1_000.0, 2_000.0]),
        1788: _serie(1788, ["1995-12-31", "1996-12-31"], [5_000.0, 8_000.0]),
    }
    longo, largo = montar_tabelas(series, [1995, 1996], catalogo=catalogo)
    assert set(longo["serie"]) == {
        "Tesouro Nacional — Conta única",
        "Base monetária restrita (resultado)",
    }
    assert largo.loc[largo["ano"] == 1995, "tesouro_conta_unica"].iloc[0] == 1.0
    assert largo.loc[largo["ano"] == 1996, "base_monetaria_restrita"].iloc[0] == 8.0
    assert list(largo["fechamento"]) == ["fim_de_ano", "fim_de_ano"]


def test_baixar_serie_renomeia_mes(monkeypatch):
    def fake_baixar(codigo, inicio="01/01/1995", fim=None):
        return pd.DataFrame(
            {"mes": pd.to_datetime(["1995-12-01"]), "valor": [123.0]}
        )

    out = baixar_serie(1810, baixar=fake_baixar)
    assert list(out.columns) == ["codigo", "data", "valor"]
    assert out.iloc[0]["codigo"] == 1810
    assert out.iloc[0]["valor"] == 123.0


def test_gravar_saidas_e_markdown(tmp_path: Path):
    catalogo = [SERIES[0]]
    series = {1810: _serie(1810, ["1995-12-31"], [1_500_000.0])}
    longo, largo = montar_tabelas(series, [1995], catalogo=catalogo)
    caminhos = gravar_saidas(longo, largo, tmp_path, stem="teste_fatores")
    assert caminhos["csv_largo"].exists()
    assert caminhos["xlsx"].exists()
    md = caminhos["md"].read_text(encoding="utf-8")
    assert "1.500,0" in md or formatar_milhoes(1500.0) in md
    assert "1810" in markdown_tabela(largo, catalogo=catalogo)
    xl = pd.ExcelFile(caminhos["xlsx"])
    assert "Anual_R$_milhoes" in xl.sheet_names
    assert "Codigos_SGS" in xl.sheet_names
