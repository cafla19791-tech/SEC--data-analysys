"""Testes dos fatores condicionantes da base monetária (SGS, fim de período)."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from scripts.fatores_condicionantes_base_monetaria import (
    ESTOQUES,
    FATORES_SOMA,
    SERIES,
    carregar_painel,
    estoque_fim,
    fluxo_ano,
    identidade_mensal,
    processar,
    soma_fatores_mes,
    tabela_anual,
    tabela_dezembro,
    tabela_mensal,
)


def _painel_sintetico() -> pd.DataFrame:
    """Dez/1999 a dez/2001 com identidade Σ fatores = Δ base."""
    meses = pd.date_range("1999-12-01", "2001-12-01", freq="MS")
    # Estoque da base: 100.000 em dez/99, sobe 1.000/mês em 2000 e 2.000/mês em 2001
    base = []
    nivel = 100_000.0
    for m in meses:
        if m.year >= 2000:
            nivel += 1_000.0 if m.year == 2000 else 2_000.0
        base.append(nivel)
    # Distribui o fluxo do mês entre os fatores (soma = Δ)
    tn, ext, tit, dep, red, der, lin, outr = [], [], [], [], [], [], [], []
    pmc, res = [], []
    prev = None
    for m, b in zip(meses, base):
        delta = 0.0 if prev is None else b - prev
        prev = b
        tn.append(0.60 * delta)
        ext.append(0.25 * delta)
        tit.append(-0.10 * delta)
        dep.append(0.15 * delta)
        red.append(0.0)
        der.append(0.05 * delta)
        lin.append(0.0)
        outr.append(0.05 * delta)  # 0.60+0.25-0.10+0.15+0.05+0.05 = 1.00
        pmc.append(0.80 * b)
        res.append(0.20 * b)
    prim = [t * 0.4 for t in tit]
    sec = [t * 0.6 for t in tit]
    return pd.DataFrame(
        {
            1810: tn,
            1811: ext,
            1809: tit,
            29004: prim,
            29006: sec,
            1815: dep,
            12484: red,
            12487: der,
            28724: lin,
            1818: outr,
            1788: base,
            1786: pmc,
            1787: res,
        },
        index=meses,
    )


def test_catalogo_oficial():
    assert {s.codigo for s in FATORES_SOMA} == {
        1810,
        1811,
        1809,
        1815,
        12484,
        12487,
        28724,
        1818,
    }
    assert 29004 not in {s.codigo for s in FATORES_SOMA}
    assert {s.codigo for s in ESTOQUES} == {1788, 1786, 1787}
    assert len(SERIES) == 13


def test_identidade_mensal_sintetica():
    p = _painel_sintetico()
    ident = identidade_mensal(p)
    assert not ident.empty
    assert (ident["residuo"].abs() < 1e-6).all()
    dez00 = pd.Timestamp("2000-12-01")
    assert soma_fatores_mes(p, dez00) == pytest.approx(1_000.0)
    assert estoque_fim(p, 1788, 2000) == pytest.approx(112_000.0)


def test_fluxo_anual_e_delta_estoque():
    p = _painel_sintetico()
    # 12 meses × 1.000 em 2000
    assert fluxo_ano(p, 1810, 2000) == pytest.approx(0.60 * 12_000)
    assert fluxo_ano(p, 1809, 2000) == pytest.approx(-0.10 * 12_000)
    anual = tabela_anual(p, [2000, 2001])
    soma = anual.loc[anual["Item"].str.startswith("Variação da base"), 2000].iloc[0]
    delta = anual.loc[anual["Item"].str.startswith("Memória: Δ estoque"), 2000].iloc[0]
    assert soma == pytest.approx(12_000.0)
    assert delta == pytest.approx(12_000.0)
    # 2001: 12 × 2.000
    soma01 = anual.loc[anual["Item"].str.startswith("Variação da base"), 2001].iloc[0]
    assert soma01 == pytest.approx(24_000.0)
    base01 = anual.loc[anual["Item"] == "Base monetária restrita", 2001].iloc[0]
    assert base01 == pytest.approx(136_000.0)


def test_primario_mais_secundario_igual_total():
    p = _painel_sintetico()
    assert fluxo_ano(p, 29004, 2001) + fluxo_ano(p, 29006, 2001) == pytest.approx(
        fluxo_ano(p, 1809, 2001)
    )


def test_dezembro_e_papel_moeda_mais_reservas():
    p = _painel_sintetico()
    dez = tabela_dezembro(p, [2000])
    tn = dez.loc[dez["SGS"] == 1810, 2000].iloc[0]
    assert tn == pytest.approx(600.0)
    pmc = estoque_fim(p, 1786, 2000)
    res = estoque_fim(p, 1787, 2000)
    assert pmc + res == pytest.approx(estoque_fim(p, 1788, 2000))


def test_mensal_tem_residuo():
    p = _painel_sintetico()
    m = tabela_mensal(p)
    assert "Σ fatores" in m.columns
    assert "Resíduo" in m.columns
    assert m.loc[m["ano"] == 2000, "Resíduo"].abs().max() < 1e-6


def _gravar_csvs(pasta: Path, painel: pd.DataFrame) -> dict[int, Path]:
    pasta.mkdir(parents=True, exist_ok=True)
    arquivos = {}
    for codigo in painel.columns:
        dest = pasta / f"{codigo}.csv"
        pd.DataFrame({"mes": painel.index, "valor": painel[codigo].to_numpy()}).to_csv(
            dest, index=False
        )
        arquivos[int(codigo)] = dest
    return arquivos


def test_carregar_painel_e_planilha(tmp_path: Path):
    painel = _painel_sintetico()
    arquivos = _gravar_csvs(tmp_path / "in", painel)
    saida = tmp_path / "fatores.xlsx"
    path = processar(
        pasta_cache=tmp_path / "cache",
        saida=saida,
        usar_cache=False,
        arquivos=arquivos,
    )
    assert path.exists()
    nomes = pd.ExcelFile(path).sheet_names
    assert nomes[0] == "Metodologia"
    for aba in ("Anual", "Dezembro", "Mensal", "Identidade", "Grafico"):
        assert aba in nomes
    anual = pd.read_excel(path, sheet_name="Anual", header=3)
    assert "Item" in anual.columns
    assert any("Tesouro Nacional" in str(x) for x in anual["Item"])
    # 2000 e 2001 presentes; 2026 ausente no sintético
    assert 2000 in anual.columns or "2000" in [str(c) for c in anual.columns]


def test_carregar_painel_local(tmp_path: Path):
    painel = _painel_sintetico()
    pasta = tmp_path / "in"
    pasta.mkdir()
    arquivos = _gravar_csvs(pasta, painel)
    out = carregar_painel(tmp_path / "cache", arquivos=arquivos)
    assert 1788 in out.columns
    assert out.index.min().date() == date(1999, 12, 1)
    assert out.index.max().date() == date(2001, 12, 1)
