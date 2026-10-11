"""Testes do discriminativo de taxas básicas reais por país e período."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from scripts.discriminativo_juros_reais_paises import (
    PERIODOS,
    Periodo,
    calcular_pais,
    calcular_periodos,
    cagr,
    completar_policy_euro,
    fator_inflacao_ipc,
    fator_nominal_policy,
    fator_proporcional,
    preparar_cpi,
    preparar_policy,
    processar,
    ranquear,
    resultados_para_df,
    taxa_real_fisher,
    escrever_planilha,
)


def test_periodos_oficiais_cinco_recortes():
    assert len(PERIODOS) == 5
    assert [p.id for p in PERIODOS] == [1, 2, 3, 4, 5]
    assert (PERIODOS[0].inicio, PERIODOS[0].fim) == (date(1995, 1, 1), date(2002, 12, 31))
    assert (PERIODOS[1].inicio, PERIODOS[1].fim) == (date(2003, 1, 1), date(2016, 5, 11))
    assert (PERIODOS[2].inicio, PERIODOS[2].fim) == (date(2016, 5, 12), date(2018, 12, 31))
    assert (PERIODOS[3].inicio, PERIODOS[3].fim) == (date(2019, 1, 1), date(2022, 12, 31))
    assert (PERIODOS[4].inicio, PERIODOS[4].fim) == (date(2023, 1, 1), date(2026, 8, 28))
    abas = [p.aba for p in PERIODOS]
    assert len(set(abas)) == 5
    assert abas[3].startswith("4_2019")
    assert abas[4].startswith("5_2023")


def test_fisher_basico():
    assert taxa_real_fisher(0.15, 0.05) == pytest.approx((1.15 / 1.05) - 1)
    assert taxa_real_fisher(0.10, 0.10) == pytest.approx(0.0)
    assert taxa_real_fisher(0.05, 0.10) < 0


def test_cagr_taxa_constante():
    # 10% a.a. capitalizado por 365,25 dias (base do script) → fator 1,10
    fat = (1.10)  # definição: (1+i)^{365,25/365,25}
    assert cagr(fat, 36525 / 100) == pytest.approx(0.10)
    assert fator_proporcional(0.10, 365) == pytest.approx(1.10 ** (365 / 365.25))


def test_fator_nominal_mes_cheio():
    policy = pd.Series({date(2000, 1, 1): 0.12})
    fat, dias, p0, p1 = fator_nominal_policy(
        policy, date(2000, 1, 1), date(2000, 1, 31)
    )
    assert dias == 31
    assert fat == pytest.approx((1.12) ** (1 / 12))
    assert p0 == date(2000, 1, 1)
    assert p1 == date(2000, 1, 1)


def test_fator_inflacao_prorata():
    ipc = pd.Series(
        {
            date(2016, 4, 1): 100.0,
            date(2016, 5, 1): 101.0,
        }
    )
    fat_11, dias_11, _, _ = fator_inflacao_ipc(
        ipc, date(2016, 5, 1), date(2016, 5, 11)
    )
    fat_20, dias_20, _, _ = fator_inflacao_ipc(
        ipc, date(2016, 5, 12), date(2016, 5, 31)
    )
    assert dias_11 == 11
    assert dias_20 == 20
    # 11 + 20 = 31 → produto dos pró-ratas reconstitui o MoM
    assert fat_11 * fat_20 == pytest.approx(1.01)
    assert fat_11 == pytest.approx(1.01 ** (11 / 31))


def test_calcular_pais_constante_fisher():
    """Taxa 15% a.a. e IPC +5% a.a. (≈0,407% a.m.) → real ≈ 9,52% a.a."""
    meses = pd.date_range("2000-01-01", "2001-12-01", freq="MS")
    policy = pd.Series({d.date(): 0.15 for d in meses})
    ipc_vals = {}
    nivel = 100.0
    mom = (1.05) ** (1 / 12)
    ipc_vals[date(1999, 12, 1)] = nivel
    for d in meses:
        nivel *= mom
        ipc_vals[d.date()] = nivel
    ipc = pd.Series(ipc_vals)
    yoy = pd.Series({d.date(): 0.05 for d in meses})
    periodo = Periodo(
        9, date(2000, 1, 1), date(2001, 12, 31), "x", "2000-2001"
    )
    r = calcular_pais("XX", policy, ipc, yoy, periodo)
    assert r.no_ranking
    assert r.cagr_real == pytest.approx(taxa_real_fisher(0.15, 0.05), rel=1e-6)
    assert r.taxa_real_acumulada > 0
    assert r.cagr_nominal == pytest.approx(0.15, rel=2e-3)
    assert r.cagr_inflacao == pytest.approx(0.05, rel=2e-3)


def test_ranking_decrescente_e_cobertura():
    periodo = Periodo(1, date(2000, 1, 1), date(2000, 12, 31), "p", "2000")

    def _pais(codigo: str, taxa: float, ipc0: float = 100.0, ipc1: float = 105.0):
        policy = pd.Series({date(2000, m, 1): taxa for m in range(1, 13)})
        ipc = pd.Series(
            {date(1999, 12, 1): ipc0, **{date(2000, m, 1): ipc1 for m in range(1, 13)}}
        )
        # IPC crescente uniforme
        niveis = {date(1999, 12, 1): ipc0}
        for m in range(1, 13):
            niveis[date(2000, m, 1)] = ipc0 * (ipc1 / ipc0) ** (m / 12)
        ipc = pd.Series(niveis)
        yoy = pd.Series({date(2000, m, 1): 0.05 for m in range(1, 13)})
        return calcular_pais(codigo, policy, ipc, yoy, periodo)

    alto = _pais("HI", 0.20)
    medio = _pais("MD", 0.10)
    baixo = _pais("LO", 0.02)
    # cobertura insuficiente: só 1 mês
    policy_curto = pd.Series({date(2000, 1, 1): 0.50})
    ipc_curto = pd.Series({date(1999, 12, 1): 100.0, date(2000, 1, 1): 101.0})
    yoy_curto = pd.Series({date(2000, 1, 1): 0.01})
    curto = calcular_pais("ZZ", policy_curto, ipc_curto, yoy_curto, periodo)

    ranked = ranquear([baixo, alto, curto, medio])
    assert [r.codigo for r in ranked if r.no_ranking] == ["HI", "MD", "LO"]
    assert ranked[0].ranking == 1
    assert ranked[0].taxa_real_acumulada >= ranked[1].taxa_real_acumulada
    assert any(r.codigo == "ZZ" and r.ranking == 0 for r in ranked)


def _bis_policy_csv(path: Path) -> Path:
    rows = [
        "FREQ,REF_AREA,UNIT_MEASURE,TIME_PERIOD,OBS_VALUE",
        "M,BR,368,1999-12,19.0",
        "M,BR,368,2000-01,19.0",
        "M,BR,368,2000-02,19.0",
        "M,US,368,1999-12,5.5",
        "M,US,368,2000-01,5.5",
        "M,US,368,2000-02,5.5",
    ]
    dest = path / "policy.csv"
    dest.write_text("\n".join(rows), encoding="utf-8")
    return dest


def _bis_cpi_csv(path: Path) -> Path:
    # unidade 628 = índice; 771 = YoY %
    rows = [
        "FREQ,REF_AREA,UNIT_MEASURE,TIME_PERIOD,OBS_VALUE",
        "M,BR,628,1999-12,100",
        "M,BR,628,2000-01,100.5",
        "M,BR,628,2000-02,101.0",
        "M,BR,771,2000-01,6.0",
        "M,BR,771,2000-02,6.0",
        "M,US,628,1999-12,100",
        "M,US,628,2000-01,100.2",
        "M,US,628,2000-02,100.4",
        "M,US,771,2000-01,3.0",
        "M,US,771,2000-02,3.0",
    ]
    dest = path / "cpi.csv"
    dest.write_text("\n".join(rows), encoding="utf-8")
    return dest


def test_completar_policy_euro():
    policy = pd.DataFrame(
        {
            "REF_AREA": ["DE", "DE", "XM", "XM", "XM"],
            "mes": [
                date(1998, 11, 1),
                date(1998, 12, 1),
                date(1998, 12, 1),
                date(1999, 1, 1),
                date(1999, 2, 1),
            ],
            "taxa_aa": [0.03, 0.03, 0.025, 0.03, 0.03],
        }
    )
    out = completar_policy_euro(policy)
    de = out[out.REF_AREA == "DE"].set_index("mes")["taxa_aa"]
    assert date(1998, 12, 1) in de.index
    assert de.loc[date(1998, 12, 1)] == pytest.approx(0.03)  # nacional prevalece
    assert de.loc[date(1999, 1, 1)] == pytest.approx(0.03)
    assert de.loc[date(1999, 2, 1)] == pytest.approx(0.03)


def test_preparar_series(tmp_path: Path):
    pol = preparar_policy(pd.read_csv(_bis_policy_csv(tmp_path)))
    ipc, yoy = preparar_cpi(pd.read_csv(_bis_cpi_csv(tmp_path)))
    assert set(pol["REF_AREA"]) == {"BR", "US"}
    assert pol.loc[pol.REF_AREA == "BR", "taxa_aa"].iloc[0] == pytest.approx(0.19)
    assert ipc.loc[ipc.REF_AREA == "BR", "ipc"].min() == 100
    assert yoy.loc[yoy.REF_AREA == "US", "inflacao_12m"].iloc[0] == pytest.approx(0.03)


def test_planilha_uma_aba_por_periodo(tmp_path: Path):
    meses = [date(2000, m, 1) for m in range(1, 3)]
    policy = pd.DataFrame(
        {
            "REF_AREA": ["BR", "BR", "US", "US"],
            "mes": meses + meses,
            "taxa_aa": [0.19, 0.19, 0.05, 0.05],
        }
    )
    ipc = pd.DataFrame(
        {
            "REF_AREA": ["BR"] * 3 + ["US"] * 3,
            "mes": [date(1999, 12, 1), *meses, date(1999, 12, 1), *meses],
            "ipc": [100, 100.5, 101.0, 100, 100.2, 100.4],
        }
    )
    yoy = pd.DataFrame(
        {
            "REF_AREA": ["BR", "BR", "US", "US"],
            "mes": meses + meses,
            "inflacao_12m": [0.06, 0.06, 0.03, 0.03],
        }
    )
    periodos = (
        Periodo(1, date(2000, 1, 1), date(2000, 2, 29), "1_teste", "jan-fev/2000"),
    )
    # monkeypatch PERIODOS used by escrever_planilha via calcular + write with custom
    por = calcular_periodos(policy, ipc, yoy, periodos, cobertura_minima=0.5)
    df = resultados_para_df(por[1])
    assert list(df["País"])[0] in {"Brasil", "Estados Unidos"}
    saida = tmp_path / "out.xlsx"
    from scripts import discriminativo_juros_reais_paises as mod

    original = mod.PERIODOS
    try:
        mod.PERIODOS = periodos
        escrever_planilha(por, saida, n_paises=2)
    finally:
        mod.PERIODOS = original
    xl = pd.ExcelFile(saida)
    assert "Metodologia" in xl.sheet_names
    assert "1_teste" in xl.sheet_names
    assert "Comparativo" in xl.sheet_names
    aba = pd.read_excel(saida, sheet_name="1_teste", header=3)
    assert "Taxa básica real acumulada" in aba.columns
    assert "CAGR real" in aba.columns
    # ranking decrescente (ignora linhas fora do ranking)
    validos = aba[aba["No ranking"] == "sim"]
    reais = validos["Taxa básica real acumulada"].astype(float)
    assert reais.is_monotonic_decreasing


def test_cli_com_csv_local(tmp_path: Path):
    policy = _bis_policy_csv(tmp_path)
    cpi = _bis_cpi_csv(tmp_path)
    saida = tmp_path / "disc.xlsx"
    # Períodos oficiais quase sem overlap com 2000-01/02 — ainda assim a planilha
    # precisa nascer com as 5 abas.
    path = processar(
        pasta_cache=tmp_path / "cache",
        saida=saida,
        usar_cache=False,
        policy_csv=policy,
        cpi_csv=cpi,
        cobertura_minima=0.01,
    )
    assert path.exists()
    nomes = pd.ExcelFile(path).sheet_names
    assert nomes[0] == "Metodologia"
    assert "Comparativo" in nomes
    for p in PERIODOS:
        assert p.aba in nomes
