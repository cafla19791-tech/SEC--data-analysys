"""Testes unitários do gerador de fluxos detalhados (carência + impacto)."""

from datetime import datetime

import numpy as np
import pandas as pd

from scripts.gerar_fluxos import (
    SelicSerie,
    gerar_fluxos_contrato,
    limpar_valor,
    meses_ate_impacto,
    parse_datas,
    taxa_mensal_composta,
)


def test_limpar_valor_br_e_us():
    s = pd.Series(["1.234,56", "5.0", "10,5", "1000"])
    out = limpar_valor(s)
    assert list(out) == [1234.56, 5.0, 10.5, 1000.0]


def test_parse_datas_iso_e_br():
    s = pd.Series(["2009-03-15", "15/03/2009", "2009-03-15T00:00:00"])
    out = parse_datas(s)
    assert list(out.dt.strftime("%Y-%m-%d")) == [
        "2009-03-15",
        "2009-03-15",
        "2009-03-15",
    ]


def test_meses_ate_impacto():
    assert meses_ate_impacto(datetime(2026, 6, 30)) == 0
    assert meses_ate_impacto(datetime(2025, 6, 30)) == 12
    assert meses_ate_impacto(datetime(2026, 1, 15)) == 5


def test_taxa_mensal_composta():
    m = taxa_mensal_composta(0.145)
    assert abs(m - ((1.145) ** (1 / 12) - 1)) < 1e-12
    assert m != 0.145 / 12  # composta ≠ linear


def test_carencia_nao_consome_amortizacao():
    """Bug original: data=contr+(carencia+p) E em_carencia=p<=carencia no loop 1..n.

    Com carencia=2 e n=3, o saldo final deve zerar e deve haver 2 meses de carência.
    """
    data = pd.Timestamp("2009-01-31")
    fluxos = gerar_fluxos_contrato(
        data_contr=data,
        valor=300.0,
        taxa_juros_aa=0.06,
        carencia=2,
        n=3,
        contrato_id=0,
        instituicao="BANCO TESTE SA",
    )

    assert len(fluxos) == 5  # 2 carência + 3 amort
    assert [f["em_carencia"] for f in fluxos] == [True, True, False, False, False]
    assert [f["amortizacao"] for f in fluxos] == [0.0, 0.0, 100.0, 100.0, 100.0]
    assert fluxos[0]["saldo"] == 300.0
    assert fluxos[2]["saldo"] == 300.0  # ainda cheio ao sair da carência
    assert fluxos[-1]["saldo"] == 100.0
    # ContAgil: base no dia 15
    assert str(fluxos[0]["data_fluxo"]) == "2009-01-15"
    assert str(fluxos[1]["data_fluxo"]) == "2009-02-15"
    assert str(fluxos[2]["data_fluxo"]) == "2009-03-15"
    assert fluxos[0]["Instituição Financeira"] == "BANCO TESTE SA"
    assert "taxa_selic_mensal" in fluxos[0]
    assert "taxa_contrato_mensal" in fluxos[0]
    assert "spread" in fluxos[0]
    assert "impacto_fiscal" in fluxos[0]


def test_spread_e_taxas_compostas_constantes_no_contrato():
    fluxos = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2010-01-15"),
        valor=1200.0,
        taxa_juros_aa=0.06,
        carencia=0,
        n=12,
        contrato_id=1,
        selic_aa=0.145,
    )
    selic_m = taxa_mensal_composta(0.145)
    contrato_m = taxa_mensal_composta(0.06)
    expected_spread = (1.0 + (selic_m - contrato_m)) ** 12

    assert all(abs(f["taxa_selic_mensal"] - selic_m) < 1e-8 for f in fluxos)
    assert all(abs(f["taxa_contrato_mensal"] - contrato_m) < 1e-8 for f in fluxos)
    assert all(abs(f["spread"] - expected_spread) < 1e-6 for f in fluxos)
    # subsídio = saldo × (selic_m − contrato_m)
    assert abs(fluxos[0]["subsidio"] - round(1200.0 * (selic_m - contrato_m), 2)) < 0.011


def test_sem_carencia_amortiza_tudo():
    fluxos = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2010-01-15"),
        valor=1000.0,
        taxa_juros_aa=0.10,
        carencia=0,
        n=4,
        contrato_id=1,
    )
    assert len(fluxos) == 4
    assert all(not f["em_carencia"] for f in fluxos)
    assert abs(sum(f["amortizacao"] for f in fluxos) - 1000.0) < 0.01


def test_impacto_capitaliza_ate_jun_2026():
    fluxos = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2026-06-15"),
        valor=1200.0,
        taxa_juros_aa=0.0,  # taxa 0 → subsídio = saldo * selic_m
        carencia=0,
        n=1,
        contrato_id=2,
        selic_aa=0.12,
        data_impacto=datetime(2026, 6, 30),
    )
    assert len(fluxos) == 1
    # data_fluxo = 2026-06-15 → 0 meses até jun/2026 → impacto == subsidio
    assert fluxos[0]["impacto_fiscal"] == fluxos[0]["subsidio"]


def test_impacto_via_fator_selic_stp():
    # Fatores sintéticos: dobro em 12 meses
    datas = np.array(
        [
            np.datetime64("2025-06-15"),
            np.datetime64("2026-06-30"),
        ],
        dtype="datetime64[ns]",
    )
    fatores = np.array([1.0, 2.0], dtype=float)
    serie = SelicSerie(datas, fatores)

    fluxos = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2025-06-15"),
        valor=1000.0,
        taxa_juros_aa=0.0,
        carencia=0,
        n=1,
        contrato_id=3,
        selic_aa=0.12,
        data_impacto=datetime(2026, 6, 30),
        selic_serie=serie,
    )
    assert len(fluxos) == 1
    # impacto ≈ subsidio × 2/1
    assert abs(fluxos[0]["impacto_fiscal"] - round(fluxos[0]["subsidio"] * 2.0, 2)) < 0.02
