"""Testes do cálculo do spread do banco (coluna Juros) por contrato."""

from __future__ import annotations

from datetime import datetime

import pandas as pd
import pytest

from scripts.gerar_fluxos import (
    calcular_spread_banco_contrato,
    gerar_fluxos_contrato,
    taxa_mensal_composta,
    taxa_spread_banco_mensal,
)


def test_taxa_spread_banco_mensal_seis_porcento():
    m = taxa_spread_banco_mensal(6.0)
    assert m == pytest.approx((1.06) ** (1.0 / 12.0) - 1.0, rel=1e-9)
    assert taxa_spread_banco_mensal(0) == 0.0
    assert taxa_spread_banco_mensal(None) == 0.0


def test_calcular_spread_banco_contrato_nominal():
    res = calcular_spread_banco_contrato(
        pd.Timestamp("2010-01-15"),
        valor=120_000.0,
        carencia=0,
        n=12,
        juros_pct=12.0,
    )
    taxa_m = taxa_spread_banco_mensal(12.0)
    # SAC: saldos 120k, 110k, ..., 10k
    esperado = sum((120_000.0 - i * 10_000.0) * taxa_m for i in range(12))
    assert res["parcelas"] == 12
    assert res["spread_banco_nominal"] == pytest.approx(round(esperado, 2), abs=0.02)


def test_spread_banco_na_parcela_bate_com_agregado():
    fluxos = gerar_fluxos_contrato(
        pd.Timestamp("2010-01-15"),
        valor=100_000.0,
        taxa_juros_aa=0.06,
        carencia=3,
        n=12,
        contrato_id=0,
        juros_pct=6.0,
        custo_financeiro="TJLP",
    )
    assert "spread_banco" in fluxos[0]
    # 1ª parcela: 100000 * taxa_banco
    assert fluxos[0]["spread_banco"] == pytest.approx(
        round(100_000.0 * taxa_spread_banco_mensal(6.0), 2), abs=0.01
    )
    soma = round(sum(f["spread_banco"] for f in fluxos), 2)
    res = calcular_spread_banco_contrato(
        pd.Timestamp("2010-01-15"),
        100_000.0,
        3,
        12,
        6.0,
    )
    assert soma == pytest.approx(res["spread_banco_nominal"], abs=0.05)


def test_spread_banco_diferente_do_fator_spread_contagil():
    """Coluna `spread` é fator ContAgil; `spread_banco` é R$."""
    fluxos = gerar_fluxos_contrato(
        pd.Timestamp("2010-01-15"),
        valor=50_000.0,
        taxa_juros_aa=0.05,
        carencia=0,
        n=6,
        contrato_id=1,
        juros_pct=5.0,
        custo_financeiro="TAXA FIXA",
        selic_aa=0.145,
    )
    selic_m = taxa_mensal_composta(0.145)
    contrato_m = taxa_mensal_composta(0.05)
    fator = (1.0 + (selic_m - contrato_m)) ** 6
    assert fluxos[0]["spread"] == pytest.approx(fator, rel=1e-5)
    assert fluxos[0]["spread_banco"] > 1.0  # valor em R$, não fator ~1.x
    assert fluxos[0]["spread_banco"] != pytest.approx(fluxos[0]["spread"], abs=0.5)


def test_cli_import_marker():
    from scripts import spread_banco_contratos as mod

    assert "spread-banco" in mod.MARKER
