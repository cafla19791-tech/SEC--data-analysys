"""Discriminativos de indiretas por ano do contrato (não do fluxo)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from scripts.discriminativos_indiretas_ano_contrato import (
    main,
    repartir_streaming,
)
from scripts.gerar_fluxos import gerar_fluxos_contrato


def test_fluxo_carrega_ano_contrato():
    rows = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2022-12-12"),
        valor=180_000.0,
        taxa_juros_aa=0.06,
        carencia=0,
        n=180,
        contrato_id=7,
        instituicao="BANCO X",
        selic_aa=0.145,
        data_impacto=datetime(2026, 6, 30),
        custo_financeiro="TAXA FIXA",
        juros_pct=6.0,
    )
    assert len(rows) == 180
    assert {r["ano_contrato"] for r in rows} == {2022}
    assert rows[0]["data_contratacao"].year == 2022
    # parcelas atravessam vários anos de fluxo
    anos_fluxo = {r["ano_fluxo"] for r in rows}
    assert 2022 in anos_fluxo
    assert max(anos_fluxo) > 2022


def test_repartir_por_ano_contrato(tmp_path: Path):
    # contrato 2022 com parcelas em 2022 e 2023
    rows_a = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2022-12-15"),
        valor=12_000.0,
        taxa_juros_aa=0.05,
        carencia=0,
        n=3,
        contrato_id=1,
        selic_aa=0.145,
        data_impacto=datetime(2026, 6, 30),
        juros_pct=5.0,
        custo_financeiro="TAXA FIXA",
    )
    # contrato 2009
    rows_b = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2009-03-15"),
        valor=24_000.0,
        taxa_juros_aa=0.06,
        carencia=0,
        n=2,
        contrato_id=2,
        selic_aa=0.145,
        data_impacto=datetime(2026, 6, 30),
        juros_pct=6.0,
        custo_financeiro="TAXA FIXA",
    )
    pasta = tmp_path / "saida"
    pasta.mkdir()
    pd.DataFrame(rows_a + rows_b).to_csv(pasta / "fluxos_teste.csv", index=False)

    out = tmp_path / "disc"
    stats = repartir_streaming([pasta / "fluxos_teste.csv"], out, chunksize=10)
    assert set(stats) == {2009, 2022}
    assert stats[2022]["qtd_parcelas"] == 3
    assert stats[2009]["qtd_parcelas"] == 2
    assert stats[2022]["qtd_contratos"] == 1

    df2022 = pd.read_csv(out / "fluxos_ano_contrato_2022.csv")
    assert len(df2022) == 3
    assert set(df2022["ano_contrato"]) == {2022}
    # mesmo com data_fluxo em 2023, permanece no arquivo 2022
    assert df2022["ano_fluxo"].max() >= 2022


def test_main_cli(tmp_path: Path):
    rows = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2020-01-15"),
        valor=6000.0,
        taxa_juros_aa=0.05,
        carencia=0,
        n=2,
        contrato_id=9,
        selic_aa=0.145,
        data_impacto=datetime(2026, 6, 30),
        juros_pct=5.0,
        custo_financeiro="TAXA FIXA",
    )
    pasta = tmp_path / "saida"
    pasta.mkdir()
    pd.DataFrame(rows).to_csv(pasta / "fluxos_0.csv", index=False)
    saida = tmp_path / "out"
    rc = main(["--pasta", str(pasta), "--saida", str(saida)])
    assert rc == 0
    assert (saida / "RESUMO_POR_ANO_CONTRATO.xlsx").exists()
    assert (saida / "fluxos_ano_contrato_2020.csv").exists()
