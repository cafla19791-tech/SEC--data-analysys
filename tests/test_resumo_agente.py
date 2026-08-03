"""Testes do resumo por agente financeiro."""

from pathlib import Path

import pandas as pd

from scripts.gerar_fluxos import (
    AGENTE_NAO_INFORMADO,
    agregar_por_agente,
    gerar_fluxos,
    load_from_csv,
    processar_em_lotes,
    resumo_from_agent_agg,
)


SAMPLE = Path(__file__).resolve().parents[1] / "data" / "sample_operacoes_com_agente.csv"


def test_load_sample_preserva_agente():
    df = load_from_csv(SAMPLE)
    assert "agente" in df.columns
    assert df["agente"].nunique() >= 5
    assert AGENTE_NAO_INFORMADO not in set(df["agente"]) or df["agente"].nunique() > 1


def test_agregar_por_agente_liga_por_contrato():
    contratos = load_from_csv(SAMPLE).head(5)
    fluxos = gerar_fluxos(contratos)
    resumo = agregar_por_agente(fluxos, contratos)

    assert list(resumo.columns) == [
        "Agente",
        "Qtd Contratos",
        "Total Subsídio (R$)",
        "Impacto Fiscal 2026 (R$)",
    ]
    assert len(resumo) >= 2
    # ordenado por subsídio desc
    assert resumo["Total Subsídio (R$)"].is_monotonic_decreasing
    # soma de contratos no resumo = contratos processados
    assert int(resumo["Qtd Contratos"].sum()) == len(contratos)


def test_processar_em_lotes_acumula_por_agente(tmp_path: Path):
    contratos = load_from_csv(SAMPLE)
    csv_path = tmp_path / "fluxos.csv"
    stats = processar_em_lotes(contratos, csv_path, lote=5)

    assert stats["n_agentes"] >= 5
    resumo = resumo_from_agent_agg(stats["por_agente"])
    assert int(resumo["Qtd Contratos"].sum()) == stats["n_contratos_ok"]
    assert abs(resumo["Total Subsídio (R$)"].sum() - stats["total_subsidio"]) < 1.0
