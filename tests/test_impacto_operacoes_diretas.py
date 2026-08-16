"""Testes do pipeline de impacto das OPERACOES DIRETAS."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

from scripts.gerar_fluxos import (
    AGENTE_BNDES_DIRETA,
    SelicSerie,
    _mapear_colunas_contratos,
    _prepare_contracts,
    load_from_excel,
)
from scripts.impacto_operacoes_diretas import (
    MARKER,
    agregar_diretas,
    apresentar_diretas,
    gerar_fluxos_diretas,
    resolver_excel_diretas,
)


def _serie() -> SelicSerie:
    datas = np.array(
        [
            np.datetime64("2009-01-01"),
            np.datetime64("2009-02-16"),
            np.datetime64("2026-06-30"),
        ],
        dtype="datetime64[ns]",
    )
    return SelicSerie(datas, np.array([1.0, 1.5, 3.0]))


def _excel_diretas(path: Path) -> None:
    pd.DataFrame(
        {
            "Data da Contratação": ["15/03/2009", "20/04/2009"],
            "Valor Contratado Reais": [100000.0, 50000.0],
            "Juros": [6.0, 2.0],
            "Prezo - carencia (meses)": [0, 0],
            "Prazo - amortizaca (meses)": [3, 2],
            "Forma de Apoio": ["DIRETA", "DIRETA"],
            "Custo financeiro": ["TAXA FIXA", "TJLP"],
            "Instituição Financeira Credenciada": ["-", "-"],
        }
    ).to_excel(path, index=False)


def test_marker():
    assert "impacto-operacoes-diretas" in MARKER


def test_mapear_valor_contratado_e_typos(tmp_path: Path):
    excel = tmp_path / "OPERACOES DIRETAS.xlsx"
    _excel_diretas(excel)
    df = load_from_excel(excel, header=0)
    assert "valor_desembolsado" in df.columns
    assert len(df) == 2
    assert set(df["agente"].unique()) == {AGENTE_BNDES_DIRETA}


def test_mapear_so_contratado():
    raw = pd.DataFrame(
        {
            "data_da_contratacao": ["2009-03-15"],
            "valor_contratado_reais": [1000.0],
            "juros": [6.0],
            "prazo_carencia_meses": [0],
            "prazo_amortizacao_meses": [2],
            "forma_de_apoio": ["DIRETA"],
        }
    )
    mapped, _ = _mapear_colunas_contratos(raw)
    out = _prepare_contracts(mapped)
    assert float(out.loc[0, "valor_desembolsado"]) == 1000.0
    assert out.loc[0, "agente"] == AGENTE_BNDES_DIRETA


def test_pipeline_fluxos_agregar_apresentacao(tmp_path: Path):
    excel = tmp_path / "OPERACOES DIRETAS.xlsx"
    _excel_diretas(excel)
    assert resolver_excel_diretas(excel) == excel

    pasta_fluxos = tmp_path / "fluxos_diretas"
    pasta_impacto = tmp_path / "impacto_diretas"
    pasta_saida = tmp_path / "saida"
    pasta_saida.mkdir()

    csv_path = gerar_fluxos_diretas(excel, pasta_fluxos, fatores=_serie(), header=0)
    assert csv_path.exists()
    df = pd.read_csv(csv_path)
    assert "impacto_fiscal" in df.columns
    assert len(df) == 5  # 3+2

    info = agregar_diretas(pasta_fluxos, pasta_impacto)
    assert (pasta_impacto / "impacto_fiscal_por_ano.csv").exists()
    assert (pasta_impacto / "resumo_por_agente.csv").exists()
    assert info["result"]["parcelas"] == 5

    out = apresentar_diretas(pasta_impacto, pasta_saida)
    assert out.exists()
    assert out.name == "APRESENTACAO_IMPACTO_BNDES_DIRETAS.xlsx"
    wb = load_workbook(out)
    assert "Impacto Fiscal" in str(wb["Capa"]["A1"].value)
    assert "Diretas" in str(wb["Capa"]["A1"].value)
