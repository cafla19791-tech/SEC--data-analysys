"""Testes do agregador streaming para massas grandes de fluxos_*.csv."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from scripts.agregar_impacto_fluxos import (
    MARKER,
    agregar_streaming,
    listar_csvs_fluxos,
    salvar_resultados,
)
from scripts.impacto_fiscal_por_ano import agregar_impacto_por_ano


def _escrever_csv(path: Path, rows: list[dict]) -> None:
    pd.DataFrame(rows).to_csv(path, index=False)


def test_marker_presente():
    assert "streaming" in MARKER


def test_listar_csvs_ignora_diario(tmp_path: Path):
    _escrever_csv(
        tmp_path / "fluxos_A.csv",
        [
            {
                "data_fluxo": "2010-01-15",
                "subsidio": 10.0,
                "impacto_fiscal": 20.0,
            }
        ],
    )
    _escrever_csv(
        tmp_path / "fluxos_diarios_A.csv",
        [
            {
                "data_fluxo": "2010-01-15",
                "subsidio": 1.0,
                "impacto_fiscal": 2.0,
            }
        ],
    )
    csvs = listar_csvs_fluxos(tmp_path)
    assert len(csvs) == 1
    assert csvs[0].name == "fluxos_A.csv"


def test_agregar_streaming_modo_coluna_multiplos_arquivos(tmp_path: Path):
    _escrever_csv(
        tmp_path / "fluxos_parte1.csv",
        [
            {
                "contrato": 1,
                "Instituição Financeira": "BANCO A",
                "data_fluxo": "2009-04-15",
                "subsidio": 100.0,
                "impacto_fiscal": 1000.0,
                "mes": 1,
            },
            {
                "contrato": 2,
                "Instituição Financeira": "BANCO B",
                "data_fluxo": "2010-04-15",
                "subsidio": 50.0,
                "impacto_fiscal": 400.0,
                "mes": 1,
            },
        ],
    )
    _escrever_csv(
        tmp_path / "fluxos_parte2.csv",
        [
            {
                "contrato": 3,
                "Instituição Financeira": "BANCO A",
                "data_fluxo": "2010-05-15",
                "subsidio": 50.0,
                "impacto_fiscal": 350.0,
                "mes": 2,
            },
        ],
    )

    result = agregar_streaming(
        listar_csvs_fluxos(tmp_path),
        modo="coluna",
        chunksize=2,
    )

    assert result["parcelas"] == 3
    por_ano = result["por_ano"].set_index("Ano")
    assert por_ano.loc[2009, "Soma Subsídio Nominal (R$)"] == 100.0
    assert por_ano.loc[2009, "Impacto Fiscal 2026 (R$)"] == 1000.0
    assert por_ano.loc[2010, "Soma Subsídio Nominal (R$)"] == 100.0
    assert por_ano.loc[2010, "Impacto Fiscal 2026 (R$)"] == 750.0
    assert int(por_ano["Quantidade de Parcelas"].sum()) == 3

    # Paridade com agregador in-memory (modo coluna)
    df_all = pd.concat(
        [pd.read_csv(p) for p in listar_csvs_fluxos(tmp_path)],
        ignore_index=True,
    )
    ref = agregar_impacto_por_ano(df_all, modo="coluna")
    merged = por_ano.reset_index().merge(ref, on="Ano", suffixes=("_s", "_r"))
    for col in (
        "Soma Subsídio Nominal (R$)",
        "Impacto Fiscal 2026 (R$)",
        "Quantidade de Parcelas",
    ):
        assert (merged[f"{col}_s"] == merged[f"{col}_r"]).all()

    por_ag = result["por_agente"].set_index("Instituição Financeira")
    assert por_ag.loc["BANCO A", "Impacto Fiscal 2026 (R$)"] == 1350.0
    assert por_ag.loc["BANCO A", "Qtd Contratos"] == 2
    assert por_ag.loc["BANCO B", "Impacto Fiscal 2026 (R$)"] == 400.0


def test_salvar_resultados(tmp_path: Path):
    _escrever_csv(
        tmp_path / "fluxos_x.csv",
        [
            {
                "contrato": 1,
                "Instituição Financeira": "X",
                "data_fluxo": "2015-01-15",
                "subsidio": 1.0,
                "impacto_fiscal": 2.0,
            }
        ],
    )
    result = agregar_streaming(listar_csvs_fluxos(tmp_path), modo="coluna")
    out = tmp_path / "out"
    paths = salvar_resultados(result, out)
    assert paths["impacto_ano_xlsx"].exists()
    assert paths["agente_xlsx"].exists()
    assert paths["workbook"].exists()
    xl = pd.ExcelFile(paths["workbook"])
    assert "Impacto_Por_Ano" in xl.sheet_names
    assert "Por_Agente" in xl.sheet_names
    assert "Totais" in xl.sheet_names


def test_modo_coluna_sem_impacto_falha(tmp_path: Path):
    _escrever_csv(
        tmp_path / "fluxos_y.csv",
        [{"data_fluxo": "2010-01-15", "subsidio": 1.0}],
    )
    with pytest.raises(ValueError, match="impacto"):
        agregar_streaming(listar_csvs_fluxos(tmp_path), modo="coluna")


def test_agregar_por_ano_contrato(tmp_path: Path):
    """Parcelas em anos diferentes do fluxo somam no ano da contratacao."""
    _escrever_csv(
        tmp_path / "fluxos_c.csv",
        [
            {
                "contrato": 1,
                "Instituição Financeira": "BNDES",
                "data_fluxo": "2010-04-15",
                "ano_contrato": 2009,
                "data_contratacao": "2009-03-15",
                "subsidio": 100.0,
                "impacto_fiscal": 1000.0,
            },
            {
                "contrato": 1,
                "Instituição Financeira": "BNDES",
                "data_fluxo": "2011-04-15",
                "ano_contrato": 2009,
                "data_contratacao": "2009-03-15",
                "subsidio": 50.0,
                "impacto_fiscal": 400.0,
            },
            {
                "contrato": 2,
                "Instituição Financeira": "BNDES",
                "data_fluxo": "2012-01-15",
                "ano_contrato": 2011,
                "data_contratacao": "2011-06-01",
                "subsidio": 20.0,
                "impacto_fiscal": 100.0,
            },
        ],
    )
    result = agregar_streaming(
        listar_csvs_fluxos(tmp_path),
        modo="coluna",
        agrupar_por="contrato",
    )
    assert result["agrupar_por"] == "contrato"
    por_ano = result["por_ano"].set_index("Ano do Contrato")
    assert list(por_ano.index.astype(int)) == [2009, 2011]
    assert por_ano.loc[2009, "Soma Subsídio Nominal (R$)"] == 150.0
    assert por_ano.loc[2009, "Impacto Fiscal 2026 (R$)"] == 1400.0
    assert int(por_ano.loc[2009, "Quantidade de Parcelas"]) == 2
    assert por_ano.loc[2011, "Impacto Fiscal 2026 (R$)"] == 100.0

    out = tmp_path / "out"
    paths = salvar_resultados(result, out)
    xl = pd.ExcelFile(paths["workbook"])
    assert "Impacto_Por_Ano_Contrato" in xl.sheet_names
    assert "Ano do Contrato" in pd.read_excel(paths["workbook"], sheet_name="Impacto_Por_Ano_Contrato").columns


def test_agregar_contrato_sem_coluna_falha(tmp_path: Path):
    _escrever_csv(
        tmp_path / "fluxos_z.csv",
        [
            {
                "data_fluxo": "2010-01-15",
                "subsidio": 1.0,
                "impacto_fiscal": 2.0,
            }
        ],
    )
    with pytest.raises(ValueError, match="ano_contrato|data_contratacao"):
        agregar_streaming(
            listar_csvs_fluxos(tmp_path), modo="coluna", agrupar_por="contrato"
        )
