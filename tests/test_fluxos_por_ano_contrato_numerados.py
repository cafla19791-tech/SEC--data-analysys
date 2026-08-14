"""Fluxos a partir de BNDES_INDIRETAS_NUMERADOS (uma aba por ano)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from scripts.fluxos_por_ano_contrato_numerados import listar_abas_ano, processar
from scripts.gerar_fluxos import gerar_fluxos_contrato, normalizar_colunas


def _make_numerados(path: Path) -> Path:
    header = [
        "Número do contrato",
        "Data da contratação",
        "Valor desembolsado R$",
        "Instituição Financeira Credenciada",
        "Custo financeiro",
        "Juros",
        "Prazo - Carência (meses)",
        "Prazo - Amortização (meses)",
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "2002"
    ws.append(header)
    ws.append(["1-2002", "2002-01-10", 1000, "BANCO X", "TAXA FIXA", 5, 0, 3])
    ws.append(["2-2002", "2002-06-10", 2000, "BANCO X", "TAXA FIXA", 5, 0, 3])
    ws2 = wb.create_sheet("2003")
    ws2.append(header)
    ws2.append(["1-2003", "2003-03-15", 1500, "BANCO Y", "TAXA FIXA", 4, 0, 2])
    wb.save(path)
    return path


def test_listar_abas_ano(tmp_path: Path):
    p = _make_numerados(tmp_path / "num.xlsx")
    assert listar_abas_ano(p) == ["2002", "2003"]


def test_numero_contrato_no_fluxo():
    rows = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2022-12-12"),
        valor=3000.0,
        taxa_juros_aa=0.05,
        carencia=0,
        n=3,
        contrato_id=0,
        juros_pct=5.0,
        custo_financeiro="TAXA FIXA",
        numero_contrato="7-2022",
        selic_aa=0.145,
    )
    assert len(rows) == 3
    assert {r["numero_contrato"] for r in rows} == {"7-2022"}
    assert {r["ano_contrato"] for r in rows} == {2022}


def test_processar_cria_abas_e_csv(tmp_path: Path):
    numerados = _make_numerados(tmp_path / "BNDES_INDIRETAS_NUMERADOS.xlsx")
    saida = tmp_path / "saida"
    xlsx = processar(numerados, saida, fatores=0.145, lote=10)
    assert xlsx.exists()
    wb = load_workbook(xlsx)
    assert "2002" in wb.sheetnames
    assert "2003" in wb.sheetnames
    assert "RESUMO" in wb.sheetnames

    csv2002 = saida / "fluxos_por_ano_contrato" / "2002.csv"
    assert csv2002.exists()
    df = pd.read_csv(csv2002)
    assert "ano_contrato" in df.columns
    assert set(df["ano_contrato"]) == {2002}
    assert "numero_contrato" in df.columns
    assert set(df["numero_contrato"]) == {"1-2002", "2-2002"}
    # 2 contratos × 3 parcelas
    assert len(df) == 6

    df3 = pd.read_csv(saida / "fluxos_por_ano_contrato" / "2003.csv")
    assert len(df3) == 2
    assert set(df3["numero_contrato"]) == {"1-2003"}


def test_retomar_pula_ano_existente(tmp_path: Path):
    numerados = _make_numerados(tmp_path / "BNDES_INDIRETAS_NUMERADOS.xlsx")
    saida = tmp_path / "saida"
    processar(numerados, saida, fatores=0.145, lote=10)
    csv2002 = saida / "fluxos_por_ano_contrato" / "2002.csv"
    mtime = csv2002.stat().st_mtime
    processar(numerados, saida, fatores=0.145, lote=10, retomar=True)
    assert csv2002.stat().st_mtime == mtime
    resumo = pd.read_csv(saida / "fluxos_por_ano_contrato" / "RESUMO.csv")
    assert set(resumo["status"]) == {"retomado"}


def test_normalizar_preserva_numero():
    bruto = pd.DataFrame(
        {
            "Número do contrato": ["1-2002"],
            "Data da contratação": ["2002-01-10"],
            "Valor desembolsado R$": [1000],
            "Instituição Financeira Credenciada": ["BANCO X"],
            "Custo financeiro": ["TAXA FIXA"],
            "Juros": [5],
            "Prazo - Carência (meses)": [0],
            "Prazo - Amortização (meses)": [3],
        }
    )
    out = normalizar_colunas(bruto)
    assert "numero_contrato" in out.columns
    assert out.loc[0, "numero_contrato"] == "1-2002"
