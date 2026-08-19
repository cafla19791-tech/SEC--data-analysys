"""Testes do script ContAgil na raiz (gerar_fluxos.py)."""

from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import importlib.util
_spec = importlib.util.spec_from_file_location("gerar_fluxos_root", ROOT / "gerar_fluxos.py")
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)

calcular_impacto_fiscal_real = _mod.calcular_impacto_fiscal_real
gerar_fluxos = _mod.gerar_fluxos
load_selic = _mod.load_selic
parse_args = _mod.parse_args


def _selic_sintetica() -> pd.DataFrame:
    dates = pd.date_range("2009-01-01", "2026-06-30", freq="D")
    # ~0.01% a.d. → fator crescente
    fator = np.cumprod(np.full(len(dates), 1.0001))
    return pd.DataFrame({"data": dates, "fator": fator})


def test_parse_fluxo_diario_flag():
    args = parse_args(
        ["--excel", "x.xlsx", "--fluxo-diario", "--output-dir", "out", "--max-contratos", "3"]
    )
    assert args.fluxo_diario is True
    assert args.max_contratos == 3
    assert args.output_dir == "out"


def test_parse_massa_dados_contagil_cli():
    """CLI ContAgil WinPython: --massa-dados / --pasta-saida / --arquivo-selic."""
    args = parse_args(
        [
            "--massa-dados",
            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\dados",
            "--pasta-saida",
            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida",
            "--arquivo-selic",
            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\STP-20260716182715078 (1).xlsx",
        ]
    )
    assert args.massa_dados.endswith("dados")
    assert args.output_dir.endswith("saida")
    assert args.arquivo_selic.endswith("STP-20260716182715078 (1).xlsx")
    assert args.excel is None


def test_calcular_impacto_usa_data_da_parcela():
    selic = pd.DataFrame(
        {
            "data": pd.to_datetime(["2009-02-15", "2009-02-16", "2026-06-30"]),
            "fator": [1.0, 2.0, 4.0],
        }
    )
    # ContAgil col D: nearest na própria parcela (15/02 → fator 1) → 4/1 = 4×
    assert calcular_impacto_fiscal_real(100.0, datetime(2009, 2, 15), selic) == 400.0


def test_gerar_fluxos_resumo_e_diario(tmp_path: Path):
    ops = pd.DataFrame(
        {
            "data_da_contratacao": ["15/01/2009"],
            "valor_desembolsado_reais": [1200.0],
            "juros": [6.0],
            "prazo_carencia_meses": [1],
            "prazo_amortizacao_meses": [2],
            "instituicao_financeira_credenciada": ["BANCO TESTE"],
        }
    )
    selic = _selic_sintetica()
    diario = tmp_path / "fluxos_diarios_detalhados.xlsx"
    out = gerar_fluxos(
        ops, selic, fluxo_diario=True, max_contratos=1, saida_diario=diario
    )

    assert len(out) == 1
    assert out.iloc[0]["amortizacao_mensal"] == 600.0
    assert out.iloc[0]["subsidio_acumulado"] != 0.0
    assert "impacto_fiscal_real" in out.columns
    assert diario.exists()

    diarios = pd.read_excel(diario)
    assert len(diarios) >= 60  # ~3 meses
    assert "taxa_selic_diaria" in diarios.columns
    assert diarios["dia_parcela"].sum() == 3
    assert set(diarios["Instituição Financeira"]) == {"BANCO TESTE"}


def test_load_selic_placeholder_sem_arquivo():
    df = load_selic(None)
    assert len(df) > 1000
    assert {"data", "fator"} <= set(df.columns)


def test_gerar_fluxos_colunas_contagil_portugues(tmp_path: Path):
    """Excel ContAgil com headers PT (R$, parênteses, hífens)."""
    root_main = _mod.main

    dados = tmp_path / "dados"
    saida = tmp_path / "saida"
    dados.mkdir()
    pd.DataFrame(
        {
            "Data da contratação": ["15/03/2009"],
            "Valor Desembolsado R$ (*)": [90000.0],
            "Juros": [6.0],
            "Prazo - Carência (meses)": [0],
            "Prazo - Amortização (meses)": [3],
            "Instituição Financeira Credenciada": ["BANCO X"],
            "Custo financeiro": ["TAXA FIXA"],
        }
    ).to_excel(dados / "ops.xlsx", index=False)
    selic = tmp_path / "STP.xlsx"
    # Fatores diários crescentes na col D → SELIC mensal > taxa do contrato
    dates = pd.date_range("2009-01-01", "2026-06-30", freq="D")
    fator = np.cumprod(np.full(len(dates), 1.0004))
    pd.DataFrame(
        {
            "data": dates,
            "b": 0,
            "c": 0,
            "d": fator,  # ContAgil: coluna D
            "e": 0,
        }
    ).to_excel(selic, index=False)

    rc = root_main(
        [
            "--massa-dados",
            str(dados),
            "--pasta-saida",
            str(saida),
            "--arquivo-selic",
            str(selic),
        ]
    )
    assert rc == 0
    out = pd.read_excel(saida / "fluxos_ops.xlsx")
    assert len(out) == 1
    assert float(out.iloc[0]["amortizacao_mensal"]) == 30000.0
    assert float(out.iloc[0]["subsidio_acumulado"]) > 0
    assert float(out.iloc[0]["impacto_fiscal_real"]) > 0
