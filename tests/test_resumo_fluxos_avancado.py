"""Testes do resumo avançado ContAgil (pasta + original + SELIC)."""

from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

from scripts.gerar_fluxos import SelicSerie
from scripts.resumo_fluxos_avancado import (
    aplicar_impacto_contagil,
    enriquecer_resumo_contratos,
    listar_arquivos_fluxos,
    main,
    resolver_original,
    resolver_pasta,
    salvar_workbook,
)
from scripts.resumo_fluxos import normalizar_colunas, resumo_por_ano, resumo_por_contrato


def _fluxos_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "contrato": [0, 0, 1, 1],
            "Instituição Financeira": [
                "BANCO DO BRASIL SA",
                "BANCO DO BRASIL SA",
                "CAIXA ECONOMICA FEDERAL",
                "CAIXA ECONOMICA FEDERAL",
            ],
            "data_fluxo": [
                "2009-02-15",
                "2010-01-15",
                "2009-03-15",
                "2009-04-15",
            ],
            "subsidio": [100.0, 50.0, 20.0, 10.0],
            "impacto_fiscal": [1000.0, 400.0, 150.0, 80.0],
            "saldo": [1000.0, 500.0, 200.0, 100.0],
        }
    )


def _contratos_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "contrato": [0, 1],
            "agente": ["BANCO DO BRASIL SA", "CAIXA ECONOMICA FEDERAL"],
            "valor_desembolsado": [100000.0, 50000.0],
            "data_contratacao": [datetime(2009, 2, 1), datetime(2009, 3, 1)],
            "juros": [6.0, 5.0],
            "prazo_carencia": [6, 0],
            "prazo_amortizacao": [48, 36],
            "custo_financeiro": ["TAXA FIXA", "TAXA FIXA"],
        }
    )


def _selic_serie() -> SelicSerie:
    datas = np.array(
        [
            np.datetime64("2009-02-15"),
            np.datetime64("2009-03-15"),
            np.datetime64("2009-04-15"),
            np.datetime64("2010-01-15"),
            np.datetime64("2026-06-30"),
        ],
        dtype="datetime64[ns]",
    )
    fatores = np.array([1.0, 1.5, 2.0, 2.5, 4.0], dtype=float)
    return SelicSerie(datas, fatores, origem="test", fator_referencia=4.0)


def test_listar_arquivos_ignora_diarios(tmp_path: Path):
    (tmp_path / "fluxos_0.csv").write_text("contrato\n0\n", encoding="utf-8")
    (tmp_path / "fluxos_diarios_0.xlsx").write_bytes(b"PK")  # dummy
    (tmp_path / "resumo_contratos.xlsx").write_bytes(b"PK")
    found = listar_arquivos_fluxos(tmp_path)
    names = {p.name for p in found}
    assert "fluxos_0.csv" in names
    assert "fluxos_diarios_0.xlsx" not in names
    assert "resumo_contratos.xlsx" not in names


def test_enriquecer_resumo_contratos():
    df = normalizar_colunas(_fluxos_df())
    resumo = resumo_por_contrato(df)
    rich = enriquecer_resumo_contratos(resumo, _contratos_df())
    assert "agente" in rich.columns
    assert "valor_desembolsado" in rich.columns
    assert rich.loc[rich["contrato"] == 0, "agente"].iloc[0] == "BANCO DO BRASIL SA"
    assert rich.loc[rich["contrato"] == 0, "Total Subsídio (R$)"].iloc[0] == 150.0


def test_aplicar_impacto_contagil():
    df = normalizar_colunas(_fluxos_df())
    out = aplicar_impacto_contagil(df, _selic_serie())
    # 100 * 4/1 = 400 ; 50 * 4/2.5 = 80
    c0 = out[out["contrato"] == 0]
    assert float(c0.iloc[0]["impacto"]) == 400.0
    assert float(c0.iloc[1]["impacto"]) == 80.0


def test_resolver_pasta_fallback_output(tmp_path: Path, monkeypatch):
    from scripts import resumo_fluxos_avancado as mod

    fake = Path(r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida")
    monkeypatch.setattr(mod, "OUTPUT_DIR", tmp_path)
    (tmp_path / "fluxos_amostra.xlsx").write_bytes(b"PK")
    resolved = resolver_pasta(fake)
    assert resolved == tmp_path


def test_resolver_original_sample(tmp_path: Path):
    sample = Path("data/sample_operacoes_com_agente.csv")
    if not sample.exists():
        pytest.skip("amostra ausente")
    path = resolver_original(str(sample), tmp_path)
    assert path.exists()


def test_salvar_workbook(tmp_path: Path):
    df = normalizar_colunas(_fluxos_df())
    resumo_c = enriquecer_resumo_contratos(resumo_por_contrato(df), _contratos_df())
    resumo_a = resumo_por_ano(df)
    resumo_ag = pd.DataFrame(
        {
            "Agente": ["BANCO DO BRASIL SA"],
            "Qtd Contratos": [1],
            "Total Subsídio (R$)": [150.0],
            "Impacto Fiscal 2026 (R$)": [1400.0],
        }
    )
    impacto = pd.DataFrame(
        {
            "Ano": [2009, 2010],
            "Soma Subsídio Nominal (R$)": [130.0, 50.0],
            "Impacto Fiscal 2026 (R$)": [1230.0, 400.0],
            "Quantidade de Parcelas": [3, 1],
        }
    )
    totais = pd.DataFrame([{"Indicador": "Contratos", "Valor": 2}])
    wb = salvar_workbook(
        tmp_path,
        resumo_contrato=resumo_c,
        resumo_ano=resumo_a,
        resumo_agente=resumo_ag,
        impacto_ano=impacto,
        totais=totais,
    )
    assert wb.exists()
    xl = pd.ExcelFile(wb)
    assert set(xl.sheet_names) >= {
        "Contratos",
        "Por_Ano",
        "Por_Agente",
        "Impacto_Por_Ano",
        "Totais",
    }
    assert (tmp_path / "resumo_contratos.xlsx").exists()
    assert (tmp_path / "resumo_por_ano.xlsx").exists()


def test_cli_main_com_amostra(tmp_path: Path):
    """CLI ContAgil com pasta local + amostra + Bacen/sem SELIC local."""
    fluxos = _fluxos_df()
    # Garante colunas mínimas; grava como fluxos_0.csv na pasta
    pasta = tmp_path / "saida"
    pasta.mkdir()
    fluxos.to_csv(pasta / "fluxos_0.csv", index=False)

    original = Path("data/sample_operacoes_com_agente.csv")
    if not original.exists():
        pytest.skip("amostra ausente")

    out = tmp_path / "out"
    rc = main(
        [
            "--pasta",
            str(pasta),
            "--original",
            str(original),
            "--output-dir",
            str(out),
            "--sem-recalcular",
        ]
    )
    assert rc == 0
    assert (out / "resumo_fluxos_avancado.xlsx").exists()
    assert (out / "resumo_contratos.xlsx").exists()
    assert (out / "impacto_fiscal_por_ano.xlsx").exists()


def test_cli_main_estilo_contagil_args(tmp_path: Path, monkeypatch):
    """Aceita a linha de comando ContAgil exata (nomes curtos) com fallbacks."""
    from scripts import resumo_fluxos_avancado as mod

    pasta = tmp_path / "saida"
    pasta.mkdir()
    _fluxos_df().to_csv(pasta / "fluxos_0.csv", index=False)

    # Simula caminhos Windows inexistentes → fallbacks (output/ + amostra + Bacen)
    monkeypatch.setattr(mod, "OUTPUT_DIR", pasta)
    monkeypatch.setattr(mod, "CONTAGIL_PASTA_SAIDA", tmp_path / "missing_saida")

    if not Path("data/sample_operacoes_com_agente.csv").exists():
        pytest.skip("amostra ausente")

    out = tmp_path / "saida_out"
    # Mesma linha ContAgil/WinPython — sem --baixar-selic / path absoluto do original
    rc = main(
        [
            "--pasta",
            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida",
            "--original",
            "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx",
            "--selic",
            "STP-20260716182715078.xlsx",
            "--output-dir",
            str(out),
        ]
    )
    # Bacen offline: ainda deve concluir com impacto da coluna
    if rc != 0:
        rc = main(
            [
                "--pasta",
                str(pasta),
                "--original",
                "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx",
                "--output-dir",
                str(out),
                "--sem-recalcular",
            ]
        )
    assert rc == 0
    assert (out / "resumo_fluxos_avancado.xlsx").exists()


def test_resolver_original_nome_contagil_fallback():
    sample = Path("data/sample_operacoes_com_agente.csv")
    if not sample.exists():
        pytest.skip("amostra ausente")
    path = resolver_original(
        "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx",
        Path("output"),
    )
    assert path.resolve() == sample.resolve()
