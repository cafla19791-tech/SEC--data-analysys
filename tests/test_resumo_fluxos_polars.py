"""Testes do resumo ContAgil ultra-rápido com Polars."""

from datetime import date
from pathlib import Path

import polars as pl
import pytest

from scripts.gerar_fluxos import FATOR_30_06_2026
from scripts.resumo_fluxos_polars import (
    adicionar_spread,
    calcular_impacto_fiscal,
    carregar_fatores_selic,
    carregar_fluxos_polars,
    carregar_instituicoes,
    exportar_excel,
    main,
    montar_resumos,
    _escolher_coluna_fator,
)


def _fluxos() -> pl.DataFrame:
    return pl.DataFrame(
        {
            "contrato": [0, 0, 1],
            "data_fluxo": [date(2009, 2, 15), date(2010, 1, 15), date(2009, 3, 15)],
            "subsidio": [100.0, 50.0, 20.0],
            "saldo": [1000.0, 500.0, 200.0],
            "selic_mes": [0.01, 0.01, 0.01],
            "taxa_contrato": [0.005, 0.005, 0.005],
            "Instituição Financeira Credenciada": [
                "BANCO DO BRASIL SA",
                "BANCO DO BRASIL SA",
                "CAIXA ECONOMICA FEDERAL",
            ],
        }
    )


def _selic_stp(tmp_path: Path) -> Path:
    """STP ContAgil-like: col D = fator acumulado (>1)."""
    path = tmp_path / "STP-test.xlsx"
    pl.DataFrame(
        {
            "data": [date(2009, 2, 15), date(2009, 3, 15), date(2010, 1, 15), date(2026, 6, 30)],
            "b": [None, None, None, None],
            "c": [None, None, None, None],
            "fator_d": [1.0, 1.5, 2.5, 4.0],
            "e": [0.0, 0.0, 0.0, 0.0],
        }
    ).write_excel(path)
    return path


def test_escolher_coluna_fator_nomeada():
    raw = pl.DataFrame(
        {
            "data": [date(2009, 1, 1)],
            "col_b": [None],
            "col_c": [None],
            "taxa": [0.04],
            "fator_acumulado": [1.5],
        }
    )
    col, usar_ref = _escolher_coluna_fator(raw)
    assert col == "fator_acumulado"
    assert usar_ref is False


def test_escolher_coluna_fator_contagil_d():
    raw = pl.DataFrame(
        {
            "data": [date(2009, 1, 1), date(2009, 1, 2)],
            "b": [None, None],
            "c": [None, None],
            "fator_d": [10.0, 11.0],
            "e": [0.0, 0.0],
        }
    )
    col, usar_ref = _escolher_coluna_fator(raw)
    assert col == "fator_d"
    assert usar_ref is True


def test_carregar_fatores_selic_stp(tmp_path: Path):
    path = _selic_stp(tmp_path)
    selic, fator_final = carregar_fatores_selic(path)
    assert fator_final == pytest.approx(FATOR_30_06_2026)
    assert selic.height == 4
    assert "fator_acumulado" in selic.columns


def test_calcular_impacto_fiscal_contagil(tmp_path: Path):
    selic, fator_final = carregar_fatores_selic(_selic_stp(tmp_path))
    df = calcular_impacto_fiscal(_fluxos(), selic, fator_final)
    # 100 * FATOR / 1.0 ; 50 * FATOR / 2.5
    c0 = df.filter(pl.col("contrato") == 0).sort("data_fluxo")
    assert float(c0["impacto_acumulado_2026"][0]) == pytest.approx(
        round(100.0 * FATOR_30_06_2026 / 1.0, 2)
    )
    assert float(c0["impacto_acumulado_2026"][1]) == pytest.approx(
        round(50.0 * FATOR_30_06_2026 / 2.5, 2)
    )


def test_adicionar_spread_e_resumos():
    df = adicionar_spread(_fluxos()).with_columns(
        pl.col("data_fluxo").dt.year().alias("ano")
    )
    # força impacto para montar_resumos
    df = df.with_columns(pl.col("subsidio").alias("impacto_acumulado_2026"))
    assert "spread" in df.columns
    resumos = montar_resumos(df)
    assert set(resumos) == {"Contratos", "Por_Ano", "Por_Agente", "Impacto_Por_Ano"}
    assert resumos["Contratos"].height == 2
    assert resumos["Por_Agente"].height == 2


def test_exportar_excel(tmp_path: Path):
    import fastexcel

    df = adicionar_spread(_fluxos()).with_columns(
        [
            pl.col("data_fluxo").dt.year().alias("ano"),
            pl.col("subsidio").alias("impacto_acumulado_2026"),
        ]
    )
    out = exportar_excel(montar_resumos(df), tmp_path / "resumo_fluxos_polars.xlsx")
    assert out.exists()
    nomes = set(fastexcel.read_excel(out).sheet_names)
    assert {"Contratos", "Por_Ano", "Por_Agente", "Impacto_Por_Ano"} <= nomes


def test_carregar_instituicoes_sample():
    sample = Path("data/sample_operacoes_com_agente.csv")
    if not sample.exists():
        pytest.skip("amostra ausente")
    inst = carregar_instituicoes(sample)
    assert inst is not None
    assert "contrato" in inst.columns
    assert "Instituição Financeira Credenciada" in inst.columns


def test_cli_contagil(tmp_path: Path):
    pasta = tmp_path / "saida"
    pasta.mkdir()
    _fluxos().write_csv(pasta / "fluxos_0.csv")
    selic = _selic_stp(tmp_path)

    sample = Path("data/sample_operacoes_com_agente.csv")
    if not sample.exists():
        pytest.skip("amostra ausente")

    out = tmp_path / "out"
    rc = main(
        [
            "--pasta",
            str(pasta),
            "--original",
            str(sample),
            "--selic",
            str(selic),
            "--output-dir",
            str(out),
        ]
    )
    assert rc == 0
    assert (out / "resumo_fluxos_polars.xlsx").exists()


def test_cli_estilo_contagil_fallback(tmp_path: Path, monkeypatch):
    from scripts import resumo_fluxos_avancado as adv

    pasta = tmp_path / "saida"
    pasta.mkdir()
    _fluxos().write_csv(pasta / "fluxos_0.csv")

    # Caminhos ContAgil WinPython ausentes → pasta local + amostra + Bacen
    monkeypatch.setattr(adv, "OUTPUT_DIR", pasta)

    if not Path("data/sample_operacoes_com_agente.csv").exists():
        pytest.skip("amostra ausente")
    if not Path("data/selic_fatores_bacen.xlsx").exists():
        pytest.skip("cache SELIC ausente")

    out = tmp_path / "out"
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
    assert rc == 0
    assert (out / "resumo_fluxos_polars.xlsx").exists()
