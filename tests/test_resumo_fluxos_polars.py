"""Testes do resumo ContAgil ultra-rápido com Polars."""

from datetime import date
from pathlib import Path

import polars as pl
import pytest

from scripts.gerar_fluxos import FATOR_30_06_2026
from scripts.resumo_fluxos_polars import (
    RELATORIO_NAME,
    WORKBOOK_NAME,
    adicionar_spread,
    adicionar_taxa_contrato_efetiva,
    anexar_tjlp_mensal,
    calcular_impacto_fiscal,
    carregar_fatores_selic,
    carregar_instituicoes,
    carregar_selic_auto,
    carregar_selic_mensal,
    carregar_tjlp_mensal,
    exportar_excel,
    gerar_graficos,
    gerar_relatorio_executivo,
    main,
    montar_resumos,
    montar_totais,
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
    assert "taxa_contrato_efetiva" in df.columns
    resumos = montar_resumos(df)
    assert set(resumos) == {
        "Contratos",
        "Por_Ano",
        "Por_Agente",
        "Impacto_Por_Ano",
        "Totais_Gerais",
    }
    assert resumos["Contratos"].height == 2
    assert resumos["Por_Agente"].height == 2
    assert "Total Subsídio" in resumos["Totais_Gerais"]["Métrica"].to_list()


def test_taxa_contrato_efetiva_tjlp():
    df = pl.DataFrame(
        {
            "contrato": [0, 1],
            "data_fluxo": [date(2009, 2, 15), date(2009, 3, 15)],
            "subsidio": [1.0, 1.0],
            "saldo": [1.0, 1.0],
            "juros": [6.0, 6.0],
            "encargo_financeiro": ["TJLP", "TAXA FIXA"],
        }
    )
    out = adicionar_taxa_contrato_efetiva(df)
    # TJLP mensal > TAXA FIXA mensal para mesmo juros
    assert float(out["taxa_contrato_efetiva"][0]) > float(out["taxa_contrato_efetiva"][1])


def _selic_taxas_xlsx(tmp_path: Path) -> Path:
    path = tmp_path / "selic_bacen.xlsx"
    pl.DataFrame(
        {
            "Data": ["01/01/09", "02/01/09", "03/01/09", "15/02/09"],
            "11 - Taxa de juros - Selic - % a.d.": [0.04, 0.04, 0.04, 0.04],
        }
    ).write_excel(path)
    return path


def _tjlp_xlsx(tmp_path: Path) -> Path:
    path = tmp_path / "tjlp_mensal.xlsx"
    pl.DataFrame(
        {
            "Data": ["01/01/09", "01/02/09", "01/03/09"],
            " Taxas de juros - TJLP mensal - % a.m.": [0.5, 0.5, 0.6],
        }
    ).write_excel(path)
    return path


def test_carregar_selic_mensal_contagil(tmp_path: Path):
    path = _selic_taxas_xlsx(tmp_path)
    selic, fator_final, modo = carregar_selic_auto(path)
    assert modo == "taxas"
    assert float(selic["selic_taxa"][0]) == pytest.approx(0.0004)
    assert fator_final == pytest.approx(float(selic["fator_acumulado"].max()))


def test_carregar_tjlp_e_anexar(tmp_path: Path):
    tjlp = carregar_tjlp_mensal(_tjlp_xlsx(tmp_path))
    assert float(tjlp["tjlp_mensal"][0]) == pytest.approx(0.005)
    df = anexar_tjlp_mensal(_fluxos(), tjlp)
    assert "tjlp_mensal" in df.columns
    assert df["tjlp_mensal"].null_count() == 0

    # Com série TJLP + encargo: efetiva ContAgil (1+tjlp)*(1+spread)-1
    work = df.with_columns(
        [
            pl.lit("TJLP").alias("encargo_financeiro"),
            pl.lit(0.002).alias("taxa_contrato"),
        ]
    )
    out = adicionar_taxa_contrato_efetiva(work)
    esperado = (1.0 + 0.005) * (1.0 + 0.002) - 1.0
    assert float(out["taxa_contrato_efetiva"][0]) == pytest.approx(esperado)


def test_exportar_excel(tmp_path: Path):
    import fastexcel

    df = adicionar_spread(_fluxos()).with_columns(
        [
            pl.col("data_fluxo").dt.year().alias("ano"),
            pl.col("subsidio").alias("impacto_acumulado_2026"),
        ]
    )
    out = exportar_excel(montar_resumos(df), tmp_path / WORKBOOK_NAME)
    assert out.exists()
    nomes = set(fastexcel.read_excel(out).sheet_names)
    assert {
        "Contratos",
        "Por_Ano",
        "Por_Agente",
        "Impacto_Por_Ano",
        "Totais_Gerais",
    } <= nomes


def test_relatorio_e_graficos(tmp_path: Path):
    df = adicionar_spread(_fluxos()).with_columns(
        [
            pl.col("data_fluxo").dt.year().alias("ano"),
            pl.col("subsidio").alias("impacto_acumulado_2026"),
        ]
    )
    rel = gerar_relatorio_executivo(df, tmp_path)
    assert rel.exists()
    texto = rel.read_text(encoding="utf-8")
    assert "Relatório Executivo" in texto
    assert "Total de Contratos" in texto

    png, html = gerar_graficos(df, tmp_path)
    assert png.exists()
    assert html.exists()
    assert montar_totais(df).height == 5


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
    tjlp = _tjlp_xlsx(tmp_path)
    rc = main(
        [
            "--pasta",
            str(pasta),
            "--original",
            str(sample),
            "--selic",
            str(selic),
            "--tjlp",
            str(tjlp),
            "--output-dir",
            str(out),
        ]
    )
    assert rc == 0
    assert (out / WORKBOOK_NAME).exists()
    assert (out / RELATORIO_NAME).exists()
    assert (out / "grafico_interativo.html").exists()


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
    assert (out / WORKBOOK_NAME).exists()
    assert (out / RELATORIO_NAME).exists()
