"""Testes do resumo ContAgil Polars com SELIC/TJLP mensais."""

from datetime import date
from pathlib import Path

import polars as pl
import pytest

from scripts.resumo_fluxos_polars import (
    RELATORIO_NAME,
    WORKBOOK_NAME,
    adicionar_spread,
    adicionar_taxa_contrato_efetiva,
    anexar_selic_mensal,
    anexar_tjlp_mensal,
    calcular_impacto_fiscal,
    carregar_instituicoes,
    carregar_selic_mensal,
    carregar_tjlp_mensal,
    exportar_excel,
    fator_final_mensal,
    gerar_graficos,
    gerar_relatorio_executivo,
    main,
    montar_resumos,
    montar_totais,
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


def _selic_mensal_xlsx(tmp_path: Path) -> Path:
    path = tmp_path / "selic_mensal.xlsx"
    pl.DataFrame(
        {
            "Data": ["01/01/09", "01/02/09", "01/03/09", "01/01/10", "01/06/26"],
            "Taxa Selic mensal - % a.m.": [1.0, 1.0, 1.0, 0.8, 0.9],
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
    selic = carregar_selic_mensal(_selic_mensal_xlsx(tmp_path))
    assert "selic_mensal" in selic.columns
    assert "fator_acumulado" in selic.columns
    # 1% a.m. → 0.01 decimal
    assert float(selic["selic_mensal"][0]) == pytest.approx(0.01)
    # um ponto por mês
    assert selic.height == 5
    # cumprod: 1.01, 1.01^2, ...
    assert float(selic["fator_acumulado"][0]) == pytest.approx(1.01)
    assert float(selic["fator_acumulado"][1]) == pytest.approx(1.01 * 1.01)


def test_carregar_tjlp_e_anexar(tmp_path: Path):
    tjlp = carregar_tjlp_mensal(_tjlp_xlsx(tmp_path))
    assert float(tjlp["tjlp_mensal"][0]) == pytest.approx(0.005)
    assert "fator_acumulado" in tjlp.columns

    df = anexar_tjlp_mensal(_fluxos(), tjlp)
    assert "tjlp_mensal" in df.columns
    assert df["tjlp_mensal"].null_count() == 0

    work = df.with_columns(
        [
            pl.lit("TJLP").alias("encargo_financeiro"),
            pl.lit(0.002).alias("taxa_contrato"),
        ]
    )
    out = adicionar_taxa_contrato_efetiva(work)
    esperado = (1.0 + 0.005) * (1.0 + 0.002) - 1.0
    assert float(out["taxa_contrato_efetiva"][0]) == pytest.approx(esperado)


def test_impacto_capitalizacao_mensal(tmp_path: Path):
    selic = carregar_selic_mensal(_selic_mensal_xlsx(tmp_path))
    fator_final = fator_final_mensal(selic)
    df = anexar_selic_mensal(_fluxos(), selic)
    out = calcular_impacto_fiscal(df, selic, fator_final)

    # Parcela em fev/2009: fator = 1.01^2 (jan+fev)
    c0 = out.filter(pl.col("contrato") == 0).sort("data_fluxo")
    fator_fev = float(
        selic.filter(pl.col("data") == date(2009, 2, 1))["fator_acumulado"][0]
    )
    esperado = round(100.0 * fator_final / fator_fev, 2)
    assert float(c0["impacto_acumulado_2026"][0]) == pytest.approx(esperado)


def test_taxa_contrato_efetiva_tjlp_base():
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
    assert float(out["taxa_contrato_efetiva"][0]) > float(out["taxa_contrato_efetiva"][1])


def test_adicionar_spread_e_resumos(tmp_path: Path):
    selic = carregar_selic_mensal(_selic_mensal_xlsx(tmp_path))
    df = anexar_selic_mensal(_fluxos(), selic)
    df = adicionar_spread(df).with_columns(pl.col("data_fluxo").dt.year().alias("ano"))
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
    assert "Total Subsídio" in resumos["Totais_Gerais"]["Métrica"].to_list()


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
    assert "mensal" in texto.lower() or "selic_mensal" in texto.lower()

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


def test_cli_contagil_mensal(tmp_path: Path):
    pasta = tmp_path / "saida"
    pasta.mkdir()
    _fluxos().write_csv(pasta / "fluxos_0.csv")
    selic = _selic_mensal_xlsx(tmp_path)
    tjlp = _tjlp_xlsx(tmp_path)

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


def test_cli_estilo_contagil_paths(tmp_path: Path, monkeypatch):
    from scripts import resumo_fluxos_avancado as adv

    pasta = tmp_path / "saida"
    pasta.mkdir()
    _fluxos().write_csv(pasta / "fluxos_0.csv")
    monkeypatch.setattr(adv, "OUTPUT_DIR", pasta)

    # Amostras locais do repo
    if not Path("data/sample_operacoes_com_agente.csv").exists():
        pytest.skip("amostra ausente")
    if not Path("data/selic_mensal.xlsx").exists():
        pytest.skip("selic_mensal amostra ausente")
    if not Path("data/tjlp_mensal.xlsx").exists():
        pytest.skip("tjlp_mensal amostra ausente")

    out = tmp_path / "out"
    rc = main(
        [
            "--pasta",
            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida",
            "--original",
            "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx",
            "--selic",
            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\selic_mensal.xlsx",
            "--tjlp",
            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\tjlp_mensal.xlsx",
            "--output-dir",
            str(out),
            "--sem-graficos",
        ]
    )
    assert rc == 0
    assert (out / WORKBOOK_NAME).exists()
