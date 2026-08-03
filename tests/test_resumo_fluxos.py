"""Testes do resumo de fluxos por contrato e por ano."""

from pathlib import Path

import pandas as pd
import pytest

from scripts.resumo_fluxos import (
    carregar_fluxos,
    normalizar_colunas,
    resolver_fluxos,
    resumo_por_ano,
    resumo_por_contrato,
    salvar_resumos,
)


def _df_base() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "contrato": [0, 0, 0, 1, 1],
            "data_fluxo": [
                "2009-02-15",
                "2010-01-15",
                "2010-02-15",
                "2009-03-15",
                "2009-04-15",
            ],
            "subsidio": [100.0, 50.0, 25.0, 10.0, 5.0],
            "impacto": [1000.0, 400.0, 200.0, 80.0, 40.0],
            "saldo": [1000.0, 500.0, 250.0, 200.0, 100.0],
        }
    )


def test_resumo_por_contrato_soma_e_saldo_final():
    df = normalizar_colunas(_df_base())
    resumo = resumo_por_contrato(df)

    assert list(resumo.columns) == [
        "Total Subsídio (R$)",
        "Impacto Fiscal 2026 (R$)",
        "Saldo Final (R$)",
        "Quantidade de Parcelas",
    ]
    assert resumo.loc[0, "Total Subsídio (R$)"] == 175.0
    assert resumo.loc[0, "Impacto Fiscal 2026 (R$)"] == 1600.0
    assert resumo.loc[0, "Saldo Final (R$)"] == 250.0
    assert int(resumo.loc[0, "Quantidade de Parcelas"]) == 3
    assert resumo.loc[1, "Saldo Final (R$)"] == 100.0


def test_resumo_por_ano_por_contrato():
    df = normalizar_colunas(_df_base())
    resumo = resumo_por_ano(df)

    assert resumo.loc[(0, 2009), "Total Subsídio (R$)"] == 100.0
    assert resumo.loc[(0, 2010), "Total Subsídio (R$)"] == 75.0
    assert resumo.loc[(0, 2010), "Impacto Fiscal 2026 (R$)"] == 600.0
    assert resumo.loc[(1, 2009), "Total Subsídio (R$)"] == 15.0


def test_alias_impacto_fiscal_e_saldo_fiscal():
    df = pd.DataFrame(
        {
            "contrato": [0, 0],
            "data_fluxo": ["2009-02-15", "2009-03-15"],
            "subsidio": [10.0, 20.0],
            "impacto_fiscal": [100.0, 200.0],
            "saldo_fiscal": [500.0, 400.0],
        }
    )
    work = normalizar_colunas(df)
    resumo = resumo_por_contrato(work)
    assert resumo.loc[0, "Impacto Fiscal 2026 (R$)"] == 300.0
    assert resumo.loc[0, "Saldo Final (R$)"] == 400.0


def test_saldo_final_respeita_ordem_cronologica():
    """'last' deve ser o saldo na última data_fluxo, não a ordem do arquivo."""
    df = pd.DataFrame(
        {
            "contrato": [0, 0],
            "data_fluxo": ["2010-01-15", "2009-01-15"],  # fora de ordem
            "subsidio": [1.0, 1.0],
            "impacto": [1.0, 1.0],
            "saldo": [100.0, 999.0],
        }
    )
    resumo = resumo_por_contrato(normalizar_colunas(df))
    assert resumo.loc[0, "Saldo Final (R$)"] == 100.0


def test_normalizar_sem_impacto_falha():
    df = pd.DataFrame(
        {
            "contrato": [0],
            "data_fluxo": ["2009-01-15"],
            "subsidio": [1.0],
        }
    )
    with pytest.raises(ValueError, match="impacto"):
        normalizar_colunas(df)


def test_salvar_resumos(tmp_path: Path):
    df = normalizar_colunas(_df_base())
    rc = resumo_por_contrato(df)
    ra = resumo_por_ano(df)
    x_c, x_a = salvar_resumos(rc, ra, tmp_path)
    assert x_c.exists()
    assert x_a.exists()
    assert x_c.with_suffix(".csv").exists()
    assert x_a.with_suffix(".csv").exists()
    lido = pd.read_excel(x_c, index_col=0)
    assert lido.loc[0, "Total Subsídio (R$)"] == 175.0


def test_resolver_fluxos_explicito(tmp_path: Path):
    path = tmp_path / "fluxos_0.csv"
    _df_base().to_csv(path, index=False)
    assert resolver_fluxos(path) == path
    loaded = carregar_fluxos(path)
    assert "contrato" in loaded.columns


def test_cli_main(tmp_path: Path):
    from scripts.resumo_fluxos import main

    path = tmp_path / "fluxos_0.csv"
    _df_base().to_csv(path, index=False)
    out = tmp_path / "saida"
    rc = main(["--fluxos", str(path), "--output-dir", str(out), "--contrato", "0"])
    assert rc == 0
    assert (out / "resumo_contratos.xlsx").exists()
    assert (out / "resumo_por_ano.xlsx").exists()
