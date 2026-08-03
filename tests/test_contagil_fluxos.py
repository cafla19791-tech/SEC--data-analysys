"""Testes do entrypoint ContAgil (massa_dados → fluxos_*.xlsx)."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from scripts.contagil_fluxos import (
    listar_excels,
    main as contagil_main,
    parse_args,
    processar_arquivo,
    processar_pasta_dados,
)
from scripts.gerar_fluxos import SelicSerie


def _serie_sintetica() -> SelicSerie:
    datas = np.array(
        [
            np.datetime64("2009-01-01"),
            np.datetime64("2009-02-16"),
            np.datetime64("2026-06-30"),
        ],
        dtype="datetime64[ns]",
    )
    return SelicSerie(datas, np.array([1.0, 1.5, 3.0]))


def _excel_contratos(path: Path) -> None:
    """Excel no layout ContAgil (header na 1ª linha, colunas PT)."""
    pd.DataFrame(
        {
            "Data da contratação": ["15/03/2009", "20/04/2009"],
            "Valor Desembolsado R$ (*)": [100000.0, 50000.0],
            "Juros": [6.0, 2.0],
            "Prazo - Carência (meses)": [0, 0],
            "Prazo - Amortização (meses)": [3, 2],
            "Instituição Financeira Credenciada": ["BANCO A", "BANCO B"],
            "Custo financeiro": ["TAXA FIXA", "TJLP"],
        }
    ).to_excel(path, index=False)


def test_listar_excels(tmp_path: Path):
    (tmp_path / "a.xlsx").write_bytes(b"dummy")
    (tmp_path / "b.txt").write_text("x")
    assert [p.name for p in listar_excels(tmp_path)] == ["a.xlsx"]


def test_processar_arquivo_grava_fluxos(tmp_path: Path):
    src = tmp_path / "operacoes.xlsx"
    saida = tmp_path / "saida"
    _excel_contratos(src)

    out = processar_arquivo(src, saida, _serie_sintetica(), header=0)
    assert out.name == "fluxos_operacoes.xlsx"
    assert out.exists()
    df = pd.read_excel(out)
    assert len(df) == 5  # 3 + 2 parcelas
    assert "impacto_fiscal" in df.columns
    assert "Instituição Financeira" in df.columns


def test_processar_pasta_dados_ignora_stp(tmp_path: Path):
    dados = tmp_path / "dados"
    saida = tmp_path / "saida"
    dados.mkdir()
    _excel_contratos(dados / "lote1.xlsx")
    # STP na mesma pasta não deve gerar fluxos
    pd.DataFrame(
        {
            "data": ["01/01/2009", "30/06/2026"],
            "b": [0, 0],
            "c": [0, 0],
            "d": [0.01, 0.01],
            "fator": [1.0, 2.0],
        }
    ).to_excel(dados / "STP-20260716182715078 (1).xlsx", index=False)

    outs = processar_pasta_dados(dados, saida, _serie_sintetica(), header=0)
    assert len(outs) == 1
    assert outs[0].name == "fluxos_lote1.xlsx"


def test_parse_args_massa_dados_alias():
    """CLI ContAgil: --massa-dados é alias de --pasta-dados."""
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
    assert args.pasta_dados == Path(
        r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\dados"
    )
    assert args.pasta_saida == Path(
        r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida"
    )
    assert args.arquivo_selic == Path(
        r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\STP-20260716182715078 (1).xlsx"
    )


def test_main_massa_dados_cli(tmp_path: Path):
    """Smoke: comando ContAgil com --massa-dados + STP local."""
    dados = tmp_path / "dados"
    saida = tmp_path / "saida"
    dados.mkdir()
    _excel_contratos(dados / "lote_demo.xlsx")
    selic = tmp_path / "STP-20260716182715078 (1).xlsx"
    pd.DataFrame(
        {
            "data": ["01/01/2009", "16/02/2009", "30/06/2026"],
            "b": [0, 0, 0],
            "c": [0, 0, 0],
            "d": [0.01, 0.01, 0.01],
            "fator": [1.0, 1.5, 3.0],
        }
    ).to_excel(selic, index=False)

    rc = contagil_main(
        [
            "--massa-dados",
            str(dados),
            "--pasta-saida",
            str(saida),
            "--arquivo-selic",
            str(selic),
            "--excel-header",
            "0",
        ]
    )
    assert rc == 0
    out = saida / "fluxos_lote_demo.xlsx"
    assert out.exists()
    df = pd.read_excel(out)
    assert len(df) == 5
    assert "impacto_fiscal" in df.columns


def test_massa_dados_inexistente_erro_claro(tmp_path: Path):
    missing = tmp_path / "dados_inexistente"
    try:
        processar_pasta_dados(missing, tmp_path / "saida", _serie_sintetica())
        assert False, "deveria falhar"
    except FileNotFoundError as exc:
        assert "Massa de dados não encontrada" in str(exc)


def test_main_massa_dados_com_fluxo_diario(tmp_path: Path):
    dados = tmp_path / "dados"
    saida = tmp_path / "saida"
    dados.mkdir()
    _excel_contratos(dados / "lote_diario.xlsx")
    selic = tmp_path / "STP-demo.xlsx"
    pd.DataFrame(
        {
            "data": ["01/01/2009", "16/02/2009", "30/06/2026"],
            "b": [0, 0, 0],
            "c": [0, 0, 0],
            "d": [0.01, 0.01, 0.01],
            "fator": [1.0, 1.5, 3.0],
        }
    ).to_excel(selic, index=False)

    rc = contagil_main(
        [
            "--massa-dados",
            str(dados),
            "--pasta-saida",
            str(saida),
            "--arquivo-selic",
            str(selic),
            "--excel-header",
            "0",
            "--fluxo-diario",
        ]
    )
    assert rc == 0
    assert (saida / "fluxos_lote_diario.xlsx").exists()
    assert (saida / "fluxos_diarios_lote_diario.xlsx").exists()
