"""Testes do entrypoint ContAgil SAC (gerar_fluxo_sac.py)."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from scripts.gerar_fluxo_sac import (
    VALOR_CONTRATO0_PADRAO,
    _aplicar_correcao_contrato0,
    _resolver_pasta_saida,
    main as sac_main,
    parse_args,
)


def test_parse_corrigir_contrato0_default():
    args = parse_args(["--corrigir-contrato0", "--input", "x.csv"])
    assert args.corrigir_contrato0 == VALOR_CONTRATO0_PADRAO


def test_aplicar_correcao_contrato0():
    df = pd.DataFrame(
        {
            "contrato": [0, 1],
            "valor_desembolsado": [9400.0, 50000.0],
        }
    )
    out = _aplicar_correcao_contrato0(df, 485000.0)
    assert out.loc[0, "valor_desembolsado"] == 485000.0
    assert out.loc[1, "valor_desembolsado"] == 50000.0


def test_resolver_pasta_saida_contagil_ausente(tmp_path: Path, monkeypatch):
    import scripts.gerar_fluxo_sac as mod

    monkeypatch.setattr(mod, "LOCAL_SAIDA", tmp_path / "saida_local")
    monkeypatch.setattr(mod, "OUTPUT_DIR", tmp_path / "output")
    saida = _resolver_pasta_saida(
        Path(r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida")
    )
    assert saida == tmp_path / "saida_local"


def test_main_smoke_sample(tmp_path: Path):
    saida = tmp_path / "saida"
    sample = Path(__file__).resolve().parents[1] / "data" / "sample_operacoes_com_agente.csv"
    selic = tmp_path / "STP.xlsx"
    dates = pd.date_range("2009-01-01", "2026-06-30", freq="D")
    fator = np.cumprod(np.full(len(dates), 1.0001))
    pd.DataFrame(
        {"data": dates, "b": 0, "c": 0, "d": 0.01, "fator": fator}
    ).to_excel(selic, index=False)

    rc = sac_main(
        [
            "--input",
            str(sample),
            "--pasta-saida",
            str(saida),
            "--arquivo-selic",
            str(selic),
            "--corrigir-contrato0",
            "--stem",
            "fluxos_completos_corrigido",
        ]
    )
    assert rc == 0
    assert (saida / "fluxos_completos_corrigido.csv").exists()
    assert (saida / "resumo_por_agente.xlsx").exists()
    fluxos = pd.read_csv(saida / "fluxos_completos_corrigido.csv")
    assert len(fluxos) > 0
    assert "saldo_fiscal" in fluxos.columns
    assert "impacto_fiscal" in fluxos.columns
