"""Testes unitários dos parsers e da capitalização Selic do gerador de fluxos."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from gerar_fluxos_contratos import (  # noqa: E402
    SelicSerie,
    limpar_valor,
    parse_datas,
    taxa_mensal_from_row,
)


def test_limpar_valor_br_e_us():
    s = pd.Series(["R$ 1.234,56", "5,5", "5.0", "44000.0", "250000,50", "1.234.567,89"])
    assert limpar_valor(s).tolist() == [1234.56, 5.5, 5.0, 44000.0, 250000.50, 1234567.89]


def test_parse_datas_iso_e_br():
    d = pd.Series(["2009-01-02T00:00:00", "15/03/2009", "02/01/2009", "2009-12-31"])
    parsed = parse_datas(d)
    assert [str(x.date()) for x in parsed] == [
        "2009-01-02",
        "2009-03-15",
        "2009-01-02",
        "2009-12-31",
    ]


def test_taxa_mensal_fixa_e_tjlp():
    assert abs(taxa_mensal_from_row("TAXA FIXA", 5.5) - ((1.055) ** (1 / 12) - 1)) < 1e-12
    assert abs(taxa_mensal_from_row("TJLP", 2.0) - ((1.08) ** (1 / 12) - 1)) < 1e-12
    assert abs(taxa_mensal_from_row("TLP + TAXA FIXA", 1.5) - ((1.075) ** (1 / 12) - 1)) < 1e-12


def test_selic_fator_rapido_e_capitalizar():
    datas = np.array(
        ["2009-01-01", "2009-02-01", "2010-01-01", "2026-06-30"],
        dtype="datetime64[ns]",
    )
    fatores = np.array([1.0, 1.01, 1.15, 8.0])
    serie = SelicSerie(datas, fatores)

    f = serie.fator_rapido([np.datetime64("2009-01-15", "ns"), np.datetime64("2009-02-15", "ns")])
    assert list(f) == [1.0, 1.01]

    impacto = serie.capitalizar(
        np.array([100.0]),
        [np.datetime64("2009-02-01", "ns")],
        datetime(2026, 6, 30),
    )
    assert abs(impacto[0] - 100.0 * (8.0 / 1.01)) < 1e-9
