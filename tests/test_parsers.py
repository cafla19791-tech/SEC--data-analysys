"""Testes unitários dos parsers do gerador de fluxos."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from gerar_fluxos_contratos import limpar_valor, parse_datas, taxa_mensal_from_row  # noqa: E402


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
