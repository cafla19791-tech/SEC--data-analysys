"""Testes do discriminante de lucros líquidos dos bancos."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from discriminante_lucros_bancos import (  # noqa: E402
    DADOS,
    FX_MEDIO,
    NUBANK_USD,
    build_frame,
    write_excel,
)


def test_cinco_bancos_e_anos():
    df = build_frame()
    assert list(df.index) == DADOS["Banco"]
    for col in ("2023", "2024", "2025", "2026_1S"):
        assert col in df.columns
        assert (df[col] > 0).all()


def test_itau_lider_em_todos_periodos():
    df = build_frame()
    for col in ("2023", "2024", "2025", "2026_1S"):
        assert df[col].idxmax() == "Itaú Unibanco"


def test_nubank_conversao_fx():
    """Conversão US$ → R$ deve bater com a PTAX média usada."""
    idx = DADOS["Banco"].index("Nubank (Nu Holdings)")
    for periodo in ("2023", "2024", "2025", "2026_1S"):
        esperado = round(NUBANK_USD[periodo] * FX_MEDIO[periodo], 2)
        obtido = round(DADOS[periodo][idx], 2)
        assert abs(esperado - obtido) < 0.02, (periodo, esperado, obtido)


def test_totais_crescem_2023_a_2025():
    df = build_frame()
    t23 = df["2023"].sum()
    t24 = df["2024"].sum()
    t25 = df["2025"].sum()
    assert t23 < t24 < t25


def test_excel_gerado(tmp_path):
    df = build_frame()
    path = tmp_path / "disc.xlsx"
    write_excel(df, path)
    assert path.exists() and path.stat().st_size > 1000
    sheets = pd.ExcelFile(path).sheet_names
    assert "Discriminante" in sheets
    assert "Nubank_USD" in sheets
    assert "Ranking" in sheets
