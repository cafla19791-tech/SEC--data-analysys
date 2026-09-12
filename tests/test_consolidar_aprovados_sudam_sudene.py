"""Testes do consolidador de aprovados Sudam."""

from __future__ import annotations

from pathlib import Path

from pypdf import PdfWriter
from reportlab.pdfgen import canvas

from scripts.consolidar_aprovados_sudam_sudene import MARKER, parse_sudam_pdf


def _mini_pdf(path: Path, lines: list[str]) -> None:
    """PDF simples com texto extraível."""
    c = canvas.Canvas(str(path))
    y = 800
    for line in lines:
        c.drawString(30, y, line[:110])
        y -= 14
    c.save()


def test_marker():
    assert "sudam" in MARKER


def test_parse_sudam_pdf(tmp_path: Path):
    pdf = tmp_path / "sudam_2015.pdf"
    _mini_pdf(
        pdf,
        [
            "EMPRESA CNPJ/MF MUNICÍPIO UF SETOR PRODUTO/SERVIÇO ENQUADRAMENTO PLEITO MODALIDADE LAUDO DATA LAUDO N.º/ANO",
            "FACEPA FABRICA S/A 04.909.479/0001-06 BELEM PA IND Embalagens Art 2 Reducao Implantacao 27/03/2015 001/2015",
            "FWP SOUZA LTDA 12.972.611/0001-80 SANTANA AP IND Colchao Art 2 Reducao Implantacao 15/06/2015 002/2015",
        ],
    )
    rows = parse_sudam_pdf(pdf, 2015)
    assert len(rows) >= 2
    assert rows[0]["CNPJ"] == "04.909.479/0001-06"
    assert rows[0]["Órgão"] == "SUDAM"
    assert "FACEPA" in rows[0]["Empresa"].upper() or "FABRICA" in rows[0]["Empresa"].upper()
