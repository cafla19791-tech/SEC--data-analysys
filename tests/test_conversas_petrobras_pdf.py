"""O PDF das conversas não pode ter página em branco."""

from __future__ import annotations

from pathlib import Path

import fitz

from scripts.sanear_pdf_conversas_petrobras import DEFAULT_PDF, paginas_em_branco, texto_util


def test_pdf_existe():
    assert DEFAULT_PDF.exists()


def test_nenhuma_pagina_em_branco():
    doc = fitz.open(DEFAULT_PDF)
    try:
        assert paginas_em_branco(doc) == []
        for i, page in enumerate(doc, 1):
            assert len(texto_util(page)) >= 40, f"página {i} sem conteúdo"
    finally:
        doc.close()


def test_pagina_3_tem_inicio_das_conversas():
    doc = fitz.open(DEFAULT_PDF)
    try:
        assert doc.page_count >= 3
        texto = doc[2].get_text()
        assert "Página 3" in texto
        assert "Traz os links do form 20-da SEC" in texto or "20-F" in texto
        assert "0001119639" in texto or "Mensagem" in texto
    finally:
        doc.close()
