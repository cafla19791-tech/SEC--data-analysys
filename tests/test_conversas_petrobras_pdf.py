"""O PDF das conversas não pode ter página em branco e segue a sequência pedida."""

from __future__ import annotations

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


def test_caixa_operacional_vem_logo_apos_a_divida_no_indice():
    doc = fitz.open(DEFAULT_PDF)
    try:
        indice = doc[1].get_text()
        i_div = indice.lower().index("dívida bruta")
        i_cx = indice.lower().index("geração operacional de caixa")
        i_jur = indice.lower().index("juros pagos")
        i_luc = indice.lower().index("lucro líquido")
        assert i_div < i_cx < i_jur < i_luc
    finally:
        doc.close()


def test_tabelas_na_sequencia_divida_caixa_juros_lucro():
    doc = fitz.open(DEFAULT_PDF)
    try:
        texto = "".join(page.get_text() for page in doc)
    finally:
        doc.close()
    i_div = texto.index("evolução das dívidas brutas")
    i_cx = texto.index("geração operacional de caixa no anos 2002 a 2026")
    i_jur = texto.index("juros pagos pela Petrobras")
    i_luc = texto.index("evolução dos lucros líquidos")
    assert i_div < i_cx < i_jur < i_luc
    assert "643.420" in texto
    assert "664.069" in texto
