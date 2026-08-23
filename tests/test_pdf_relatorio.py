"""Testes do exportador de PDF das evoluções SGS."""

from __future__ import annotations

from pathlib import Path

import pytest

from scripts.evolucao_balanca_reservas import contar_paginas_pdf, exportar_pdf_relatorio
from scripts.imprimir_evolucoes_pdf import _pdf_base, _pdf_completo, _pdf_dbgg

ROOT = Path(__file__).resolve().parents[1]
CACHE = ROOT / "data"


def test_exportar_pdf_relatorio_cria_paginas(tmp_path: Path) -> None:
    import matplotlib.pyplot as plt

    png = tmp_path / "grafico.png"
    fig, ax = plt.subplots(figsize=(4, 2))
    ax.plot([0, 1], [0, 1])
    fig.savefig(png, dpi=80)
    plt.close(fig)
    path = exportar_pdf_relatorio(
        tmp_path / "relatorio.pdf",
        "Relatório de teste",
        ["Valores em R$ bilhões."],
        tabelas=[("Tabela", ["Ano", "Valor"], [["1995", "1,0"], ["1996", "2,0"]], None)],
        imagens=[png],
    )
    assert path.exists()
    bruto = path.read_bytes()
    assert bruto.startswith(b"%PDF")
    assert contar_paginas_pdf(path) >= 3


def test_imprimir_base_dbgg_e_compilado(tmp_path: Path) -> None:
    if not (CACHE / "sgs_1788_base.csv").exists() or not (CACHE / "sgs_4537_dbgg_pib.csv").exists():
        pytest.skip("cache SGS ausente")
    base = _pdf_base(CACHE, tmp_path)
    dbgg = _pdf_dbgg(CACHE, tmp_path)
    completo = _pdf_completo(CACHE, tmp_path)
    for path in (base, dbgg, completo):
        assert path.exists()
        assert path.read_bytes().startswith(b"%PDF")
    assert contar_paginas_pdf(base) >= 4
    assert contar_paginas_pdf(dbgg) >= 4
    assert contar_paginas_pdf(completo) >= 7
