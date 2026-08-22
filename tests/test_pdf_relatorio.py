"""Testes do exportador de PDF das evoluções SGS."""

from __future__ import annotations

from pathlib import Path

from scripts.evolucao_balanca_reservas import contar_paginas_pdf, exportar_pdf_relatorio


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
