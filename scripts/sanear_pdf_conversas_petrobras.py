#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Remove páginas em branco do PDF das conversas Petrobras e reemite o arquivo.

Uso::

  python scripts/sanear_pdf_conversas_petrobras.py
  python scripts/sanear_pdf_conversas_petrobras.py --pdf output/conversas_petrobras_20f_2026-09-01.pdf
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_PDF = ROOT / "output" / "conversas_petrobras_20f_2026-09-01.pdf"


def texto_util(page: fitz.Page) -> str:
    lines = []
    for line in page.get_text().splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith("Petrobras — conversas"):
            continue
        if s.startswith("Página "):
            continue
        lines.append(s)
    return "\n".join(lines)


def paginas_em_branco(doc: fitz.Document) -> list[int]:
    vazias = []
    for i, page in enumerate(doc):
        if len(texto_util(page)) < 40:
            vazias.append(i)
    return vazias


def renumerar_rodape(doc: fitz.Document) -> None:
    """Corrige o 'Página N' do rodapé após apagar folhas em branco."""
    for i, page in enumerate(doc, 1):
        rodape = fitz.Rect(400, page.rect.height - 50, page.rect.width, page.rect.height)
        for n in range(1, 40):
            for rect in page.search_for(f"Página {n}"):
                if rect.intersects(rodape):
                    page.add_redact_annot(rect + (-1, -1, 8, 1), fill=(1, 1, 1))
        page.apply_redactions()
        page.insert_text(
            (516.5, 813.5),
            f"Página {i}",
            fontname="times-roman",
            fontsize=8,
            color=(91 / 255, 101 / 255, 112 / 255),
        )


def sanear(pdf_path: Path) -> tuple[Path, list[int], int]:
    doc = fitz.open(pdf_path)
    vazias = paginas_em_branco(doc)
    for i in reversed(vazias):
        doc.delete_page(i)
    renumerar_rodape(doc)
    tmp = pdf_path.with_suffix(".tmp.pdf")
    doc.save(tmp, deflate=True, garbage=4)
    n = doc.page_count
    doc.close()
    shutil.move(tmp, pdf_path)
    art = Path("/opt/cursor/artifacts") / pdf_path.name
    if art.parent.is_dir():
        art.write_bytes(pdf_path.read_bytes())
    return pdf_path, vazias, n


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--pdf", type=Path, default=DEFAULT_PDF)
    args = p.parse_args()
    path, removidas, n = sanear(args.pdf)
    print(f"pdf: {path}")
    print(f"paginas_removidas: {[i + 1 for i in removidas] or 'nenhuma'}")
    print(f"paginas_finais: {n}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
