"""Converte planilhas .xlsx para PDF via LibreOffice (soffice)."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
from typing import Any


def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    # ContAgil / Windows common paths
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if Path(c).is_file():
            return c
    return None


def xlsx_para_pdf(
    xlsx_path: str | Path,
    out_dir: str | Path | None = None,
) -> dict[str, Any]:
    """Converte um .xlsx em .pdf (todas as abas) usando LibreOffice headless."""
    src = Path(xlsx_path).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"Arquivo nao encontrado: {src}")
    if src.suffix.lower() not in {".xlsx", ".xls", ".ods"}:
        raise ValueError(f"Formato nao suportado: {src.suffix}")

    dest_dir = Path(out_dir).resolve() if out_dir else src.parent
    dest_dir.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice (soffice) nao encontrado. Instale o LibreOffice "
            "ou use os PDFs pre-gerados em output/pdf/."
        )

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf:calc_pdf_Export",
        "--outdir",
        str(dest_dir),
        str(src),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    if proc.returncode != 0:
        raise RuntimeError(
            f"Falha na conversao (exit {proc.returncode}): {proc.stderr or proc.stdout}"
        )

    pdf = dest_dir / f"{src.stem}.pdf"
    if not pdf.is_file():
        raise RuntimeError(f"PDF nao gerado em {dest_dir}. stdout={proc.stdout!r}")

    return {
        "source": str(src),
        "path": str(pdf),
        "bytes": pdf.stat().st_size,
        "soffice": soffice,
    }
