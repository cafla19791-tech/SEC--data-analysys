"""Converte XLSX → PDF via LibreOffice soffice."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


def find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    raise FileNotFoundError("LibreOffice (soffice) nao encontrado no PATH")


def para_pdf(xlsx: Path, outdir: Path) -> Path:
    xlsx = Path(xlsx).resolve()
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    if not xlsx.is_file():
        raise FileNotFoundError(xlsx)
    soffice = find_soffice()
    with tempfile.TemporaryDirectory(prefix="cpi_pdf_") as tmp:
        cmd = [
            soffice,
            "--headless",
            "--nologo",
            "--nolockcheck",
            "--convert-to",
            "pdf",
            "--outdir",
            tmp,
            str(xlsx),
        ]
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        produced = list(Path(tmp).glob("*.pdf"))
        if not produced:
            raise RuntimeError(f"PDF nao gerado a partir de {xlsx}")
        dest = outdir / f"{xlsx.stem}.pdf"
        shutil.move(str(produced[0]), str(dest))
    return dest
