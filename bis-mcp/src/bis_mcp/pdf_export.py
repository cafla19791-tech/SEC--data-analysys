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
    """Converte um .xlsx em .pdf (todas as abas) usando LibreOffice headless.

    Prefira planilhas geradas com print_layout (paisagem + fit-to-width): o
    LibreOffice respeita a configuracao de pagina do .xlsx na exportacao.
    """
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

    # calc_pdf_Export: usa page setup da planilha (landscape/fit) quando presente.
    # UseSinglePageSheets=false evita esmagar abas enormes em uma unica pagina.
    convert_filter = (
        'pdf:calc_pdf_Export:'
        '{"SinglePageSheets":{"type":"boolean","value":"false"},'
        '"UseLosslessCompression":{"type":"boolean","value":"true"}}'
    )
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        convert_filter,
        "--outdir",
        str(dest_dir),
        str(src),
    ]
    # Diarios grandes podem passar de 10 min
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)
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
