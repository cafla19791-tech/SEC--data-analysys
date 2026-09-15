"""CLI ContAgil / local para WS_LONG_CPI."""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path
from urllib.request import urlretrieve

from . import loaders
from .excel_mensal import gerar_excel_mensal
from .excel_periodos import gerar_excel_periodos
from .pdf_export import para_pdf

BULK_URL = "https://data.bis.org/static/bulk/WS_LONG_CPI_csv_col.zip"


def _print(data: object) -> None:
    print(json.dumps(data, indent=2, default=str, ensure_ascii=True))


def _parse_areas(raw: str) -> set[str] | None:
    raw = (raw or "").strip()
    if not raw:
        return None
    return {a.strip().upper() for a in raw.split(",") if a.strip()}


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="bis-cpi-cli",
        description="BIS WS_LONG_CPI — Excel/PDF por país e rankings de inflação acumulada",
    )
    sub = p.add_subparsers(dest="command", required=True)

    d = sub.add_parser("download", help="Baixa e extrai WS_LONG_CPI_csv_col.zip")
    d.add_argument("--dir", default=".", help="Pasta destino")

    i = sub.add_parser("info", help="Resumo do CSV local")
    i.add_argument("--csv", default="")

    xm = sub.add_parser("excel-mensal", help="1 aba/país: índice, YoY e acumulada")
    xm.add_argument("--out", default="cpi_mensal_por_pais.xlsx")
    xm.add_argument("--csv", default="")
    xm.add_argument("--areas", default="")
    xm.add_argument("--from", dest="date_from", default="")
    xm.add_argument("--to", dest="date_to", default="")

    xp = sub.add_parser("excel-periodos", help="Rankings de inflação acumulada por período")
    xp.add_argument("--out", default="cpi_inflacao_acumulada_periodos.xlsx")
    xp.add_argument("--csv", default="")
    xp.add_argument("--areas", default="")

    pdf = sub.add_parser("para-pdf", help="Converte XLSX em PDF (LibreOffice)")
    pdf.add_argument("xlsx")
    pdf.add_argument("--outdir", default="pdf")

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.command == "download":
        dest = Path(args.dir)
        dest.mkdir(parents=True, exist_ok=True)
        zpath = dest / "WS_LONG_CPI_csv_col.zip"
        print(f"Baixando {BULK_URL} ...")
        urlretrieve(BULK_URL, zpath)
        with zipfile.ZipFile(zpath) as zf:
            zf.extractall(dest)
        csv_path = dest / "WS_LONG_CPI_csv_col.csv"
        _print({"zip": str(zpath), "csv": str(csv_path), "ok": csv_path.is_file()})
        return 0

    if args.command == "info":
        csv_path = loaders.find_csv(args.csv or None)
        _print(loaders.summarize_csv(csv_path))
        return 0

    if args.command == "excel-mensal":
        csv_path = loaders.find_csv(args.csv or None)
        out = gerar_excel_mensal(
            csv_path,
            Path(args.out),
            areas=_parse_areas(args.areas),
            date_from=args.date_from or None,
            date_to=args.date_to or None,
        )
        _print({"out": str(out), "csv": str(csv_path)})
        return 0

    if args.command == "excel-periodos":
        csv_path = loaders.find_csv(args.csv or None)
        out = gerar_excel_periodos(
            csv_path,
            Path(args.out),
            areas=_parse_areas(args.areas),
        )
        _print({"out": str(out), "csv": str(csv_path)})
        return 0

    if args.command == "para-pdf":
        dest = para_pdf(Path(args.xlsx), Path(args.outdir))
        _print({"pdf": str(dest)})
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
