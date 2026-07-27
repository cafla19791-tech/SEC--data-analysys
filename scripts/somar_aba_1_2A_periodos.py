"""Totais por item (linhas 6–177) da aba 1.2-A, somando as colunas mensais de cada período.

Para cada item (ex. A6, A7, …, A177) e cada período:

    total = SOMA(colunas_do_periodo naquela linha)

Exemplo período jan/97–dez/02 (72 meses = B:BU)::

    A6 → SOMA(B6:BU6)
    A7 → SOMA(B7:BU7)
    …

Uso::

    python somar_aba_1_2A_periodos.py "serie_historica_mai26 (2).xlsx" --out saida/aba_1_2A_totais_por_item.xlsx
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

PERIODS: list[tuple[str, datetime, datetime]] = [
    ("jan/97–dez/02", datetime(1997, 1, 1), datetime(2002, 12, 1)),
    ("jan/03–mai/16", datetime(2003, 1, 1), datetime(2016, 5, 1)),
    ("jun/16–dez/18", datetime(2016, 6, 1), datetime(2018, 12, 1)),
    ("jan/19–dez/22", datetime(2019, 1, 1), datetime(2022, 12, 1)),
    ("jan/23–mai/26", datetime(2023, 1, 1), datetime(2026, 5, 1)),
]

SHEET = "1.2-A"
ROW_START = 6
ROW_END = 177


def _load(path: str | Path):
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("Instale openpyxl: pip install openpyxl") from exc
    return load_workbook(path, read_only=True, data_only=True)


def _col_letter(col_idx_0based: int) -> str:
    from openpyxl.utils import get_column_letter

    return get_column_letter(col_idx_0based + 1)


def totails_por_item(path: str | Path) -> dict[str, Any]:
    """Retorna 172 itens × 5 períodos (soma mensal na própria linha)."""
    wb = _load(path)
    if SHEET not in wb.sheetnames:
        raise ValueError(f"Aba '{SHEET}' nao encontrada. Disponiveis: {wb.sheetnames}")
    ws = wb[SHEET]
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[4]
    unit = rows[2][0] if len(rows) > 2 else None

    period_meta: list[dict[str, Any]] = []
    for lab, start, end in PERIODS:
        cols = [
            col
            for col, h in enumerate(headers)
            if col > 0 and isinstance(h, datetime) and start <= h <= end
        ]
        period_meta.append(
            {
                "label": lab,
                "cols": cols,
                "meses": len(cols),
                "colunas_excel": f"{_col_letter(cols[0])}:{_col_letter(cols[-1])}",
                "inicio": start.strftime("%Y-%m"),
                "fim": end.strftime("%Y-%m"),
            }
        )

    itens: list[dict[str, Any]] = []
    for ridx in range(ROW_START - 1, ROW_END):
        row = rows[ridx]
        label = row[0]
        somas: dict[str, float] = {}
        for pm in period_meta:
            s = 0.0
            for c in pm["cols"]:
                v = row[c] if c < len(row) else None
                if v is None:
                    continue
                try:
                    s += float(v)
                except (TypeError, ValueError):
                    continue
            somas[pm["label"]] = round(s, 2)
        itens.append({"linha": ridx + 1, "item": label, **somas})

    return {
        "path": str(path),
        "sheet": SHEET,
        "unit": unit,
        "row_start": ROW_START,
        "row_end": ROW_END,
        "n_itens": len(itens),
        "periodos": [
            {
                "periodo": pm["label"],
                "inicio": pm["inicio"],
                "fim": pm["fim"],
                "meses": pm["meses"],
                "colunas_excel": pm["colunas_excel"],
                "formula_exemplo_A6": f"SOMA({pm['colunas_excel']}6)",
            }
            for pm in period_meta
        ],
        "itens": itens,
        "nota": (
            "Cada célula = soma das colunas mensais do período NA MESMA LINHA "
            "(ex.: jan/97–dez/02 na linha 6 = SOMA(B6:BU6)). "
            "São 172 totais por período, um por item A6–A177."
        ),
    }


# alias for older imports/tests
sum_periods = totails_por_item


def write_outputs(table: dict[str, Any], out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    period_labels = [p["periodo"] for p in table["periodos"]]

    if out.suffix.lower() == ".csv":
        with out.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["linha", "item"] + period_labels)
            w.writerow(
                ["", "colunas_excel"]
                + [p["colunas_excel"] for p in table["periodos"]]
            )
            w.writerow(["", "n_meses"] + [p["meses"] for p in table["periodos"]])
            for row in table["itens"]:
                w.writerow([row["linha"], row["item"]] + [row[p] for p in period_labels])
        return

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    xwb = Workbook()
    ws1 = xwb.active
    ws1.title = "Totais por item"
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="D9E2F3")
    ws1["A1"] = "Aba 1.2-A — Total de cada item (A6–A177) por período"
    ws1["A1"].font = bold
    ws1["A2"] = table.get("unit") or "R$ milhões — IPCA"
    ws1["A3"] = table["nota"]
    ws1["A4"] = f"Fonte: {table['path']}"

    headers = ["Linha", "Item"] + period_labels
    for c, h in enumerate(headers, 1):
        cell = ws1.cell(6, c, h)
        cell.font = bold
        cell.fill = fill
    ws1.cell(7, 2, "Colunas Excel")
    for j, p in enumerate(table["periodos"]):
        ws1.cell(7, 3 + j, p["colunas_excel"])
    ws1.cell(8, 2, "Nº meses")
    for j, p in enumerate(table["periodos"]):
        ws1.cell(8, 3 + j, p["meses"])

    for i, row in enumerate(table["itens"]):
        r = 9 + i
        ws1.cell(r, 1, row["linha"])
        ws1.cell(r, 2, row["item"])
        for j, p in enumerate(period_labels):
            cell = ws1.cell(r, 3 + j, row[p])
            cell.number_format = "#,##0.00"

    ws1.column_dimensions["A"].width = 8
    ws1.column_dimensions["B"].width = 72
    for col in range(3, 3 + len(period_labels)):
        ws1.column_dimensions[get_column_letter(col)].width = 16

    ws2 = xwb.create_sheet("Legenda periodos")
    for c, h in enumerate(
        ["Período", "Início", "Fim", "Meses", "Colunas Excel", "Fórmula ex. A6"], 1
    ):
        ws2.cell(1, c, h).font = bold
    for i, p in enumerate(table["periodos"]):
        ws2.cell(2 + i, 1, p["periodo"])
        ws2.cell(2 + i, 2, p["inicio"])
        ws2.cell(2 + i, 3, p["fim"])
        ws2.cell(2 + i, 4, p["meses"])
        ws2.cell(2 + i, 5, p["colunas_excel"])
        ws2.cell(2 + i, 6, p["formula_exemplo_A6"])
    ws2["A8"] = (
        "jan/97–dez/02 = 72 meses = B:BU (não BV). "
        "BV inicia o período seguinte (jan/03)."
    )
    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["E"].width = 14
    ws2.column_dimensions["F"].width = 18
    xwb.save(out)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Totais por item (A6–A177) da aba 1.2-A em 5 períodos"
    )
    p.add_argument("path", help="Caminho do serie_historica_*.xlsx")
    p.add_argument("--out", default="aba_1_2A_totais_por_item.xlsx")
    args = p.parse_args(argv)
    path = Path(args.path)
    if not path.exists():
        print(json.dumps({"error": f"arquivo nao encontrado: {path}"}, ensure_ascii=False))
        return 1
    try:
        table = totails_por_item(path)
        out = Path(args.out)
        write_outputs(table, out)
        print(
            json.dumps(
                {
                    "ok": True,
                    "out": str(out.resolve()),
                    "unit": table["unit"],
                    "n_itens": table["n_itens"],
                    "periodos": table["periodos"],
                    "amostra_itens": table["itens"][:5],
                    "nota": table["nota"],
                },
                indent=2,
                ensure_ascii=False,
            )
        )
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"error": str(exc)}, ensure_ascii=False))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
