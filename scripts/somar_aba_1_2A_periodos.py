"""Soma linhas 6–177 da aba 1.2-A (IPCA) em períodos fiscais definidos.

Uso::

    python somar_aba_1_2A_periodos.py "serie_historica_mai26 (2).xlsx" --out saida/aba_1_2A_somas_periodos.xlsx
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
ROW_START = 6  # 1-indexed inclusive
ROW_END = 177  # 1-indexed inclusive


def _load(path: str | Path):
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("Instale openpyxl: pip install openpyxl") from exc
    return load_workbook(path, read_only=True, data_only=True)


def sum_periods(path: str | Path) -> dict[str, Any]:
    wb = _load(path)
    if SHEET not in wb.sheetnames:
        raise ValueError(f"Aba '{SHEET}' nao encontrada. Disponiveis: {wb.sheetnames}")
    ws = wb[SHEET]
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[4]  # linha 5
    unit = rows[2][0] if len(rows) > 2 else None

    period_cols: list[list[int]] = []
    for _lab, start, end in PERIODS:
        cols = [
            col
            for col, h in enumerate(headers)
            if col > 0 and isinstance(h, datetime) and start <= h <= end
        ]
        period_cols.append(cols)

    detail: list[dict[str, Any]] = []
    period_totals = [0.0] * len(PERIODS)

    for ridx in range(ROW_START - 1, ROW_END):
        row = rows[ridx]
        label = row[0]
        sums: list[float] = []
        for pi, cols in enumerate(period_cols):
            s = 0.0
            for c in cols:
                v = row[c] if c < len(row) else None
                if v is None:
                    continue
                try:
                    s += float(v)
                except (TypeError, ValueError):
                    continue
            sums.append(s)
            period_totals[pi] += s
        detail.append(
            {
                "linha": ridx + 1,
                "discriminacao": label,
                **{PERIODS[i][0]: round(sums[i], 2) for i in range(len(PERIODS))},
                "total_5_periodos": round(sum(sums), 2),
            }
        )

    resumo = []
    for i, (lab, start, end) in enumerate(PERIODS):
        resumo.append(
            {
                "periodo": lab,
                "inicio": start.strftime("%Y-%m"),
                "fim": end.strftime("%Y-%m"),
                "meses": len(period_cols[i]),
                "soma_linhas_6_177_R$mi": round(period_totals[i], 2),
                "soma_linhas_6_177_R$bi": round(period_totals[i] / 1000.0, 2),
            }
        )

    return {
        "path": str(path),
        "sheet": SHEET,
        "unit": unit,
        "row_start": ROW_START,
        "row_end": ROW_END,
        "resumo": resumo,
        "detail": detail,
        "nota": (
            "A soma agrega todas as linhas 6–177 (totais + subitens); "
            "ha dupla contagem contabil. Use a aba Detalhe para linha a linha."
        ),
    }


def write_outputs(table: dict[str, Any], out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    period_labels = [p[0] for p in PERIODS]

    if out.suffix.lower() == ".csv":
        with out.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["linha", "discriminacao"] + period_labels + ["total_5_periodos"])
            for row in table["detail"]:
                w.writerow(
                    [row["linha"], row["discriminacao"]]
                    + [row[p] for p in period_labels]
                    + [row["total_5_periodos"]]
                )
            w.writerow([])
            w.writerow(
                ["", "SOMA linhas 6–177"]
                + [r["soma_linhas_6_177_R$mi"] for r in table["resumo"]]
                + [round(sum(r["soma_linhas_6_177_R$mi"] for r in table["resumo"]), 2)]
            )
            w.writerow(
                ["", "SOMA linhas 6–177 (R$ bi)"]
                + [r["soma_linhas_6_177_R$bi"] for r in table["resumo"]]
                + [round(sum(r["soma_linhas_6_177_R$bi"] for r in table["resumo"]), 2)]
            )
        return

    from openpyxl import Workbook
    from openpyxl.styles import Font

    xwb = Workbook()
    ws1 = xwb.active
    ws1.title = "Resumo"
    bold = Font(bold=True)
    ws1["A1"] = "Aba 1.2-A — Soma linhas 6 a 177 por período"
    ws1["A1"].font = bold
    ws1["A2"] = table.get("unit") or "R$ milhões — IPCA"
    ws1["A3"] = f"Fonte: {table['path']}"
    ws1["A5"] = "Período"
    ws1["B5"] = "Meses"
    ws1["C5"] = "Soma linhas 6–177 (R$ mi)"
    ws1["D5"] = "Soma linhas 6–177 (R$ bi)"
    for c in ("A5", "B5", "C5", "D5"):
        ws1[c].font = bold
    for i, r in enumerate(table["resumo"]):
        row = 6 + i
        ws1[f"A{row}"] = f"{i + 1}) {r['periodo']}"
        ws1[f"B{row}"] = r["meses"]
        ws1[f"C{row}"] = r["soma_linhas_6_177_R$mi"]
        ws1[f"D{row}"] = r["soma_linhas_6_177_R$bi"]
    ws1["A12"] = table["nota"]
    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["C"].width = 32
    ws1.column_dimensions["D"].width = 32

    ws2 = xwb.create_sheet("Detalhe")
    headers = ["linha", "discriminacao"] + period_labels + ["total_5_periodos"]
    for c, h in enumerate(headers, 1):
        ws2.cell(1, c, h).font = bold
    for i, row in enumerate(table["detail"], 2):
        ws2.cell(i, 1, row["linha"])
        ws2.cell(i, 2, row["discriminacao"])
        for j, p in enumerate(period_labels):
            ws2.cell(i, 3 + j, row[p])
        ws2.cell(i, 8, row["total_5_periodos"])
    tr = len(table["detail"]) + 2
    ws2.cell(tr, 2, "SOMA linhas 6–177").font = bold
    for j, r in enumerate(table["resumo"]):
        ws2.cell(tr, 3 + j, r["soma_linhas_6_177_R$mi"]).font = bold
    ws2.column_dimensions["B"].width = 70
    xwb.save(out)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Soma aba 1.2-A linhas 6–177 por período")
    p.add_argument("path", help="Caminho do serie_historica_*.xlsx")
    p.add_argument("--out", default="aba_1_2A_somas_periodos.xlsx")
    args = p.parse_args(argv)
    path = Path(args.path)
    if not path.exists():
        print(json.dumps({"error": f"arquivo nao encontrado: {path}"}, ensure_ascii=False))
        return 1
    try:
        table = sum_periods(path)
        out = Path(args.out)
        write_outputs(table, out)
        print(
            json.dumps(
                {
                    "ok": True,
                    "out": str(out.resolve()),
                    "unit": table["unit"],
                    "resumo": table["resumo"],
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
