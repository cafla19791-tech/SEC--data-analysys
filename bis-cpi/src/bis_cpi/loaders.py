"""Leitura do CSV colunar WS_LONG_CPI_csv_col.csv."""

from __future__ import annotations

import csv
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

# UNIT_MEASURE codes in the bulk file
UNIT_INDEX = "628"  # Index, 2010 = 100
UNIT_YOY = "771"  # Year-on-year changes, in per cent

_MONTH_RE = re.compile(r"^(\d{4})-(\d{2})$")
_YEAR_RE = re.compile(r"^(\d{4})$")


def _is_time_col(name: str) -> bool:
    n = (name or "").strip()
    return bool(_MONTH_RE.match(n) or _YEAR_RE.match(n))


def find_csv(explicit: str | Path | None = None) -> Path:
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit))
    here = Path.cwd()
    candidates.extend(
        [
            here / "WS_LONG_CPI_csv_col.csv",
            here.parent / "WS_LONG_CPI_csv_col.csv",
            Path("/workspace/bis_cpi/extracted/WS_LONG_CPI_csv_col.csv"),
            here / "bis_cpi" / "extracted" / "WS_LONG_CPI_csv_col.csv",
        ]
    )
    for p in candidates:
        if p.is_file():
            return p.resolve()
    raise FileNotFoundError(
        "WS_LONG_CPI_csv_col.csv nao encontrado. "
        "Baixe https://data.bis.org/static/bulk/WS_LONG_CPI_csv_col.zip"
    )


def load_series(
    csv_path: Path,
    *,
    freq: str = "M",
    unit: str = UNIT_INDEX,
    areas: set[str] | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
) -> tuple[dict[str, list[tuple[str, float]]], dict[str, str]]:
    """Return {REF_AREA: [(TIME_PERIOD, value), ...]} sorted + English names from file."""
    freq = (freq or "M").strip().upper()[:1]
    unit = (unit or UNIT_INDEX).strip()
    series: dict[str, list[tuple[str, float]]] = defaultdict(list)
    names: dict[str, str] = {}

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        time_cols = [c for c in (reader.fieldnames or []) if _is_time_col(c)]
        if freq == "M":
            time_cols = [c for c in time_cols if _MONTH_RE.match(c)]
        else:
            time_cols = [c for c in time_cols if _YEAR_RE.match(c)]

        for row in reader:
            if (row.get("FREQ") or "").strip().upper() != freq:
                continue
            if (row.get("UNIT_MEASURE") or "").strip() != unit:
                continue
            area = (row.get("REF_AREA") or "").strip().upper()
            if not area:
                continue
            if areas is not None and area not in areas:
                continue
            names[area] = (row.get("Reference area") or area).strip()
            points: list[tuple[str, float]] = []
            for col in time_cols:
                raw = (row.get(col) or "").strip()
                if not raw:
                    continue
                if date_from and col < date_from:
                    continue
                if date_to and col > date_to:
                    continue
                try:
                    val = float(raw)
                except ValueError:
                    continue
                points.append((col, val))
            if points:
                series[area] = points

    for area in series:
        series[area].sort(key=lambda x: x[0])
    return dict(series), names


def summarize_csv(csv_path: Path) -> dict[str, Any]:
    series_idx, names = load_series(csv_path, freq="M", unit=UNIT_INDEX)
    return {
        "csv": str(csv_path),
        "n_countries_monthly_index": len(series_idx),
        "countries": sorted(series_idx),
        "names": {k: names.get(k, k) for k in sorted(series_idx)},
    }
