"""Excel resumo: taxa acumulada por pais em periodos fixos (sem sab/dom).

Conversao: taxa_ad = (1 + taxa_aa/100)^(1/252) - 1
Acumulacao: fator *= (1 + taxa_ad)  apenas em dias de segunda a sexta.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, getcontext
from pathlib import Path
from typing import Any

from . import excel_diario, providers

getcontext().prec = 80

_INVALID_SHEET = re.compile(r"[\\/*?:\[\]]")


@dataclass(frozen=True)
class Periodo:
    sheet: str
    titulo: str
    inicio: date
    fim: date
    coluna_taxa: str = "Taxa acumulada (%)"


# Pedido do usuario (datas em DD/MM/AAAA).
PERIODOS: list[Periodo] = [
    Periodo(
        sheet="01_1995_a_2002",
        titulo="Taxa acumulada de 01/01/1995 a 31/12/2002",
        inicio=date(1995, 1, 1),
        fim=date(2002, 12, 31),
    ),
    Periodo(
        sheet="02_2003_a_2016-04",
        titulo="Taxa acumulada de 01/01/2003 a 30/04/2016",
        inicio=date(2003, 1, 1),
        fim=date(2016, 4, 30),
    ),
    Periodo(
        sheet="03_2016-05_a_2018",
        titulo="Taxa acumulada de 01/05/2016 a 31/12/2018",
        inicio=date(2016, 5, 1),
        fim=date(2018, 12, 31),
    ),
    Periodo(
        sheet="04_2019_a_2022",
        titulo="Taxa acumulada de 01/01/2019 a 31/12/2022",
        inicio=date(2019, 1, 1),
        fim=date(2022, 12, 31),
    ),
    Periodo(
        sheet="05_2023_a_2026-06",
        titulo="Taxa acumulada de 01/01/2023 a 30/06/2026",
        inicio=date(2023, 1, 1),
        fim=date(2026, 6, 30),
    ),
    Periodo(
        sheet="06_basica_2003_a_2026-06",
        titulo="Taxa basica acumulada de 01/01/2003 a 30/06/2026",
        inicio=date(2003, 1, 1),
        fim=date(2026, 6, 30),
        coluna_taxa="Taxa basica acumulada (%)",
    ),
]


def _parse_iso_day(value: str) -> date | None:
    s = (value or "").strip()
    if not s:
        return None
    # BIS daily: YYYY-MM-DD
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def is_weekday(d: date) -> bool:
    """True for Monday..Friday (excludes Saturday/Sunday)."""
    return d.weekday() < 5


def acumular_periodo(
    points: list[tuple[str, float]],
    inicio: date,
    fim: date,
) -> dict[str, Any] | None:
    """Acumula compostos no periodo, ignorando sabados e domingos."""
    fator = Decimal(1)
    n = 0
    primeiro: date | None = None
    ultimo: date | None = None
    for dia_s, taxa_aa in points:
        d = _parse_iso_day(dia_s)
        if d is None or d < inicio or d > fim:
            continue
        if not is_weekday(d):
            continue
        taxa_ad = Decimal(str(excel_diario.taxa_diaria_composta_aa(taxa_aa)))
        fator *= Decimal(1) + taxa_ad
        n += 1
        if primeiro is None:
            primeiro = d
        ultimo = d
    if n == 0:
        return None
    acum = (fator - Decimal(1)) * Decimal(100)
    return {
        "n_dias_uteis": n,
        "inicio_obs": primeiro.isoformat() if primeiro else None,
        "fim_obs": ultimo.isoformat() if ultimo else None,
        "taxa_acumulada": excel_diario._as_excel_number(acum),
        "taxa_acumulada_sort": float(acum)
        if abs(float(acum)) < 1e307
        else float("inf") * (1 if acum >= 0 else -1),
    }


def gerar_excel_periodos(
    out_path: str | Path,
    *,
    csv_path: str | Path | None = None,
    prefer_local: bool = True,
) -> dict[str, Any]:
    """Gera workbook com 6 abas de ranking por periodo (+ legenda)."""
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pandas e necessario. pip install pandas openpyxl") from exc

    local = providers.find_local_flat_csv(csv_path) if prefer_local or csv_path else None
    if local is None and prefer_local:
        dest = Path(out_path).resolve().parent
        providers.download_flat_csv(dest)
        local = providers.find_local_flat_csv(dest / providers.FLAT_CSV_NAME)

    date_from = min(p.inicio for p in PERIODOS).isoformat()
    date_to = max(p.fim for p in PERIODOS).isoformat()

    if local is not None:
        series, names = excel_diario._load_daily_from_flat(local, None, date_from, date_to)
        source = f"local:{local}"
    else:
        # Fallback SDMX: known set (full universe via flat is preferred).
        codes = sorted(excel_diario.AREA_NAMES.keys())
        series, names = excel_diario._load_daily_from_sdmx(codes, date_from, date_to)
        source = "sdmx"

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    legenda = pd.DataFrame(
        [
            {"Item": "Fonte", "Valor": "BIS WS_CBPOL (frequencia diaria)"},
            {"Item": "Origem dados", "Valor": source},
            {
                "Item": "Conversao % a.a. -> % a.d.",
                "Valor": "taxa_ad = (1 + taxa_aa/100)^(1/252) - 1",
            },
            {
                "Item": "Dias excluidos",
                "Valor": "Sabados e domingos (nao entram na acumulacao)",
            },
            {
                "Item": "Acumulacao",
                "Valor": "fator *= (1 + taxa_ad); taxa_acumulada_% = (fator - 1)*100",
            },
            {
                "Item": "Ordenacao",
                "Valor": "Cada aba em ordem crescente da taxa acumulada do periodo",
            },
            {
                "Item": "Colunas",
                "Valor": "Pais | Taxa acumulada (%)  (+ codigo e dias uteis usados)",
            },
        ]
    )

    engine = "xlsxwriter"
    try:
        import xlsxwriter  # noqa: F401
    except ImportError:
        engine = "openpyxl"

    period_frames: list[tuple[Periodo, "pd.DataFrame"]] = []
    for periodo in PERIODOS:
        rows: list[dict[str, Any]] = []
        for code in sorted(series.keys()):
            stats = acumular_periodo(series[code], periodo.inicio, periodo.fim)
            if stats is None:
                continue
            pais = names.get(code) or excel_diario.AREA_NAMES.get(code, code)
            rows.append(
                {
                    "Pais": pais,
                    periodo.coluna_taxa: stats["taxa_acumulada"],
                    "Codigo": code,
                    "N_dias_uteis": stats["n_dias_uteis"],
                    "Primeira_obs": stats["inicio_obs"],
                    "Ultima_obs": stats["fim_obs"],
                    "_sort": stats["taxa_acumulada_sort"],
                }
            )
        rows.sort(key=lambda r: (r["_sort"], str(r["Pais"])))
        for r in rows:
            del r["_sort"]
        # Presentation order: Pais, taxa, then metadata
        ordered = []
        for r in rows:
            ordered.append(
                {
                    "Pais": r["Pais"],
                    periodo.coluna_taxa: r[periodo.coluna_taxa],
                    "Codigo": r["Codigo"],
                    "N_dias_uteis": r["N_dias_uteis"],
                    "Primeira_obs": r["Primeira_obs"],
                    "Ultima_obs": r["Ultima_obs"],
                }
            )
        period_frames.append((periodo, pd.DataFrame(ordered)))

    with pd.ExcelWriter(out, engine=engine) as writer:
        legenda.to_excel(writer, sheet_name="00_Legenda", index=False)
        # Indice dos periodos
        idx = pd.DataFrame(
            [
                {
                    "Aba": p.sheet,
                    "Periodo": p.titulo,
                    "Inicio": p.inicio.isoformat(),
                    "Fim": p.fim.isoformat(),
                    "N_paises": len(df),
                }
                for p, df in period_frames
            ]
        )
        idx.to_excel(writer, sheet_name="01_Indice", index=False)

        for periodo, df in period_frames:
            # Title row via sheet + first row comment in a small header sheet content:
            # Keep data starting at row 1; put titulo in a companion? User asked only
            # country + rate — include titulo as sheet name context; add Titulo row above?
            # Cleaner: write titulo in A1 merged then table — keep simple flat table.
            sheet = _INVALID_SHEET.sub("-", periodo.sheet)[:31]
            # Prepend a one-row description sheet content using two writes:
            df_out = df.copy()
            df_out.to_excel(writer, sheet_name=sheet, index=False, startrow=1)
            ws = writer.sheets[sheet]
            if engine == "xlsxwriter":
                ws.write(0, 0, periodo.titulo)
                ws.set_column(0, 0, 28)
                ws.set_column(1, 1, 26)
                ws.set_column(2, 2, 10)
                ws.set_column(3, 3, 14)
                ws.set_column(4, 5, 12)
            else:
                ws.cell(row=1, column=1, value=periodo.titulo)

    return {
        "path": str(out.resolve()),
        "source": source,
        "periodos": len(PERIODOS),
        "countries_loaded": len(series),
        "bytes": out.stat().st_size,
        "sheets": [p.sheet for p in PERIODOS],
    }
