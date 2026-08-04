"""Excel resumo: taxa acumulada por pais em periodos fixos.

Frequencia diaria (padrao):
  taxa_ad = (1 + taxa_aa/100)^(1/252) - 1
  acumula apenas segunda a sexta.

Frequencia mensal (--freq M):
  taxa_am = (1 + taxa_aa/100)^(1/12) - 1
  acumula todos os meses com observacao no periodo.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, getcontext
from pathlib import Path
from typing import Any

from . import excel_diario, excel_format, providers

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
    """Acumula compostos diarios no periodo, ignorando sabados e domingos."""
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
        "n_obs": n,
        "inicio_obs": primeiro.isoformat() if primeiro else None,
        "fim_obs": ultimo.isoformat() if ultimo else None,
        "taxa_acumulada": excel_diario._as_excel_number(acum),
        "taxa_acumulada_sort": float(acum)
        if abs(float(acum)) < 1e307
        else float("inf") * (1 if acum >= 0 else -1),
    }


def acumular_periodo_mensal(
    points: list[tuple[str, float]],
    inicio: date,
    fim: date,
) -> dict[str, Any] | None:
    """Acumula compostos mensais no periodo (TIME_PERIOD YYYY-MM)."""
    ym_from = f"{inicio:%Y-%m}"
    ym_to = f"{fim:%Y-%m}"
    fator = Decimal(1)
    n = 0
    primeiro: str | None = None
    ultimo: str | None = None
    for period, taxa_aa in points:
        ym = str(period).strip()[:7]
        if len(ym) < 7 or ym[4] != "-":
            continue
        if ym < ym_from or ym > ym_to:
            continue
        taxa_am = Decimal(str(excel_diario.taxa_mensal_composta_aa(taxa_aa)))
        fator *= Decimal(1) + taxa_am
        n += 1
        if primeiro is None:
            primeiro = ym
        ultimo = ym
    if n == 0:
        return None
    acum = (fator - Decimal(1)) * Decimal(100)
    return {
        "n_obs": n,
        "inicio_obs": primeiro,
        "fim_obs": ultimo,
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
    freq: str = "D",
) -> dict[str, Any]:
    """Gera workbook com 6 abas de ranking por periodo (+ legenda)."""
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pandas e necessario. pip install pandas openpyxl") from exc

    freq_code = excel_diario._normalize_freq_code(freq)
    if freq_code not in {"D", "M"}:
        raise ValueError("excel-periodos aceita apenas frequencia D ou M")

    local = providers.find_local_flat_csv(csv_path) if prefer_local or csv_path else None
    if local is None and prefer_local:
        dest = Path(out_path).resolve().parent
        providers.download_flat_csv(dest)
        local = providers.find_local_flat_csv(dest / providers.FLAT_CSV_NAME)

    date_from = min(p.inicio for p in PERIODOS).isoformat()
    date_to = max(p.fim for p in PERIODOS).isoformat()
    # Monthly compare works with YYYY-MM prefix.
    if freq_code == "M":
        date_from = date_from[:7]
        date_to = date_to[:7]

    if local is not None:
        series, names = excel_diario._load_freq_from_flat(
            local, None, date_from, date_to, freq=freq_code
        )
        source = f"local:{local}"
    else:
        # Fallback SDMX: known set (full universe via flat is preferred).
        codes = sorted(excel_diario.AREA_NAMES.keys())
        series, names = excel_diario._load_freq_from_sdmx(
            codes, date_from, date_to, freq=freq_code
        )
        source = "sdmx"

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    if freq_code == "M":
        legenda = pd.DataFrame(
            [
                {"Item": "Fonte", "Valor": "BIS WS_CBPOL (frequencia mensal)"},
                {"Item": "Origem dados", "Valor": source},
                {
                    "Item": "Conversao % a.a. -> % a.m.",
                    "Valor": "taxa_am = (1 + taxa_aa/100)^(1/12) - 1",
                },
                {
                    "Item": "Acumulacao",
                    "Valor": "fator *= (1 + taxa_am); taxa_acumulada_% = (fator - 1)*100",
                },
                {
                    "Item": "Ordenacao",
                    "Valor": "Cada aba em ordem crescente da taxa acumulada do periodo",
                },
                {
                    "Item": "Colunas",
                    "Valor": "Pais | Taxa acumulada (%)  (+ codigo e N_meses)",
                },
            ]
        )
        n_col = "N_meses"
        acum_fn = acumular_periodo_mensal
    else:
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
        n_col = "N_dias_uteis"
        acum_fn = acumular_periodo

    engine = "xlsxwriter"
    try:
        import xlsxwriter  # noqa: F401
    except ImportError:
        engine = "openpyxl"

    period_frames: list[tuple[Periodo, "pd.DataFrame"]] = []
    for periodo in PERIODOS:
        rows: list[dict[str, Any]] = []
        for code in sorted(series.keys()):
            stats = acum_fn(series[code], periodo.inicio, periodo.fim)
            if stats is None:
                continue
            pais = names.get(code) or excel_diario.AREA_NAMES.get(code, code)
            rows.append(
                {
                    "Pais": pais,
                    periodo.coluna_taxa: stats["taxa_acumulada"],
                    "Codigo": code,
                    n_col: stats["n_obs"],
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
                    n_col: r[n_col],
                    "Primeira_obs": r["Primeira_obs"],
                    "Ultima_obs": r["Ultima_obs"],
                }
            )
        period_frames.append((periodo, pd.DataFrame(ordered)))

    with pd.ExcelWriter(out, engine=engine) as writer:
        legenda.to_excel(writer, sheet_name="00_Legenda", index=False)
        excel_format.autosize_dataframe_sheet(
            writer, "00_Legenda", legenda, engine=engine, max_width=80, padding=4
        )
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
        excel_format.autosize_dataframe_sheet(
            writer, "01_Indice", idx, engine=engine, max_width=70, padding=4
        )

        for periodo, df in period_frames:
            sheet = _INVALID_SHEET.sub("-", periodo.sheet)[:31]
            df_out = df.copy()
            df_out.to_excel(writer, sheet_name=sheet, index=False, startrow=1)
            ws = writer.sheets[sheet]
            if engine == "xlsxwriter":
                ws.write(0, 0, periodo.titulo)
            else:
                ws.cell(row=1, column=1, value=periodo.titulo)
            excel_format.autosize_dataframe_sheet(
                writer,
                sheet,
                df_out,
                engine=engine,
                min_width=12,
                max_width=40,
                padding=4,
                extra_title_width=len(periodo.titulo) + 4,
            )

    return {
        "path": str(out.resolve()),
        "source": source,
        "periodos": len(PERIODOS),
        "countries_loaded": len(series),
        "bytes": out.stat().st_size,
        "sheets": [p.sheet for p in PERIODOS],
        "freq": freq_code,
    }
