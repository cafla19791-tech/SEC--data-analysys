"""Gera Excel ContAgil: 1 aba/pais com mes, taxa % a.m. e acumulados compostos.

Conversao anual -> mensal:
  taxa_am = (1 + taxa_aa)^(1/12) - 1
  fator_acum *= (1 + taxa_am)
"""

from __future__ import annotations

import math
import re
from collections import defaultdict
from decimal import Decimal, getcontext
from pathlib import Path
from typing import Any

from . import excel_diario, excel_format, providers

getcontext().prec = 80


def build_country_rows_mensal(points: list[tuple[str, float]]) -> list[dict[str, Any]]:
    """Mes | Taxa (% a.m.) | Taxa acumulada (%) | Taxa acumulada ano (%)."""
    parsed: list[tuple[str, Decimal, float]] = []
    for period, taxa_aa_pct in points:
        if taxa_aa_pct is None or not math.isfinite(float(taxa_aa_pct)):
            continue
        ym = str(period).strip()[:7]
        if len(ym) < 7 or ym[4] != "-":
            continue
        taxa_am = Decimal(str(excel_diario.taxa_mensal_composta_aa(taxa_aa_pct)))
        parsed.append((ym, taxa_am, float(taxa_aa_pct)))

    if not parsed:
        return []

    last_of_year: dict[int, str] = {}
    for ym, *_rest in parsed:
        year = int(ym[:4])
        if year not in last_of_year or ym > last_of_year[year]:
            last_of_year[year] = ym

    rows: list[dict[str, Any]] = []
    fator = Decimal(1)
    fator_ano: dict[int, Decimal] = defaultdict(lambda: Decimal(1))

    for ym, taxa_am, taxa_aa_pct in parsed:
        year = int(ym[:4])
        fator *= Decimal(1) + taxa_am
        fator_ano[year] *= Decimal(1) + taxa_am

        ano_val = None
        if ym == last_of_year[year]:
            ano_val = excel_diario._as_excel_number(
                (fator_ano[year] - Decimal(1)) * Decimal(100)
            )

        rows.append(
            {
                "Mês": ym,
                "Taxa (% a.m.)": float(taxa_am * Decimal(100)),
                "Taxa acumulada (%)": excel_diario._as_excel_number(
                    (fator - Decimal(1)) * Decimal(100)
                ),
                "Taxa acumulada ano (%)": ano_val,
                "_taxa_aa": taxa_aa_pct,
            }
        )
    return rows


def gerar_excel_mensal(
    out_path: str | Path,
    *,
    csv_path: str | Path | None = None,
    areas: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    prefer_local: bool = True,
) -> dict[str, Any]:
    """Gera workbook .xlsx mensal com uma aba por país."""
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pandas e necessario. pip install pandas openpyxl") from exc

    wanted: set[str] | None = None
    if areas:
        wanted = {
            providers.resolve_area(a)["code"]
            for a in re.split(r"[\s,;+|]+", areas)
            if a
        }

    local = providers.find_local_flat_csv(csv_path) if prefer_local or csv_path else None
    if local is None and prefer_local:
        dest = Path(out_path).resolve().parent
        providers.download_flat_csv(dest)
        local = providers.find_local_flat_csv(dest / providers.FLAT_CSV_NAME)

    # Monthly periods are YYYY-MM; allow YYYY-MM-DD filters via prefix compare.
    if local is not None:
        series, names = excel_diario._load_freq_from_flat(
            local, wanted, date_from, date_to, freq="M"
        )
        source = f"local:{local}"
    else:
        if wanted is None:
            wanted = set(excel_diario.AREA_NAMES.keys())
        series, names = excel_diario._load_freq_from_sdmx(
            sorted(wanted), date_from, date_to, freq="M"
        )
        source = "sdmx"

    if not series:
        raise RuntimeError("Nenhuma serie mensal encontrada para os filtros informados.")

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    indice_rows = []
    country_tables: dict[str, list[dict[str, Any]]] = {}
    for code in sorted(series.keys()):
        points = series[code]
        if not points:
            continue
        table = build_country_rows_mensal(points)
        if not table:
            continue
        country_tables[code] = table
        indice_rows.append(
            {
                "Codigo": code,
                "Pais": names.get(code, excel_diario.AREA_NAMES.get(code, code)),
                "Aba": excel_diario.sheet_name(
                    code, names.get(code, excel_diario.AREA_NAMES.get(code, code))
                ),
                "N_meses": len(table),
                "Inicio": table[0]["Mês"],
                "Fim": table[-1]["Mês"],
                "Taxa_aa_ultimo_%": table[-1]["_taxa_aa"],
                "Taxa_am_ultimo_%": table[-1]["Taxa (% a.m.)"],
                "Taxa_acumulada_final_%": table[-1]["Taxa acumulada (%)"],
            }
        )

    legenda = pd.DataFrame(
        [
            {"Item": "Fonte", "Valor": "BIS WS_CBPOL (frequencia mensal)"},
            {"Item": "Origem dados", "Valor": source},
            {
                "Item": "Conversao % a.a. -> % a.m.",
                "Valor": "taxa_am = (1 + taxa_aa/100)^(1/12) - 1",
            },
            {
                "Item": "Acumulacao (juros compostos)",
                "Valor": "fator *= (1 + taxa_am);  taxa_acumulada_% = (fator - 1)*100",
            },
            {
                "Item": "Taxa acumulada ano (%)",
                "Valor": "Preenchida so no ultimo mes do ano da serie",
            },
            {
                "Item": "Colunas por pais",
                "Valor": "Mes | Taxa (% a.m.) | Taxa acumulada (%) | Taxa acumulada ano (%)",
            },
        ]
    )

    engine = "xlsxwriter"
    try:
        import xlsxwriter  # noqa: F401
    except ImportError:
        engine = "openpyxl"

    with pd.ExcelWriter(out, engine=engine) as writer:
        legenda.to_excel(writer, sheet_name="00_Legenda", index=False)
        excel_format.autosize_dataframe_sheet(
            writer,
            "00_Legenda",
            legenda,
            engine=engine,
            max_width=80,
            padding=4,
            center=True,
            print_layout=True,
        )
        indice_df = pd.DataFrame(indice_rows)
        indice_df.to_excel(writer, sheet_name="01_Indice", index=False)
        excel_format.autosize_dataframe_sheet(
            writer,
            "01_Indice",
            indice_df,
            engine=engine,
            padding=4,
            center=True,
            print_layout=True,
        )

        formats_tpl = None
        if engine == "xlsxwriter":
            formats_tpl = excel_format.make_center_formats(
                writer.book,
                [None, "0.00000000", "0.000000", "0.000000"],
            )

        for code in sorted(country_tables.keys()):
            name = names.get(code, excel_diario.AREA_NAMES.get(code, code))
            aba = excel_diario.sheet_name(code, name)
            df = pd.DataFrame(country_tables[code])[
                [
                    "Mês",
                    "Taxa (% a.m.)",
                    "Taxa acumulada (%)",
                    "Taxa acumulada ano (%)",
                ]
            ]
            df.to_excel(writer, sheet_name=aba, index=False)
            excel_format.autosize_dataframe_sheet(
                writer,
                aba,
                df,
                engine=engine,
                col_formats=formats_tpl,
                min_width=12,
                max_width=36,
                padding=4,
                center=True,
                print_layout=True,
            )

    return {
        "path": str(out.resolve()),
        "source": source,
        "countries": len(country_tables),
        "rows_total": sum(len(v) for v in country_tables.values()),
        "bytes": out.stat().st_size,
        "freq": "M",
        "date_from": date_from,
        "date_to": date_to,
    }
