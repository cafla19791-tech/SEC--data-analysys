"""Gera Excel ContAgil: 1 aba/pais com dia, taxa % a.d. e acumulado composto.

Conversao anual -> diaria com ano de 252 dias uteis:
  taxa_ad = (1 + taxa_aa)^(1/252) - 1
  fator_acum *= (1 + taxa_ad)
  taxa_acumulada = (fator_acum - 1) * 100
"""

from __future__ import annotations

import csv
import math
import re
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, getcontext
from pathlib import Path
from typing import Any, Iterable

from . import excel_format, providers

getcontext().prec = 80

# Excel sheet name: max 31 chars; forbidden \ / ? * [ ]
_INVALID_SHEET = re.compile(r'[\\/*?:\[\]]')

AREA_NAMES: dict[str, str] = {
    "AR": "Argentina",
    "AU": "Australia",
    "BR": "Brasil",
    "CA": "Canada",
    "CH": "Suica",
    "CL": "Chile",
    "CN": "China",
    "CO": "Colombia",
    "CZ": "Tchequia",
    "DK": "Dinamarca",
    "ES": "Espanha",
    "GB": "Reino Unido",
    "HK": "Hong Kong",
    "HU": "Hungria",
    "ID": "Indonesia",
    "IL": "Israel",
    "IN": "India",
    "IS": "Islandia",
    "IT": "Italia",
    "JP": "Japao",
    "KR": "Coreia",
    "KW": "Kuwait",
    "MA": "Marrocos",
    "MK": "Macedonia do Norte",
    "MX": "Mexico",
    "MY": "Malasia",
    "NL": "Paises Baixos",
    "NO": "Noruega",
    "NZ": "Nova Zelandia",
    "PE": "Peru",
    "PH": "Filipinas",
    "PL": "Polonia",
    "RO": "Romenia",
    "RS": "Servia",
    "RU": "Russia",
    "SA": "Arabia Saudita",
    "SE": "Suecia",
    "TH": "Tailandia",
    "TR": "Turquia",
    "US": "Estados Unidos",
    "XM": "Zona do Euro",
    "ZA": "Africa do Sul",
}


DIAS_UTEIS_ANO = 252
MESES_ANO = 12


def taxa_diaria_composta_aa(taxa_aa_pct: float) -> float:
    """Taxa % a.a. -> taxa decimal ao dia (ano com 252 dias uteis)."""
    return (1.0 + float(taxa_aa_pct) / 100.0) ** (1.0 / DIAS_UTEIS_ANO) - 1.0


def taxa_mensal_composta_aa(taxa_aa_pct: float) -> float:
    """Taxa % a.a. -> taxa decimal ao mes: (1+r)^(1/12)-1."""
    return (1.0 + float(taxa_aa_pct) / 100.0) ** (1.0 / MESES_ANO) - 1.0


def sheet_name(code: str, name: str) -> str:
    base = f"{code} - {name}"
    base = _INVALID_SHEET.sub("-", base).strip() or code
    return base[:31]


def _load_freq_from_flat(
    csv_path: Path,
    areas: set[str] | None,
    date_from: str | None,
    date_to: str | None,
    *,
    freq: str = "D",
) -> tuple[dict[str, list[tuple[str, float]]], dict[str, str]]:
    """Return {area: [(period, taxa_aa_pct), ...]} sorted, plus area display names."""
    freq_code = _normalize_freq_code(freq)
    series: dict[str, list[tuple[str, float]]] = defaultdict(list)
    names: dict[str, str] = {}
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for raw in reader:
            data = providers._row_dict(raw)
            if providers._norm_code(data.get("FREQ")) != freq_code:
                continue
            area = providers._norm_code(data.get("REF_AREA"))
            if not area:
                continue
            if areas is not None and area not in areas:
                continue
            period = (data.get("TIME_PERIOD") or "").strip()
            if not period:
                continue
            if date_from and period < date_from:
                continue
            if date_to and period > date_to:
                continue
            obs = providers._parse_obs(data.get("OBS_VALUE"))
            if obs is None or not math.isfinite(float(obs)):
                continue
            series[area].append((period, float(obs)))
            if area not in names:
                # Prefer label after colon in original REF_AREA cell if present.
                ref_raw = ""
                for k, v in raw.items():
                    if providers._norm_key(k) == "REF_AREA" and v:
                        ref_raw = str(v)
                        break
                if ":" in ref_raw:
                    names[area] = ref_raw.split(":", 1)[1].strip()
                else:
                    names[area] = AREA_NAMES.get(area, area)
    for code in series:
        series[code].sort(key=lambda x: x[0])
    return series, names


def _normalize_freq_code(freq: str) -> str:
    f = (freq or "D").strip().upper()
    if f in {"D", "DAILY", "DIA", "DIARIO", "DIÁRIO"}:
        return "D"
    if f in {"M", "MONTHLY", "MES", "MENSAL"}:
        return "M"
    if f in {"A", "Y", "ANNUAL", "ANO", "ANUAL"}:
        return "A"
    raise ValueError(f"Frequencia invalida: {freq!r}")


def _load_daily_from_flat(
    csv_path: Path,
    areas: set[str] | None,
    date_from: str | None,
    date_to: str | None,
) -> tuple[dict[str, list[tuple[str, float]]], dict[str, str]]:
    return _load_freq_from_flat(csv_path, areas, date_from, date_to, freq="D")


def _load_freq_from_sdmx(
    area_codes: Iterable[str],
    date_from: str | None,
    date_to: str | None,
    *,
    freq: str = "D",
) -> tuple[dict[str, list[tuple[str, float]]], dict[str, str]]:
    codes = [providers.resolve_area(a)["code"] for a in area_codes]
    freq_code = _normalize_freq_code(freq)
    # Batch to keep URLs reasonable.
    series: dict[str, list[tuple[str, float]]] = defaultdict(list)
    names: dict[str, str] = {c: AREA_NAMES.get(c, c) for c in codes}
    batch_size = 8
    for i in range(0, len(codes), batch_size):
        batch = codes[i : i + batch_size]
        data = providers.get_policy_rates(
            ",".join(batch),
            freq=freq_code,
            date_from=date_from,
            date_to=date_to,
        )
        for row in data["series"]:
            series[row["ref_area"]].append((row["time_period"], float(row["value"])))
        for a in data["areas"]:
            names[a["code"]] = AREA_NAMES.get(a["code"], a.get("name") or a["code"])
    for code in series:
        series[code].sort(key=lambda x: x[0])
    return series, names


def _load_daily_from_sdmx(
    area_codes: Iterable[str],
    date_from: str | None,
    date_to: str | None,
) -> tuple[dict[str, list[tuple[str, float]]], dict[str, str]]:
    return _load_freq_from_sdmx(area_codes, date_from, date_to, freq="D")


def _as_excel_number(value: Decimal) -> float | str:
    """Excel only stores IEEE floats; keep big compounded values as text."""
    try:
        f = float(value)
    except (OverflowError, ValueError):
        return format(value, ".6e")
    if not math.isfinite(f):
        return format(value, ".6e")
    # Stay inside Excel's numeric range with margin.
    if abs(f) > 1e307:
        return format(value, ".6e")
    return f


def _parse_day(dia: str):
    return datetime.strptime(str(dia)[:10], "%Y-%m-%d").date()


def build_country_rows(points: list[tuple[str, float]]) -> list[dict[str, Any]]:
    """Dia | taxa a.d. | acumulada | acumulada mes (fim mes) | acumulada ano (fim ano)."""
    parsed: list[tuple[Any, str, Decimal, float]] = []
    for dia, taxa_aa_pct in points:
        if taxa_aa_pct is None or not math.isfinite(float(taxa_aa_pct)):
            continue
        try:
            d = _parse_day(dia)
        except ValueError:
            continue
        taxa_ad = Decimal(str(taxa_diaria_composta_aa(taxa_aa_pct)))
        parsed.append((d, str(dia)[:10], taxa_ad, float(taxa_aa_pct)))

    if not parsed:
        return []

    last_of_month: dict[tuple[int, int], Any] = {}
    last_of_year: dict[int, Any] = {}
    for d, *_rest in parsed:
        ym = (d.year, d.month)
        if ym not in last_of_month or d > last_of_month[ym]:
            last_of_month[ym] = d
        if d.year not in last_of_year or d > last_of_year[d.year]:
            last_of_year[d.year] = d

    rows: list[dict[str, Any]] = []
    fator = Decimal(1)
    fator_mes: dict[tuple[int, int], Decimal] = defaultdict(lambda: Decimal(1))
    fator_ano: dict[int, Decimal] = defaultdict(lambda: Decimal(1))

    for d, dia, taxa_ad, taxa_aa_pct in parsed:
        fator *= Decimal(1) + taxa_ad
        ym = (d.year, d.month)
        fator_mes[ym] *= Decimal(1) + taxa_ad
        fator_ano[d.year] *= Decimal(1) + taxa_ad

        mes_val = None
        ano_val = None
        if d == last_of_month[ym]:
            mes_val = _as_excel_number((fator_mes[ym] - Decimal(1)) * Decimal(100))
        if d == last_of_year[d.year]:
            ano_val = _as_excel_number((fator_ano[d.year] - Decimal(1)) * Decimal(100))

        rows.append(
            {
                "Dia": dia,
                "Taxa (% a.d.)": float(taxa_ad * Decimal(100)),
                "Taxa acumulada (%)": _as_excel_number(
                    (fator - Decimal(1)) * Decimal(100)
                ),
                "Taxa acumulada mês (%)": mes_val,
                "Taxa acumulada ano (%)": ano_val,
                "_taxa_aa": taxa_aa_pct,
                "_fator": fator,
            }
        )
    return rows


# Padrao: series diarias compostas a partir de 01/01/1995 ate o ultimo dia do pais.
DEFAULT_DATE_FROM = "1995-01-01"


def gerar_excel_diario(
    out_path: str | Path,
    *,
    csv_path: str | Path | None = None,
    areas: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    prefer_local: bool = True,
) -> dict[str, Any]:
    """Gera workbook .xlsx com uma aba por país.

    Por padrao usa dias a partir de 1995-01-01 ate o ultimo dia disponivel
    de cada pais (date_to=None). Para incluir historico anterior, passe
    date_from mais antigo (ex.: 1980-01-01).
    """
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pandas e necessario. pip install pandas openpyxl") from exc

    if date_from is None:
        date_from = DEFAULT_DATE_FROM
    if len(date_from) >= 10:
        date_from = date_from[:10]
    if date_to is not None and len(date_to) >= 10:
        date_to = date_to[:10]

    wanted: set[str] | None = None
    if areas:
        wanted = {providers.resolve_area(a)["code"] for a in re.split(r"[\s,;+|]+", areas) if a}

    local = providers.find_local_flat_csv(csv_path) if prefer_local or csv_path else None
    if local is None and prefer_local:
        # Auto-download flat zip next to output for ContAgil reuse.
        dest = Path(out_path).resolve().parent
        providers.download_flat_csv(dest)
        local = providers.find_local_flat_csv(dest / providers.FLAT_CSV_NAME)

    if local is not None:
        series, names = _load_daily_from_flat(local, wanted, date_from, date_to)
        source = f"local:{local}"
    else:
        if wanted is None:
            # Sensible default set if online-only.
            wanted = {
                "BR",
                "US",
                "XM",
                "GB",
                "JP",
                "CN",
                "AR",
                "MX",
                "CL",
                "CO",
                "CA",
                "AU",
                "CH",
                "SE",
                "NO",
                "KR",
                "IN",
                "ZA",
                "TR",
            }
        series, names = _load_daily_from_sdmx(sorted(wanted), date_from, date_to)
        source = "sdmx"

    if not series:
        raise RuntimeError("Nenhuma serie diaria encontrada para os filtros informados.")

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    # Indice first.
    indice_rows = []
    country_tables: dict[str, list[dict[str, Any]]] = {}
    for code in sorted(series.keys()):
        points = series[code]
        if not points:
            continue
        table = build_country_rows(points)
        country_tables[code] = table
        last_aa = table[-1]["_taxa_aa"]
        indice_rows.append(
            {
                "Codigo": code,
                "Pais": names.get(code, AREA_NAMES.get(code, code)),
                "Aba": sheet_name(code, names.get(code, AREA_NAMES.get(code, code))),
                "N_dias": len(table),
                "Inicio": table[0]["Dia"],
                "Fim": table[-1]["Dia"],
                "Taxa_aa_ultimo_%": last_aa,
                "Taxa_ad_ultimo_%": table[-1]["Taxa (% a.d.)"],
                "Taxa_acumulada_final_%": table[-1]["Taxa acumulada (%)"],
            }
        )

    legenda = pd.DataFrame(
        [
            {"Item": "Fonte", "Valor": "BIS WS_CBPOL (frequencia diaria)"},
            {"Item": "Origem dados", "Valor": source},
            {
                "Item": "Periodo",
                "Valor": (
                    f"De {date_from} ate o ultimo dia disponivel de cada pais"
                    + (f" (limite --to {date_to})" if date_to else "")
                ),
            },
            {
                "Item": "Conversao % a.a. -> % a.d.",
                "Valor": "taxa_ad = (1 + taxa_aa/100)^(1/252) - 1   [ano com 252 dias uteis]",
            },
            {
                "Item": "Acumulacao (juros compostos)",
                "Valor": "fator *= (1 + taxa_ad);  taxa_acumulada_% = (fator - 1)*100",
            },
            {
                "Item": "Taxa acumulada mes (%)",
                "Valor": (
                    "Preenchida so no ultimo dia do mes da serie; "
                    "compostos apenas com as taxas a.d. daquele mes"
                ),
            },
            {
                "Item": "Taxa acumulada ano (%)",
                "Valor": (
                    "Preenchida so no ultimo dia do ano da serie; "
                    "compostos apenas com as taxas a.d. daquele ano"
                ),
            },
            {
                "Item": "Colunas por pais",
                "Valor": (
                    "Dia | Taxa (% a.d.) | Taxa acumulada (%) | "
                    "Taxa acumulada mes (%) | Taxa acumulada ano (%)"
                ),
            },
            {
                "Item": "Observacao",
                "Valor": (
                    "OBS_VALUE do BIS esta em % a.a. (policy rate). "
                    "A acumulacao comeca no primeiro dia >= "
                    f"{date_from} disponivel de cada pais. "
                    "Em series muito longas (hiperinflacao), a acumulada pode aparecer "
                    "em notacao cientifica (texto) por limite numerico do Excel."
                ),
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
                [None, "0.00000000", "0.000000", "0.000000", "0.000000"],
            )

        for code in sorted(country_tables.keys()):
            name = names.get(code, AREA_NAMES.get(code, code))
            aba = sheet_name(code, name)
            df = pd.DataFrame(country_tables[code])[
                [
                    "Dia",
                    "Taxa (% a.d.)",
                    "Taxa acumulada (%)",
                    "Taxa acumulada mês (%)",
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
        "date_from": date_from,
        "date_to": date_to,
    }
