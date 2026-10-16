#!/usr/bin/env python3
"""Gera planilha com Meta Selic (BCB/Copom) e taxas básicas por país (BIS).

Fontes:
- Meta Selic / vigência: https://www.bcb.gov.br/controleinflacao/historicotaxasjuros
  API: https://www.bcb.gov.br/api/servico/sitebcb/historicotaxasjuros
- Taxas diárias e mensais por país: BIS WS_CBPOL
  https://data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip

Aba por país: colunas diárias e mensais a partir de 01/01/1996.
"""

from __future__ import annotations

import argparse
import json
import re
from datetime import date, datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import requests

START = date(1996, 1, 1)
COPOM_API = "https://www.bcb.gov.br/api/servico/sitebcb/historicotaxasjuros"
BIS_ZIP_URL = "https://data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip"
TZ_BR = ZoneInfo("America/Sao_Paulo")

COUNTRY_ORDER = [
    "Argentina",
    "Australia",
    "Austria",
    "Belgium",
    "Brazil",
    "Canada",
    "Chile",
    "China",
    "Colombia",
    "Croatia",
    "Czechia",
    "Denmark",
    "Euro area",
    "France",
    "Germany",
    "Greece",
    "Hong Kong SAR",
    "Hungary",
    "Iceland",
    "India",
    "Indonesia",
    "Israel",
    "Italy",
    "Japan",
    "Korea",
    "Kuwait",
    "Malaysia",
    "Mexico",
    "Morocco",
    "Netherlands",
    "New Zealand",
    "North Macedonia",
    "Norway",
    "Peru",
    "Philippines",
    "Poland",
    "Portugal",
    "Romania",
    "Russia",
    "Saudi Arabia",
    "Serbia",
    "South Africa",
    "Spain",
    "Sweden",
    "Switzerland",
    "Thailand",
    "Türkiye",
    "United Kingdom",
    "United States",
]


def parse_br_date(value: str | None) -> date | None:
    if not value:
        return None
    # API returns ISO with Z; interpret in America/Sao_Paulo (as BCB page)
    dt = datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(TZ_BR)
    return dt.date()


def sheet_name(country: str) -> str:
    # Excel sheet names: max 31 chars, no \ / ? * [ ]
    name = country.replace("Türkiye", "Turkiye")
    name = re.sub(r"[\\/?*\[\]]", "", name)
    name = name.replace(" ", "_")
    return name[:31]


def fetch_copom(cache: Path) -> list[dict]:
    if cache.exists():
        data = json.loads(cache.read_text(encoding="utf-8"))
        return data["conteudo"]
    r = requests.get(COPOM_API, timeout=60, headers={"Accept": "application/json"})
    r.raise_for_status()
    cache.parent.mkdir(parents=True, exist_ok=True)
    cache.write_text(r.text, encoding="utf-8")
    return r.json()["conteudo"]


def build_meta_selic(rows: list[dict]) -> pd.DataFrame:
    records = []
    for r in rows:
        inicio = parse_br_date(r.get("DataInicioVigencia"))
        if inicio is None or inicio < date(1996, 7, 1):
            continue
        fim = parse_br_date(r.get("DataFimVigencia"))
        reuniao = parse_br_date(r.get("DataReuniaoCopom"))
        records.append(
            {
                "N_Reuniao_Copom": int(r["NumeroReuniaoCopom"]),
                "Reuniao_Extraordinaria": bool(r.get("ReuniaoExtraordinaria")),
                "Data_Reuniao": reuniao,
                "Inicio_Vigencia": inicio,
                "Fim_Vigencia": fim,
                "Meta_Selic": r.get("MetaSelic"),
                "TBAN": r.get("TaxaTban"),
                "Selic_Efetiva_Vigencia": r.get("TaxaSelicEfetivaVigencia"),
                "Selic_Efetiva_Anualizada_%_aa": r.get("TaxaSelicEfetivaAnualizada"),
                "Vies": r.get("Vies"),
                "Uso_Meta_Selic": r.get("UsoMetaSelic"),
                "Decisao_Monocratica_Pres": r.get("descisaoMonocraticaPres"),
            }
        )
    df = pd.DataFrame(records)
    df = df.sort_values(
        ["Inicio_Vigencia", "N_Reuniao_Copom"], ascending=[True, True]
    ).reset_index(drop=True)
    return df


def ensure_bis_csv(bis_dir: Path) -> Path:
    csv_path = bis_dir / "WS_CBPOL_csv_flat.csv"
    if csv_path.exists() and csv_path.stat().st_size > 1_000_000:
        return csv_path
    bis_dir.mkdir(parents=True, exist_ok=True)
    zip_path = bis_dir / "WS_CBPOL_csv_flat.zip"
    print(f"Baixando BIS WS_CBPOL de {BIS_ZIP_URL} ...")
    with requests.get(BIS_ZIP_URL, stream=True, timeout=180) as r:
        r.raise_for_status()
        with open(zip_path, "wb") as f:
            for chunk in r.iter_content(1 << 20):
                if chunk:
                    f.write(chunk)
    import zipfile

    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(bis_dir)
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV BIS não encontrado em {bis_dir}")
    return csv_path


def load_bis_from_1996(csv_path: Path) -> pd.DataFrame:
    usecols = [
        "FREQ:Frequency",
        "REF_AREA:Reference area",
        "TIME_PERIOD:Time period or range",
        "OBS_VALUE:Observation Value",
        "SOURCE_REF:Publication Source",
        "COMPILATION:Compilation",
        "SUPP_INFO_BREAKS:Supplemental information and breaks",
        "TITLE:Title",
    ]
    parts: list[pd.DataFrame] = []
    for chunk in pd.read_csv(
        csv_path, usecols=usecols, chunksize=150_000, low_memory=False
    ):
        chunk = chunk.rename(
            columns={
                "FREQ:Frequency": "freq",
                "REF_AREA:Reference area": "area",
                "TIME_PERIOD:Time period or range": "period",
                "OBS_VALUE:Observation Value": "value",
                "SOURCE_REF:Publication Source": "source",
                "COMPILATION:Compilation": "compilation",
                "SUPP_INFO_BREAKS:Supplemental information and breaks": "breaks",
                "TITLE:Title": "title",
            }
        )
        chunk["period"] = chunk["period"].astype(str)
        chunk = chunk[chunk["period"] >= "1996-01"]
        if not chunk.empty:
            parts.append(chunk)
    if not parts:
        raise RuntimeError("Nenhum dado BIS >= 1996-01 encontrado")
    df = pd.concat(parts, ignore_index=True)
    # area like "BR: Brazil"
    df["country"] = df["area"].str.split(": ", n=1).str[-1].str.strip()
    df["freq_code"] = df["freq"].str.split(": ", n=1).str[0].str.strip()
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    return df


def country_frames(df: pd.DataFrame, country: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    sub = df[df["country"] == country]
    daily = (
        sub[sub["freq_code"] == "D"][["period", "value"]]
        .dropna(subset=["value"])
        .drop_duplicates(subset=["period"])
        .sort_values("period")
        .reset_index(drop=True)
    )
    daily.columns = ["Data", "Taxa_Diaria_%_aa"]
    monthly = (
        sub[sub["freq_code"] == "M"][["period", "value"]]
        .dropna(subset=["value"])
        .drop_duplicates(subset=["period"])
        .sort_values("period")
        .reset_index(drop=True)
    )
    monthly.columns = ["Ano_Mes", "Taxa_Mensal_%_aa"]
    meta_row = sub.iloc[0] if len(sub) else None
    meta = {
        "country": country,
        "source": meta_row["source"] if meta_row is not None else "",
        "compilation": meta_row["compilation"] if meta_row is not None else "",
        "breaks": meta_row["breaks"] if meta_row is not None else "",
        "title": meta_row["title"] if meta_row is not None else "",
    }
    return daily, monthly, meta


def write_workbook(
    out_path: Path,
    meta_selic: pd.DataFrame,
    bis: pd.DataFrame,
) -> None:
    countries = sorted(bis["country"].unique(), key=lambda c: (COUNTRY_ORDER.index(c) if c in COUNTRY_ORDER else 999, c))
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        # Índice
        idx_rows = [
            {
                "Aba": "Meta_Selic_BCB",
                "Conteudo": "Meta Selic e períodos de vigência (Copom/BCB) desde 01/07/1996",
                "Fonte": "https://www.bcb.gov.br/controleinflacao/historicotaxasjuros",
            }
        ]
        for c in countries:
            idx_rows.append(
                {
                    "Aba": sheet_name(c),
                    "Conteudo": f"Taxas diárias e mensais — {c} a partir de 01/01/1996",
                    "Fonte": "BIS WS_CBPOL (Central bank policy rates)",
                }
            )
        idx = pd.DataFrame(idx_rows)
        idx.to_excel(writer, sheet_name="Indice", index=False, startrow=4)
        wb = writer.book
        ws = writer.sheets["Indice"]
        title_fmt = wb.add_format({"bold": True, "font_size": 14})
        note_fmt = wb.add_format({"text_wrap": True, "valign": "top"})
        header_fmt = wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "white"})
        ws.write(0, 0, "Taxas básicas de juros — Meta Selic (BCB) e bancos centrais (BIS)", title_fmt)
        ws.write(
            1,
            0,
            "Gerado a partir do histórico Copom/BCB e do dataset BIS WS_CBPOL. "
            "A planilha original em formato colunar do Drive estoura o limite de colunas do Excel "
            "(corta em 1988); aqui usamos o CSV flat completo do BIS, filtrado a partir de 01/01/1996.",
            note_fmt,
        )
        ws.set_row(1, 45)
        ws.write(2, 0, f"Gerado em {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")
        for col, width in enumerate([24, 70, 70]):
            ws.set_column(col, col, width)
        for col_num, value in enumerate(idx.columns.values):
            ws.write(4, col_num, value, header_fmt)

        # Meta Selic
        meta_out = meta_selic.copy()
        for col in ["Data_Reuniao", "Inicio_Vigencia", "Fim_Vigencia"]:
            meta_out[col] = pd.to_datetime(meta_out[col])
        meta_out.to_excel(writer, sheet_name="Meta_Selic_BCB", index=False, startrow=3)
        ws_m = writer.sheets["Meta_Selic_BCB"]
        ws_m.write(0, 0, "Meta Selic definida pelo Copom — períodos de vigência", title_fmt)
        ws_m.write(
            1,
            0,
            "Fonte: BCB — Histórico de taxas de juros. Até ~1997/98 o campo Meta_Selic reflete a meta "
            "mensal publicada na época; a partir de 1998/99 a Meta Selic é em % a.a. "
            "Use Selic_Efetiva_Anualizada_%_aa para a taxa efetiva anualizada na vigência. "
            "Fim_Vigencia vazio = vigência corrente.",
            note_fmt,
        )
        ws_m.set_row(1, 50)
        for col_num, value in enumerate(meta_out.columns.values):
            ws_m.write(3, col_num, value, header_fmt)
        ws_m.set_column(0, 0, 16)
        ws_m.set_column(1, 1, 18)
        ws_m.set_column(2, 4, 14)
        ws_m.set_column(5, 8, 18)
        ws_m.set_column(9, 11, 16)

        # Países
        pct_fmt = wb.add_format({"num_format": "0.0000"})
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
        used_names: set[str] = set()
        for country in countries:
            daily, monthly, meta = country_frames(bis, country)
            # side-by-side without forcing equal length
            n = max(len(daily), len(monthly))
            out = pd.DataFrame(
                {
                    "Data_Diaria": list(daily["Data"]) + [None] * (n - len(daily)),
                    "Taxa_Diaria_%_aa": list(daily["Taxa_Diaria_%_aa"]) + [None] * (n - len(daily)),
                    "Ano_Mes": list(monthly["Ano_Mes"]) + [None] * (n - len(monthly)),
                    "Taxa_Mensal_%_aa": list(monthly["Taxa_Mensal_%_aa"]) + [None] * (n - len(monthly)),
                }
            )
            name = sheet_name(country)
            base = name
            i = 2
            while name in used_names:
                suffix = f"_{i}"
                name = (base[: 31 - len(suffix)] + suffix)
                i += 1
            used_names.add(name)

            out.to_excel(writer, sheet_name=name, index=False, startrow=5)
            ws_c = writer.sheets[name]
            ws_c.write(0, 0, f"Taxa básica — {country}", title_fmt)
            ws_c.write(1, 0, f"Fonte: {meta['source']}")
            compilation = str(meta["compilation"] or "")[:900]
            breaks = str(meta["breaks"] or "")[:500]
            ws_c.write(2, 0, f"Compilação: {compilation}", note_fmt)
            ws_c.write(3, 0, f"Quebras/notas: {breaks}" if breaks and breaks != "nan" else "Quebras/notas: —", note_fmt)
            ws_c.set_row(2, 40)
            ws_c.set_row(3, 30)
            for col_num, value in enumerate(out.columns.values):
                ws_c.write(5, col_num, value, header_fmt)
            ws_c.set_column(0, 0, 14, date_fmt)
            ws_c.set_column(1, 1, 16, pct_fmt)
            ws_c.set_column(2, 2, 12)
            ws_c.set_column(3, 3, 16, pct_fmt)
            ws_c.write(0, 5, f"Obs. diárias: {len(daily)}")
            ws_c.write(1, 5, f"Obs. mensais: {len(monthly)}")

        # Fontes
        fontes = pd.DataFrame(
            [
                {
                    "Item": "Meta Selic / vigência",
                    "URL": "https://www.bcb.gov.br/controleinflacao/historicotaxasjuros",
                    "API": COPOM_API,
                },
                {
                    "Item": "Taxas bancos centrais (diária/mensal)",
                    "URL": "https://data.bis.org/topics/CBPOL/data",
                    "API": BIS_ZIP_URL,
                },
                {
                    "Item": "Série SGS 432 (Meta Selic diária, desde 05/03/1999)",
                    "URL": "https://dadosabertos.bcb.gov.br/dataset/432-taxa-de-juros---meta-selic-definida-pelo-copom",
                    "API": "https://api.bcb.gov.br/dados/serie/bcdata.sgs.432/dados?formato=json",
                },
            ]
        )
        fontes.to_excel(writer, sheet_name="Fontes", index=False)
        ws_f = writer.sheets["Fontes"]
        for col, width in enumerate([45, 80, 80]):
            ws_f.set_column(col, col, width)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--output",
        type=Path,
        default=Path("output/taxas_juros_bc_paises_1996.xlsx"),
    )
    ap.add_argument("--bis-dir", type=Path, default=Path("data/bis"))
    ap.add_argument("--copom-cache", type=Path, default=Path("data/selic/historico_copom.json"))
    args = ap.parse_args()

    print("1/3 Meta Selic (Copom/BCB)...")
    copom = fetch_copom(args.copom_cache)
    meta = build_meta_selic(copom)
    print(f"   {len(meta)} períodos desde {meta['Inicio_Vigencia'].min()} até {meta['Inicio_Vigencia'].max()}")

    print("2/3 Taxas BIS por país...")
    csv_path = ensure_bis_csv(args.bis_dir)
    bis = load_bis_from_1996(csv_path)
    print(f"   {bis['country'].nunique()} países, {len(bis):,} observações >= 1996-01")

    print("3/3 Gravando Excel...")
    write_workbook(args.output, meta, bis)
    print(f"OK: {args.output} ({args.output.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
