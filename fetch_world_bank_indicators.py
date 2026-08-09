#!/usr/bin/env python3
"""Fetch World Bank featured indicators and write one Excel sheet per indicator."""

from __future__ import annotations

import io
import re
import time
import zipfile
from pathlib import Path

import pandas as pd
import requests

# Featured indicators from the World Bank screenshot, grouped by category.
INDICATORS: list[tuple[str, str, str]] = [
    # Growth and economic structure
    ("Growth and economic structure", "GDP (current US$)", "NY.GDP.MKTP.CD"),
    ("Growth and economic structure", "GDP growth (annual %)", "NY.GDP.MKTP.KD.ZG"),
    (
        "Growth and economic structure",
        "Agriculture, value added (annual % growth)",
        "NV.AGR.TOTL.KD.ZG",
    ),
    (
        "Growth and economic structure",
        "Industry, value added (annual % growth)",
        "NV.IND.TOTL.KD.ZG",
    ),
    (
        "Growth and economic structure",
        "Manufacturing, value added (annual % growth)",
        "NV.IND.MANF.KD.ZG",
    ),
    (
        "Growth and economic structure",
        "Services, value added (annual % growth)",
        "NV.SRV.TOTL.KD.ZG",
    ),
    (
        "Growth and economic structure",
        "Final consumption expenditure (annual % growth)",
        "NE.CON.TOTL.KD.ZG",
    ),
    (
        "Growth and economic structure",
        "Gross capital formation (annual % growth)",
        "NE.GDI.TOTL.KD.ZG",
    ),
    (
        "Growth and economic structure",
        "Exports of goods and services (annual % growth)",
        "NE.EXP.GNFS.KD.ZG",
    ),
    (
        "Growth and economic structure",
        "Imports of goods and services (annual % growth)",
        "NE.IMP.GNFS.KD.ZG",
    ),
    (
        "Growth and economic structure",
        "Agriculture, value added (% of GDP)",
        "NV.AGR.TOTL.ZS",
    ),
    (
        "Growth and economic structure",
        "Industry, value added (% of GDP)",
        "NV.IND.TOTL.ZS",
    ),
    (
        "Growth and economic structure",
        "Services, value added (% of GDP)",
        "NV.SRV.TOTL.ZS",
    ),
    (
        "Growth and economic structure",
        "Final consumption expenditure (% of GDP)",
        "NE.CON.TOTL.ZS",
    ),
    (
        "Growth and economic structure",
        "Gross capital formation (% of GDP)",
        "NE.GDI.TOTL.ZS",
    ),
    (
        "Growth and economic structure",
        "Exports of goods and services (% of GDP)",
        "NE.EXP.GNFS.ZS",
    ),
    (
        "Growth and economic structure",
        "Imports of goods and services (% of GDP)",
        "NE.IMP.GNFS.ZS",
    ),
    # Income and savings
    (
        "Income and savings",
        "GNI per capita, Atlas method (current US$)",
        "NY.GNP.PCAP.CD",
    ),
    (
        "Income and savings",
        "GNI per capita, PPP (current international $)",
        "NY.GNP.PCAP.PP.CD",
    ),
    ("Income and savings", "Population, total", "SP.POP.TOTL"),
    ("Income and savings", "Gross savings (% of GDP)", "NY.GNS.ICTR.ZS"),
    (
        "Income and savings",
        "Adjusted net savings, including particulate emission damage (% of GNI)",
        "NY.ADJ.SVNG.GN.ZS",
    ),
    # Balance of payments
    ("Balance of payments", "Export value index (2000 = 100)", "TX.VAL.MRCH.XD.WD"),
    ("Balance of payments", "Import value index (2000 = 100)", "TM.VAL.MRCH.XD.WD"),
    (
        "Balance of payments",
        "Personal remittances, received (% of GDP)",
        "BX.TRF.PWKR.DT.GD.ZS",
    ),
    (
        "Balance of payments",
        "Current account balance (% of GDP)",
        "BN.CAB.XOKA.GD.ZS",
    ),
    (
        "Balance of payments",
        "Foreign direct investment, net inflows (% of GDP)",
        "BX.KLT.DINV.WD.GD.ZS",
    ),
    # Prices and terms of trade
    ("Prices and terms of trade", "Consumer price index (2010 = 100)", "FP.CPI.TOTL"),
    (
        "Prices and terms of trade",
        "Export unit value index (2000 = 100)",
        "TX.UVI.MRCH.XD.WD",
    ),
    (
        "Prices and terms of trade",
        "Import unit value index (2000 = 100)",
        "TM.UVI.MRCH.XD.WD",
    ),
    (
        "Prices and terms of trade",
        "Net barter terms of trade index (2000 = 100)",
        "TT.PRI.MRCH.XD.WD",
    ),
    # Labor and productivity
    (
        "Labor and productivity",
        "GDP per person employed (constant 2011 PPP $)",
        "SL.GDP.PCAP.EM.KD",
    ),
    (
        "Labor and productivity",
        "Unemployment, total (% of total labor force) (modeled ILO estimate)",
        "SL.UEM.TOTL.ZS",
    ),
    (
        "Labor and productivity",
        "Agriculture, value added per worker (constant 2010 US$)",
        "NV.AGR.EMPL.KD",
    ),
    (
        "Labor and productivity",
        "Industry, value added per worker (constant 2010 US$)",
        "NV.IND.EMPL.KD",
    ),
    (
        "Labor and productivity",
        "Services, value added per worker (constant 2010 US$)",
        "NV.SRV.EMPL.KD",
    ),
]

DOWNLOAD_URL = "https://api.worldbank.org/v2/en/indicator/{code}"
SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "world-bank-indicators-fetcher/1.0"})


def sheet_name(code: str, used: set[str]) -> str:
    """Use indicator codes as sheet names (unique and <= 31 chars)."""
    base = re.sub(r"[\[\]\*:/\\?]", "_", code)[:31]
    candidate = base
    i = 2
    while candidate in used:
        suffix = f"_{i}"
        candidate = f"{base[: 31 - len(suffix)]}{suffix}"
        i += 1
    used.add(candidate)
    return candidate


def fetch_indicator(code: str) -> pd.DataFrame:
    """Download one indicator CSV (wide: countries x years) from World Bank."""
    last_error: Exception | None = None
    for attempt in range(1, 5):
        try:
            response = SESSION.get(
                DOWNLOAD_URL.format(code=code),
                params={"downloadformat": "csv"},
                timeout=180,
            )
            response.raise_for_status()
            with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
                data_name = next(
                    n
                    for n in zf.namelist()
                    if n.startswith("API_") and n.endswith(".csv")
                )
                with zf.open(data_name) as fh:
                    # World Bank CSVs have 4 metadata rows before the header.
                    frame = pd.read_csv(fh, skiprows=4)
            break
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            time.sleep(1.5 * attempt)
    else:
        raise RuntimeError(f"Failed to download {code}: {last_error}") from last_error

    # Drop trailing empty column that often appears in WDI exports.
    frame = frame.loc[:, ~frame.columns.str.match(r"^Unnamed")]
    keep = ["Country Name", "Country Code", "Indicator Name", "Indicator Code"]
    year_cols = [c for c in frame.columns if c.isdigit()]
    frame = frame[keep + year_cols].copy()

    # Convert year columns to numeric; keep empty cells as NaN.
    for col in year_cols:
        frame[col] = pd.to_numeric(frame[col], errors="coerce")

    # Keep rows with at least one observation.
    if year_cols:
        frame = frame.loc[frame[year_cols].notna().any(axis=1)].copy()

    return frame.sort_values(["Country Name", "Country Code"]).reset_index(drop=True)


def year_range(frame: pd.DataFrame) -> tuple[int | None, int | None]:
    year_cols = [c for c in frame.columns if str(c).isdigit()]
    if not year_cols:
        return None, None
    present = [int(c) for c in year_cols if frame[c].notna().any()]
    if not present:
        return None, None
    return min(present), max(present)


def main() -> None:
    output_path = Path("world_bank_featured_indicators.xlsx")
    used_sheet_names: set[str] = {"Index", "Summary"}
    summary_rows: list[dict] = []

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        index_df = pd.DataFrame(
            [
                {"Category": category, "Indicator": name, "Code": code}
                for category, name, code in INDICATORS
            ]
        )
        index_df.to_excel(writer, sheet_name="Index", index=False)

        for i, (category, name, code) in enumerate(INDICATORS, start=1):
            print(f"[{i}/{len(INDICATORS)}] Fetching {code} — {name}", flush=True)
            try:
                data = fetch_indicator(code)
                status = "ok"
                n_countries = int(data["Country Name"].nunique()) if not data.empty else 0
                year_min, year_max = year_range(data)
            except Exception as exc:  # noqa: BLE001
                print(f"  ERROR: {exc}", flush=True)
                data = pd.DataFrame(
                    [{"Error": str(exc), "Indicator": name, "Code": code}]
                )
                status = f"error: {exc}"
                n_countries = 0
                year_min = None
                year_max = None

            tab = sheet_name(code, used_sheet_names)
            data.to_excel(writer, sheet_name=tab, index=False)
            summary_rows.append(
                {
                    "Category": category,
                    "Indicator": name,
                    "Code": code,
                    "Sheet": tab,
                    "Status": status,
                    "Countries/aggregates": n_countries,
                    "Year min": year_min,
                    "Year max": year_max,
                }
            )
            print(
                f"  -> sheet '{tab}' | rows={n_countries} | years={year_min}-{year_max}",
                flush=True,
            )
            time.sleep(0.2)

        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

    size_mb = output_path.stat().st_size / (1024 * 1024)
    print(f"\nWrote {output_path} ({size_mb:.2f} MB)", flush=True)


if __name__ == "__main__":
    main()
