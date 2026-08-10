#!/usr/bin/env python3
"""Fetch World Bank featured indicators and write one Excel sheet per indicator."""

from __future__ import annotations

import argparse
import io
import re
import time
import zipfile
from pathlib import Path

import pandas as pd
import requests

# Featured indicators from World Bank screenshots, grouped by category.
INDICATORS_GROWTH_MACRO: list[tuple[str, str, str]] = [
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

INDICATORS_BUSINESS_INFRA: list[tuple[str, str, str]] = [
    # Business environment
    ("Business environment", "Time required to start a business (days)", "IC.REG.DURS"),
    ("Business environment", "Time required to get electricity (days)", "IC.ELC.TIME"),
    (
        "Business environment",
        "Firms expected to give gifts in meetings with tax officials (% of firms)",
        "IC.TAX.GIFT.ZS",
    ),
    (
        "Business environment",
        "Firms with female top manager (% of firms)",
        "IC.FRM.FEMM.ZS",
    ),
    # Financial access and stability
    (
        "Financial access and stability",
        "Depositors with commercial banks (per 1,000 adults)",
        "FB.CBK.DPTR.P3",
    ),
    (
        "Financial access and stability",
        "Borrowers from commercial banks (per 1,000 adults)",
        "FB.CBK.BRWR.P3",
    ),
    (
        "Financial access and stability",
        "Commercial bank branches (per 100,000 adults)",
        "FB.CBK.BRCH.P5",
    ),
    (
        "Financial access and stability",
        "Bank nonperforming loans to total gross loans (%)",
        "FB.AST.NPER.ZS",
    ),
    # Stock markets
    (
        "Stock markets",
        "Market capitalization of listed domestic companies (% of GDP)",
        "CM.MKT.LCAP.GD.ZS",
    ),
    (
        "Stock markets",
        "Stocks traded, turnover ratio of domestic shares (%)",
        "CM.MKT.TRNR",
    ),
    # Government finance and taxes
    (
        "Government finance and taxes",
        "Revenue, excluding grants (current LCU)",
        "GC.REV.XGRT.CN",
    ),
    ("Government finance and taxes", "Expense (current LCU)", "GC.XPN.TOTL.CN"),
    (
        "Government finance and taxes",
        "Net lending (+) / net borrowing (-) (current LCU)",
        "GC.NLD.TOTL.CN",
    ),
    (
        "Government finance and taxes",
        "Compensation of employees (current LCU)",
        "GC.XPN.COMP.CN",
    ),
    (
        "Government finance and taxes",
        "Taxes on goods and services (current LCU)",
        "GC.TAX.GSRV.CN",
    ),
    (
        "Government finance and taxes",
        "Profit tax (% of commercial profits)",
        "IC.TAX.PRFT.CP.ZS",
    ),
    (
        "Government finance and taxes",
        "Total tax rate (% of commercial profits)",
        "IC.TAX.TOTL.CP.ZS",
    ),
    # Military and fragile situations
    (
        "Military and fragile situations",
        "Military expenditure (% of GDP)",
        "MS.MIL.XPND.GD.ZS",
    ),
    (
        "Military and fragile situations",
        "Armed forces personnel, total",
        "MS.MIL.TOTL.P1",
    ),
    (
        "Military and fragile situations",
        "Battle-related deaths (number of people)",
        "VC.BTL.DETH",
    ),
    (
        "Military and fragile situations",
        "Intentional homicides (per 100,000 people)",
        "VC.IHR.PSRC.P5",
    ),
    # Infrastructure and communications
    (
        "Infrastructure and communications",
        "Air transport, passengers carried",
        "IS.AIR.PSGR",
    ),
    (
        "Infrastructure and communications",
        "Air transport, freight (million ton-km)",
        "IS.AIR.GOOD.MT.K1",
    ),
    (
        "Infrastructure and communications",
        "Container port traffic (TEU: 20 foot equivalent units)",
        "IS.SHP.GOOD.TU",
    ),
    (
        "Infrastructure and communications",
        "Individuals using the Internet (% of population)",
        "IT.NET.USER.ZS",
    ),
    (
        "Infrastructure and communications",
        "Mobile cellular subscriptions (per 100 people)",
        "IT.CEL.SETS.P2",
    ),
    (
        "Infrastructure and communications",
        "Investment in transport with private participation (current US$)",
        "IE.PPI.TRAN.CD",
    ),
    (
        "Infrastructure and communications",
        "Investment in energy with private participation (current US$)",
        "IE.PPI.ENGY.CD",
    ),
    # Science and innovation
    (
        "Science and innovation",
        "Research and development expenditure (% of GDP)",
        "GB.XPD.RSDV.GD.ZS",
    ),
    ("Science and innovation", "Patent applications, residents", "IP.PAT.RESD"),
    (
        "Science and innovation",
        "Industrial design applications, resident, by count",
        "IP.IDS.RSCT",
    ),
    (
        "Science and innovation",
        "Scientific and technical journal articles",
        "IP.JRN.ARTC.SC",
    ),
    (
        "Science and innovation",
        "ICT goods exports (% of total goods exports)",
        "TX.VAL.ICTG.ZS.UN",
    ),
]

INDICATORS_POVERTY: list[tuple[str, str, str]] = [
    # Poverty rates at national poverty lines
    (
        "Poverty rates at national poverty lines",
        "Poverty headcount ratio at national poverty lines (% of population)",
        "SI.POV.NAHC",
    ),
    (
        "Poverty rates at national poverty lines",
        "Urban poverty headcount ratio at national poverty lines (% of urban population)",
        "SI.POV.URHC",
    ),
    (
        "Poverty rates at national poverty lines",
        "Rural poverty headcount ratio at national poverty lines (% of rural population)",
        "SI.POV.RUHC",
    ),
    (
        "Poverty rates at national poverty lines",
        "Poverty gap at national poverty lines (%)",
        "SI.POV.NAGP",
    ),
    (
        "Poverty rates at national poverty lines",
        "Urban poverty gap at national poverty lines (%)",
        "SI.POV.URGP",
    ),
    (
        "Poverty rates at national poverty lines",
        "Rural poverty gap at national poverty lines (%)",
        "SI.POV.RUGP",
    ),
    # Poverty rates at international poverty lines
    (
        "Poverty rates at international poverty lines",
        "Poverty headcount ratio at $3.00 a day (2021 PPP) (% of population)",
        "SI.POV.DDAY",
    ),
    (
        "Poverty rates at international poverty lines",
        "Poverty headcount ratio at $4.20 a day (2021 PPP) (% of population)",
        "SI.POV.LMIC",
    ),
    (
        "Poverty rates at international poverty lines",
        "Poverty headcount ratio at $8.30 a day (2021 PPP) (% of population)",
        "SI.POV.UMIC",
    ),
    (
        "Poverty rates at international poverty lines",
        "Poverty gap at $3.00 a day (2021 PPP) (%)",
        "SI.POV.GAPS",
    ),
    (
        "Poverty rates at international poverty lines",
        "Poverty gap at $4.20 a day (2021 PPP) (%)",
        "SI.POV.LMIC.GP",
    ),
    (
        "Poverty rates at international poverty lines",
        "Poverty gap at $8.30 a day (2021 PPP) (%)",
        "SI.POV.UMIC.GP",
    ),
    # Distribution of income or consumption
    (
        "Distribution of income or consumption",
        "GINI index (World Bank estimate)",
        "SI.POV.GINI",
    ),
    (
        "Distribution of income or consumption",
        "Income share held by lowest 10%",
        "SI.DST.FRST.10",
    ),
    (
        "Distribution of income or consumption",
        "Income share held by lowest 20%",
        "SI.DST.FRST.20",
    ),
    (
        "Distribution of income or consumption",
        "Income share held by second 20%",
        "SI.DST.02ND.20",
    ),
    (
        "Distribution of income or consumption",
        "Income share held by third 20%",
        "SI.DST.03RD.20",
    ),
    (
        "Distribution of income or consumption",
        "Income share held by fourth 20%",
        "SI.DST.04TH.20",
    ),
    (
        "Distribution of income or consumption",
        "Income share held by highest 20%",
        "SI.DST.05TH.20",
    ),
    (
        "Distribution of income or consumption",
        "Income share held by highest 10%",
        "SI.DST.10TH.10",
    ),
    # Shared prosperity
    (
        "Shared prosperity",
        "Annualized average growth rate in per capita real survey mean consumption or income, bottom 40% of population (%)",
        "SI.SPR.PC40.ZG",
    ),
    (
        "Shared prosperity",
        "Annualized average growth rate in per capita real survey mean consumption or income, total population (%)",
        "SI.SPR.PCAP.ZG",
    ),
    (
        "Shared prosperity",
        "Survey mean consumption or income per capita, bottom 40% of population (2021 PPP $ per day)",
        "SI.SPR.PC40",
    ),
    (
        "Shared prosperity",
        "Survey mean consumption or income per capita, total population (2021 PPP $ per day)",
        "SI.SPR.PCAP",
    ),
]

WORKBOOKS: dict[str, tuple[Path, list[tuple[str, str, str]]]] = {
    "growth-macro": (
        Path("world_bank_featured_indicators.xlsx"),
        INDICATORS_GROWTH_MACRO,
    ),
    "business-infra": (
        Path("world_bank_featured_indicators_business_infra.xlsx"),
        INDICATORS_BUSINESS_INFRA,
    ),
    "poverty": (
        Path("world_bank_featured_indicators_poverty.xlsx"),
        INDICATORS_POVERTY,
    ),
}

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

    for col in year_cols:
        frame[col] = pd.to_numeric(frame[col], errors="coerce")

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


def write_workbook(
    output_path: Path, indicators: list[tuple[str, str, str]]
) -> None:
    used_sheet_names: set[str] = {"Index", "Summary"}
    summary_rows: list[dict] = []

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        index_df = pd.DataFrame(
            [
                {"Category": category, "Indicator": name, "Code": code}
                for category, name, code in indicators
            ]
        )
        index_df.to_excel(writer, sheet_name="Index", index=False)

        for i, (category, name, code) in enumerate(indicators, start=1):
            print(f"[{i}/{len(indicators)}] Fetching {code} — {name}", flush=True)
            try:
                data = fetch_indicator(code)
                status = "ok"
                n_countries = (
                    int(data["Country Name"].nunique()) if not data.empty else 0
                )
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


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fetch World Bank featured indicators into Excel workbooks."
    )
    parser.add_argument(
        "--set",
        choices=["all", *WORKBOOKS.keys()],
        default="all",
        help="Which workbook set to regenerate (default: all).",
    )
    args = parser.parse_args()

    selected = (
        WORKBOOKS.items()
        if args.set == "all"
        else [(args.set, WORKBOOKS[args.set])]
    )
    for key, (path, indicators) in selected:
        print(f"\n=== Workbook '{key}' -> {path} ({len(indicators)} indicators) ===")
        write_workbook(path, indicators)


if __name__ == "__main__":
    main()
