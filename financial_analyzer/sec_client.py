"""Cliente para a API pública SEC EDGAR (CompanyFacts e tickers)."""

from __future__ import annotations

import time
from typing import Any

import requests

SEC_DATA_BASE = "https://data.sec.gov"
TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"

# Conceitos US-GAAP usados com frequência em 10-K / 10-Q
METRIC_CONCEPTS: dict[str, list[str]] = {
    "revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
        "SalesRevenueGoodsNet",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
    ],
    "net_income": [
        "NetIncomeLoss",
        "ProfitLoss",
        "NetIncomeLossAvailableToCommonStockholdersBasic",
    ],
    "total_assets": ["Assets"],
    "total_liabilities": [
        "Liabilities",
        "LiabilitiesAndStockholdersEquity",
    ],
    "equity": [
        "StockholdersEquity",
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
        "PartnersCapital",
    ],
    "operating_income": [
        "OperatingIncomeLoss",
    ],
    "current_assets": ["AssetsCurrent"],
    "current_liabilities": ["LiabilitiesCurrent"],
    "cash": [
        "CashAndCashEquivalentsAtCarryingValue",
        "Cash",
        "CashCashEquivalentsAndShortTermInvestments",
    ],
}


class SecClient:
    """Busca fatos financeiros XBRL na SEC EDGAR."""

    def __init__(
        self,
        user_agent: str = "FinancialAnalyzer research@example.com",
        min_interval: float = 0.12,
    ) -> None:
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate",
        }
        self.min_interval = min_interval
        self._last_request = 0.0
        self._ticker_map: dict[str, dict[str, Any]] | None = None

    def _throttle(self) -> None:
        elapsed = time.monotonic() - self._last_request
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self._last_request = time.monotonic()

    def _get_json(self, url: str) -> dict[str, Any]:
        self._throttle()
        response = requests.get(url, headers=self.headers, timeout=30)
        response.raise_for_status()
        return response.json()

    def load_tickers(self) -> dict[str, dict[str, Any]]:
        """Mapa ticker → {cik, title}."""
        if self._ticker_map is not None:
            return self._ticker_map

        raw = self._get_json(TICKERS_URL)
        mapping: dict[str, dict[str, Any]] = {}
        for entry in raw.values():
            ticker = str(entry.get("ticker", "")).upper()
            if not ticker:
                continue
            mapping[ticker] = {
                "cik": int(entry["cik_str"]),
                "title": entry.get("title", ticker),
            }
        self._ticker_map = mapping
        return mapping

    def resolve_ticker(self, ticker: str) -> dict[str, Any]:
        ticker = ticker.upper().strip()
        mapping = self.load_tickers()
        if ticker not in mapping:
            raise ValueError(
                f"Ticker '{ticker}' não encontrado na lista da SEC. "
                "Verifique o símbolo ou use --cik."
            )
        return mapping[ticker]

    def get_company_facts(self, cik: int | str) -> dict[str, Any]:
        cik_padded = str(cik).zfill(10)
        url = f"{SEC_DATA_BASE}/api/xbrl/companyfacts/CIK{cik_padded}.json"
        return self._get_json(url)

    def extract_annual_series(
        self,
        facts: dict[str, Any],
        concepts: list[str],
        years: int = 5,
    ) -> list[dict[str, Any]]:
        """Extrai série anual (10-K / FY) para o primeiro conceito disponível."""
        us_gaap = facts.get("facts", {}).get("us-gaap", {})

        for concept in concepts:
            if concept not in us_gaap:
                continue

            units = us_gaap[concept].get("units", {})
            usd = units.get("USD") or units.get("USD/shares") or []
            if not usd:
                # algumas métricas usam outras unidades; pega a primeira lista
                for values in units.values():
                    if values:
                        usd = values
                        break

            annual = [
                e
                for e in usd
                if e.get("form") in {"10-K", "10-K/A"} and e.get("fp") == "FY"
            ]
            if not annual:
                # fallback: frames anuais (CY####)
                annual = [
                    e
                    for e in usd
                    if isinstance(e.get("frame"), str)
                    and e["frame"].startswith("CY")
                    and "Q" not in e["frame"]
                ]

            if not annual:
                continue

            by_end: dict[str, dict[str, Any]] = {}
            for entry in annual:
                end = entry.get("end")
                if not end:
                    continue
                prev = by_end.get(end)
                if prev is None or entry.get("filed", "") > prev.get("filed", ""):
                    by_end[end] = {
                        "end": end,
                        "value": float(entry["val"]),
                        "filed": entry.get("filed"),
                        "concept": concept,
                        "form": entry.get("form"),
                    }

            series = sorted(by_end.values(), key=lambda x: x["end"])
            return series[-years:]

        return []

    def get_financial_snapshot(
        self,
        ticker: str | None = None,
        cik: int | str | None = None,
        years: int = 5,
    ) -> dict[str, Any]:
        """Retorna nome, CIK e séries anuais das métricas principais."""
        if cik is None:
            if not ticker:
                raise ValueError("Informe ticker ou CIK.")
            info = self.resolve_ticker(ticker)
            cik = info["cik"]
            title = info["title"]
            ticker_sym = ticker.upper()
        else:
            ticker_sym = (ticker or "").upper()
            title = None

        facts = self.get_company_facts(cik)
        entity = facts.get("entityName") or title or ticker_sym or str(cik)

        series: dict[str, list[dict[str, Any]]] = {}
        for key, concepts in METRIC_CONCEPTS.items():
            series[key] = self.extract_annual_series(facts, concepts, years=years)

        return {
            "ticker": ticker_sym,
            "cik": int(cik),
            "name": entity,
            "series": series,
        }
