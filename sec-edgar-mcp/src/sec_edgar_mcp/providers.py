"""SEC EDGAR / data.sec.gov client (no API key; User-Agent required)."""

from __future__ import annotations

import os
import re
import time
from functools import lru_cache
from typing import Any

import httpx

DATA_SEC = "https://data.sec.gov"
WWW_SEC = "https://www.sec.gov"
ARCHIVES = "https://www.sec.gov/Archives/edgar/data"

# SEC fair-access: identify app + contact email (required or 403).
DEFAULT_UA = os.getenv(
    "SEC_USER_AGENT",
    "SEC-data-analysys/0.1 (cafla19791@gmail.com)",
)

# Soft rate limit (~8 req/s to stay under 10/s).
_MIN_INTERVAL = 0.12
_last_request = 0.0


def _headers() -> dict[str, str]:
    return {
        "User-Agent": DEFAULT_UA,
        "Accept-Encoding": "gzip, deflate",
        "Accept": "application/json",
    }


def _throttle() -> None:
    global _last_request
    now = time.monotonic()
    wait = _MIN_INTERVAL - (now - _last_request)
    if wait > 0:
        time.sleep(wait)
    _last_request = time.monotonic()


def _get_json(url: str, *, timeout: float = 60.0) -> Any:
    _throttle()
    with httpx.Client(headers=_headers(), timeout=timeout, follow_redirects=True) as client:
        resp = client.get(url)
        if resp.status_code == 403:
            raise RuntimeError(
                "SEC retornou 403. Defina SEC_USER_AGENT='AppName email@dominio.com' "
                f"(atual: {DEFAULT_UA!r})."
            )
        resp.raise_for_status()
        return resp.json()


def pad_cik(cik: str | int) -> str:
    s = re.sub(r"\D", "", str(cik))
    if not s:
        raise ValueError(f"CIK invalido: {cik!r}")
    return s.zfill(10)


@lru_cache(maxsize=1)
def _ticker_map() -> dict[str, dict[str, Any]]:
    """ticker upper -> {cik, title} from company_tickers.json."""
    url = f"{WWW_SEC}/files/company_tickers.json"
    raw = _get_json(url)
    out: dict[str, dict[str, Any]] = {}
    # Shape: {"0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."}, ...}
    for row in raw.values() if isinstance(raw, dict) else raw:
        ticker = str(row.get("ticker", "")).upper().strip()
        if not ticker:
            continue
        out[ticker] = {
            "cik": pad_cik(row.get("cik_str", row.get("cik", ""))),
            "ticker": ticker,
            "title": row.get("title") or row.get("name") or "",
        }
    return out


def lookup_ticker(ticker: str) -> dict[str, Any]:
    """Resolve ticker -> CIK via SEC company_tickers.json."""
    t = ticker.strip().upper()
    info = _ticker_map().get(t)
    if info is None:
        # Fuzzy contains
        hits = [v for k, v in _ticker_map().items() if t in k or t in v["title"].upper()]
        return {
            "query": ticker,
            "found": False,
            "matches": hits[:15],
            "provider": "sec.gov/files/company_tickers.json",
        }
    return {
        "query": ticker,
        "found": True,
        **info,
        "provider": "sec.gov/files/company_tickers.json",
    }


def resolve_cik(ticker_or_cik: str) -> str:
    s = ticker_or_cik.strip()
    if re.fullmatch(r"\d{1,10}", s):
        return pad_cik(s)
    info = lookup_ticker(s)
    if not info.get("found"):
        raise ValueError(
            f"Ticker/CIK nao encontrado: {ticker_or_cik!r}. "
            "Use lookup_ticker ou o CIK numerico (ex.: 320193)."
        )
    return str(info["cik"])


def get_submissions(ticker_or_cik: str) -> dict[str, Any]:
    """Company metadata + recent filing history."""
    cik = resolve_cik(ticker_or_cik)
    data = _get_json(f"{DATA_SEC}/submissions/CIK{cik}.json")
    return {
        "cik": cik,
        "name": data.get("name"),
        "tickers": data.get("tickers"),
        "exchanges": data.get("exchanges"),
        "sic": data.get("sic"),
        "sicDescription": data.get("sicDescription"),
        "ein": data.get("ein"),
        "entityType": data.get("entityType"),
        "raw_keys": sorted(data.keys()),
        "provider": "data.sec.gov/submissions",
        "_data": data,
    }


def list_filings(
    ticker_or_cik: str,
    form: str | None = None,
    limit: int = 20,
) -> dict[str, Any]:
    """List recent filings (optionally filter by form, e.g. 10-K, 10-Q, 8-K, 20-F)."""
    limit = max(1, min(int(limit), 200))
    meta = get_submissions(ticker_or_cik)
    data = meta.pop("_data")
    recent = data.get("filings", {}).get("recent", {})
    forms = recent.get("form", [])
    dates = recent.get("filingDate", [])
    accessions = recent.get("accessionNumber", [])
    primaries = recent.get("primaryDocument", [])
    descriptions = recent.get("primaryDocDescription", [])

    form_filter = form.strip().upper() if form else None
    rows: list[dict[str, Any]] = []
    cik_int = str(int(meta["cik"]))  # archives path drops leading zeros

    for i, f in enumerate(forms):
        if form_filter and str(f).upper() != form_filter:
            continue
        acc = accessions[i] if i < len(accessions) else ""
        acc_nodash = acc.replace("-", "")
        primary = primaries[i] if i < len(primaries) else ""
        doc_url = ""
        if acc_nodash and primary:
            doc_url = f"{ARCHIVES}/{cik_int}/{acc_nodash}/{primary}"
        index_url = (
            f"{ARCHIVES}/{cik_int}/{acc_nodash}/{acc}-index.html" if acc_nodash else ""
        )
        rows.append(
            {
                "form": f,
                "filingDate": dates[i] if i < len(dates) else None,
                "accessionNumber": acc,
                "primaryDocument": primary,
                "description": descriptions[i] if i < len(descriptions) else None,
                "documentUrl": doc_url,
                "indexUrl": index_url,
            }
        )
        if len(rows) >= limit:
            break

    return {
        "cik": meta["cik"],
        "name": meta["name"],
        "tickers": meta.get("tickers"),
        "form_filter": form_filter,
        "count": len(rows),
        "filings": rows,
        "provider": "data.sec.gov/submissions",
    }


def get_company_facts(
    ticker_or_cik: str,
    concepts: list[str] | None = None,
    limit_per_concept: int = 8,
) -> dict[str, Any]:
    """XBRL company facts (us-gaap). Default: key P&L / balance concepts."""
    cik = resolve_cik(ticker_or_cik)
    data = _get_json(f"{DATA_SEC}/api/xbrl/companyfacts/CIK{cik}.json")
    facts = data.get("facts", {})
    us_gaap = facts.get("us-gaap", {})
    dei = facts.get("dei", {})

    default_concepts = concepts or [
        "Revenues",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "NetIncomeLoss",
        "Assets",
        "Liabilities",
        "StockholdersEquity",
        "EarningsPerShareDiluted",
        "OperatingIncomeLoss",
    ]
    limit_per_concept = max(1, min(int(limit_per_concept), 40))

    extracted: dict[str, Any] = {}
    for concept in default_concepts:
        node = us_gaap.get(concept) or dei.get(concept)
        if not node:
            continue
        units = node.get("units", {})
        # Prefer USD / USD/shares
        series = None
        unit_name = None
        for uname in ("USD", "USD/shares", "shares"):
            if uname in units:
                series = units[uname]
                unit_name = uname
                break
        if series is None and units:
            unit_name = next(iter(units))
            series = units[unit_name]
        if not series:
            continue
        # Sort by end date descending
        ordered = sorted(series, key=lambda x: x.get("end") or "", reverse=True)
        extracted[concept] = {
            "label": node.get("label"),
            "description": (node.get("description") or "")[:240],
            "unit": unit_name,
            "recent": [
                {
                    "end": r.get("end"),
                    "val": r.get("val"),
                    "fy": r.get("fy"),
                    "fp": r.get("fp"),
                    "form": r.get("form"),
                    "filed": r.get("filed"),
                    "frame": r.get("frame"),
                }
                for r in ordered[:limit_per_concept]
            ],
        }

    return {
        "cik": cik,
        "entityName": data.get("entityName"),
        "concepts_found": list(extracted.keys()),
        "concepts_requested": default_concepts,
        "facts": extracted,
        "provider": "data.sec.gov/api/xbrl/companyfacts",
        "note": (
            "Valores XBRL estruturados dos formularios 10-K/10-Q etc. "
            "Nao substitui a leitura do PDF/HTML completo do filing."
        ),
    }


def get_concept(
    ticker_or_cik: str,
    concept: str,
    taxonomy: str = "us-gaap",
    limit: int = 20,
) -> dict[str, Any]:
    """Single XBRL concept time series."""
    cik = resolve_cik(ticker_or_cik)
    concept = concept.strip()
    taxonomy = taxonomy.strip()
    url = f"{DATA_SEC}/api/xbrl/companyconcept/CIK{cik}/{taxonomy}/{concept}.json"
    data = _get_json(url)
    units = data.get("units", {})
    unit_name = "USD" if "USD" in units else (next(iter(units), None))
    series = units.get(unit_name, []) if unit_name else []
    ordered = sorted(series, key=lambda x: x.get("end") or "", reverse=True)
    limit = max(1, min(int(limit), 100))
    return {
        "cik": cik,
        "entityName": data.get("entityName"),
        "taxonomy": taxonomy,
        "concept": concept,
        "label": data.get("label"),
        "unit": unit_name,
        "count": len(ordered),
        "recent": ordered[:limit],
        "provider": "data.sec.gov/api/xbrl/companyconcept",
    }
