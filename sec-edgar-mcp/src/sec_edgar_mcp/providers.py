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


# Aliases comuns us-gaap -> ifrs-full (emissores estrangeiros: 20-F / 6-K).
_IFRS_ALIASES: dict[str, list[str]] = {
    "NetIncomeLoss": [
        "ProfitLossAttributableToOwnersOfParent",
        "ProfitLoss",
    ],
    "Revenues": ["Revenue"],
    "RevenueFromContractWithCustomerExcludingAssessedTax": ["Revenue"],
    "Assets": ["Assets"],
    "Liabilities": ["Liabilities"],
    "StockholdersEquity": ["Equity"],
    "OperatingIncomeLoss": ["ProfitLossFromOperatingActivities"],
    "EarningsPerShareDiluted": ["BasicEarningsLossPerShare", "DilutedEarningsLossPerShare"],
}


def _fetch_concept_raw(cik: str, taxonomy: str, concept: str) -> dict[str, Any]:
    url = f"{DATA_SEC}/api/xbrl/companyconcept/CIK{cik}/{taxonomy}/{concept}.json"
    return _get_json(url)


def _series_from_concept(data: dict[str, Any]) -> tuple[str | None, list[dict[str, Any]]]:
    units = data.get("units", {})
    unit_name = "USD" if "USD" in units else (next(iter(units), None))
    series = units.get(unit_name, []) if unit_name else []
    ordered = sorted(series, key=lambda x: x.get("end") or "", reverse=True)
    return unit_name, ordered


def _prefer_annual(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Prefer FY / frame CY#### (ano cheio), sem duplicar o mesmo end."""
    annual = [
        r
        for r in rows
        if r.get("fp") == "FY"
        or (isinstance(r.get("frame"), str) and re.fullmatch(r"CY\d{4}", r["frame"]))
    ]
    pool = annual or rows
    seen: set[str] = set()
    out: list[dict[str, Any]] = []
    for r in pool:
        key = str(r.get("end") or "")
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def get_concept(
    ticker_or_cik: str,
    concept: str,
    taxonomy: str = "auto",
    limit: int = 20,
    *,
    annual_only: bool = False,
) -> dict[str, Any]:
    """Single XBRL concept time series.

    taxonomy:
      - ``us-gaap`` / ``ifrs-full``: fixo
      - ``auto`` (padrao): tenta us-gaap; se vazio ou serie antiga (< ~3 anos),
        tenta aliases IFRS (ex.: NetIncomeLoss -> ProfitLossAttributableToOwnersOfParent).
        Necessario para emissores estrangeiros (PBR, VALE, etc.).
    """
    cik = resolve_cik(ticker_or_cik)
    concept = concept.strip()
    taxonomy = (taxonomy or "auto").strip().lower()
    if taxonomy in {"auto", "automatico", ""}:
        taxonomy = "auto"
    limit = max(1, min(int(limit), 100))

    attempts: list[tuple[str, str]] = []
    if taxonomy == "auto":
        attempts.append(("us-gaap", concept))
        for alt in _IFRS_ALIASES.get(concept, [concept]):
            attempts.append(("ifrs-full", alt))
        if concept not in _IFRS_ALIASES:
            attempts.append(("ifrs-full", concept))
    else:
        attempts.append((taxonomy, concept))

    last_err: Exception | None = None
    best: dict[str, Any] | None = None
    best_end = ""

    for tax, tag in attempts:
        try:
            data = _fetch_concept_raw(cik, tax, tag)
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            continue
        unit_name, ordered = _series_from_concept(data)
        if annual_only:
            ordered = _prefer_annual(ordered)
        else:
            seen: set[str] = set()
            dedup: list[dict[str, Any]] = []
            for r in ordered:
                key = f"{r.get('end')}|{r.get('val')}|{r.get('fp')}"
                if key in seen:
                    continue
                seen.add(key)
                dedup.append(r)
            ordered = dedup

        if not ordered:
            continue
        top_end = str(ordered[0].get("end") or "")
        candidate = {
            "cik": cik,
            "entityName": data.get("entityName"),
            "taxonomy": tax,
            "concept": tag,
            "concept_requested": concept,
            "label": data.get("label"),
            "unit": unit_name,
            "count": len(ordered),
            "recent": ordered[:limit],
            "provider": "data.sec.gov/api/xbrl/companyconcept",
        }
        if top_end > best_end:
            best = candidate
            best_end = top_end
        # Se ja temos dado recente (>= 2020), para
        if top_end >= "2020-01-01":
            best = candidate
            break

    if best is None:
        if last_err is not None:
            raise last_err
        raise ValueError(
            f"Conceito nao encontrado para {ticker_or_cik!r}: {concept!r} "
            f"(tentativas: {attempts})"
        )

    # Aviso se a serie ainda for antiga
    note = None
    if best_end and best_end < "2020-01-01":
        note = (
            "Serie antiga no taxonomy usado. Emissores estrangeiros (20-F/6-K) "
            "costumam estar em ifrs-full (ex.: ProfitLossAttributableToOwnersOfParent). "
            "Rode com --taxonomy ifrs-full."
        )
        best["note"] = note
    if taxonomy == "auto" and best["concept"] != concept:
        best["note"] = (
            (best.get("note") + " " if best.get("note") else "")
            + f"Auto: usou {best['taxonomy']}:{best['concept']} "
            f"(pedido era {concept})."
        ).strip()
    return best


def _year_from_row(row: dict[str, Any]) -> int | None:
    """Ano economico do ponto XBRL.

    Preferir ``frame`` (CY2010) ou ano de ``end``. Nao usar ``fy`` do formulario:
    em 20-F comparativos o fy e o ano do filing, nao o do periodo (ex.: fy=2009
    com end=2007-12-31).
    """
    frame = row.get("frame")
    if isinstance(frame, str):
        m = re.fullmatch(r"CY(\d{4})", frame)
        if m:
            return int(m.group(1))
        m = re.fullmatch(r"CY(\d{4})Q\d", frame)
        if m:
            return int(m.group(1))
    end = str(row.get("end") or "")
    m = re.match(r"(\d{4})", end)
    if m:
        return int(m.group(1))
    if row.get("fy") is not None:
        try:
            return int(row["fy"])
        except (TypeError, ValueError):
            return None
    return None


def get_concept_range(
    ticker_or_cik: str,
    concept: str,
    *,
    year_from: int = 2008,
    year_to: int = 2025,
    annual_only: bool = True,
) -> dict[str, Any]:
    """Serie anual unindo us-gaap (antigo) + ifrs-full (recente).

    Para PBR/VALE: US-GAAP cobre ~2008-2011; IFRS cobre anos seguintes.
    Em overlap, prefere IFRS. Valores podem nao ser estritamente comparaveis
    entre normas.
    """
    year_from = int(year_from)
    year_to = int(year_to)
    if year_from > year_to:
        year_from, year_to = year_to, year_from

    # Busca as duas fontes com limite alto
    us = None
    ifrs = None
    try:
        us = get_concept(
            ticker_or_cik,
            concept,
            taxonomy="us-gaap",
            limit=100,
            annual_only=annual_only,
        )
    except Exception:  # noqa: BLE001
        us = None
    try:
        ifrs = get_concept(
            ticker_or_cik,
            concept,
            taxonomy="auto" if concept in _IFRS_ALIASES else "ifrs-full",
            limit=100,
            annual_only=annual_only,
        )
        # Se auto caiu de volta no us-gaap, forcar ifrs aliases
        if ifrs and ifrs.get("taxonomy") == "us-gaap":
            ifrs = None
            for alt in _IFRS_ALIASES.get(concept, [concept]):
                try:
                    ifrs = get_concept(
                        ticker_or_cik,
                        alt,
                        taxonomy="ifrs-full",
                        limit=100,
                        annual_only=annual_only,
                    )
                    break
                except Exception:  # noqa: BLE001
                    continue
    except Exception:  # noqa: BLE001
        ifrs = None

    by_year: dict[int, dict[str, Any]] = {}

    def _ingest(block: dict[str, Any] | None, source: str) -> None:
        if not block:
            return
        for r in block.get("recent", []):
            y = _year_from_row(r)
            if y is None or y < year_from or y > year_to:
                continue
            entry = {
                "year": y,
                "end": r.get("end"),
                "val": r.get("val"),
                "unit": block.get("unit"),
                "fp": r.get("fp"),
                "form": r.get("form"),
                "filed": r.get("filed"),
                "frame": r.get("frame"),
                "taxonomy": block.get("taxonomy"),
                "concept": block.get("concept"),
                "source": source,
            }
            # Prefer IFRS on overlap
            if y not in by_year or source == "ifrs":
                by_year[y] = entry

    _ingest(us, "us-gaap")
    _ingest(ifrs, "ifrs")

    series = [by_year[y] for y in sorted(by_year.keys())]
    entity = (ifrs or us or {}).get("entityName")
    cik = resolve_cik(ticker_or_cik)

    missing = [y for y in range(year_from, year_to + 1) if y not in by_year]

    return {
        "cik": cik,
        "entityName": entity,
        "concept_requested": concept,
        "year_from": year_from,
        "year_to": year_to,
        "annual_only": annual_only,
        "count": len(series),
        "years_missing": missing,
        "series": series,
        "sources": {
            "us-gaap": {
                "concept": (us or {}).get("concept"),
                "label": (us or {}).get("label"),
                "points": len((us or {}).get("recent") or []),
            },
            "ifrs-full": {
                "concept": (ifrs or {}).get("concept"),
                "label": (ifrs or {}).get("label"),
                "points": len((ifrs or {}).get("recent") or []),
            },
        },
        "note": (
            "Serie mesclada us-gaap + ifrs-full. Em anos com as duas fontes, "
            "prevalece IFRS. Mudanca de norma pode afetar comparabilidade."
        ),
        "provider": "data.sec.gov/api/xbrl/companyconcept (merged)",
    }
