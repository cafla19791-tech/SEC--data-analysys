"""Minimal BCB SGS client (DBGG, BNDES) used by the annual fiscal collector."""

from __future__ import annotations

import os
import time
from datetime import date, datetime
from typing import Any

import httpx

SGS_BASE = "https://api.bcb.gov.br/dados/serie"
DEFAULT_UA = os.getenv(
    "BCB_USER_AGENT",
    os.getenv(
        "TESOURO_USER_AGENT",
        "SEC-data-analysys-tesouro-mcp/0.1 (cafla19791@gmail.com)",
    ),
)
_MIN_INTERVAL = 0.2
_last_request = 0.0

# Well-known series used by the collector.
SERIES = {
    "dbgg_rs_mi": {
        "code": 13761,
        "name": "Divida bruta do governo geral - R$ milhoes",
        "unit": "R$ milhoes",
    },
    "dbgg_pib": {
        "code": 13762,
        "name": "Divida bruta do governo geral - % PIB",
        "unit": "% PIB",
    },
    "bndes_desembolso": {
        "code": 7415,
        "name": "Desembolsos do sistema BNDES - Total",
        "unit": "R$ milhoes",
    },
}


def _headers() -> dict[str, str]:
    return {
        "User-Agent": DEFAULT_UA,
        "Accept": "application/json",
        "Accept-Encoding": "gzip, deflate",
    }


def _throttle() -> None:
    global _last_request
    now = time.monotonic()
    wait = _MIN_INTERVAL - (now - _last_request)
    if wait > 0:
        time.sleep(wait)
    _last_request = time.monotonic()


def _get_json(url: str, *, timeout: float = 90.0) -> Any:
    _throttle()
    with httpx.Client(headers=_headers(), timeout=timeout, follow_redirects=True) as client:
        resp = client.get(url)
        if resp.status_code == 404:
            raise ValueError(f"Serie BCB nao encontrada (404): {url}")
        if resp.status_code == 429:
            raise RuntimeError("BCB rate limit (429). Aguarde e tente novamente.")
        resp.raise_for_status()
        return resp.json()


def fetch_sgs_range(
    code: int,
    start: date,
    end: date,
    *,
    chunk_years: int = 3,
) -> dict[date, float]:
    """Fetch monthly/daily SGS points between dates (inclusive)."""
    if start > end:
        start, end = end, start
    out: dict[date, float] = {}
    cur = start
    while cur <= end:
        nxt = min(date(cur.year + chunk_years, 12, 31), end)
        url = (
            f"{SGS_BASE}/bcdata.sgs.{int(code)}/dados?formato=json"
            f"&dataInicial={cur.strftime('%d/%m/%Y')}"
            f"&dataFinal={nxt.strftime('%d/%m/%Y')}"
        )
        chunk = _get_json(url)
        if isinstance(chunk, list):
            for row in chunk:
                d = datetime.strptime(str(row["data"]), "%d/%m/%Y").date()
                out[d] = float(str(row["valor"]).replace(",", "."))
        if nxt >= end:
            break
        cur = date(nxt.year + 1, 1, 1)
    return out


def annual_sum(points: dict[date, float]) -> dict[int, float]:
    out: dict[int, float] = {}
    for d, v in points.items():
        out[d.year] = out.get(d.year, 0.0) + float(v)
    return out


def december_stocks(points: dict[date, float]) -> dict[int, float]:
    """Year -> value on Dec observation (date month==12)."""
    out: dict[int, float] = {}
    for d, v in points.items():
        if d.month == 12:
            out[d.year] = float(v)
    return out
