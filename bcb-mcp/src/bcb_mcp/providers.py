"""Banco Central do Brasil open-data client (SGS + OLINDA). No API key."""

from __future__ import annotations

import os
import re
import time
from datetime import date, datetime, timedelta
from typing import Any

import httpx

SGS_BASE = "https://api.bcb.gov.br/dados/serie"
OLINDA_PTAX = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata"
OLINDA_EXPECT = (
    "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata"
)

DEFAULT_UA = os.getenv(
    "BCB_USER_AGENT",
    "SEC-data-analysys-bcb-mcp/0.1 (cafla19791@gmail.com)",
)

# Soft throttle (~5 req/s).
_MIN_INTERVAL = 0.2
_last_request = 0.0

# SGS queries capped ~10 years (BCB rule since 2025).
_MAX_CHUNK_DAYS = 3650

# Common series aliases -> SGS code.
KNOWN_SERIES: dict[str, dict[str, Any]] = {
    "selic": {
        "code": 11,
        "name": "Taxa Selic (% a.d.)",
        "periodicity": "daily",
    },
    "selic_meta": {
        "code": 432,
        "name": "Meta Selic definida pelo Copom (% a.a.)",
        "periodicity": "daily",
    },
    "selic_mes": {
        "code": 4390,
        "name": "Selic acumulada no mes (%)",
        "periodicity": "monthly",
    },
    "cdi": {
        "code": 12,
        "name": "Taxa CDI (% a.d.)",
        "periodicity": "daily",
    },
    "cdi_mes": {
        "code": 4391,
        "name": "CDI acumulado no mes (%)",
        "periodicity": "monthly",
    },
    "ipca": {
        "code": 433,
        "name": "IPCA variacao mensal (%)",
        "periodicity": "monthly",
    },
    "ipca_12m": {
        "code": 13522,
        "name": "IPCA acumulado 12 meses (%)",
        "periodicity": "monthly",
    },
    "inpc": {
        "code": 188,
        "name": "INPC variacao mensal (%)",
        "periodicity": "monthly",
    },
    "igpm": {
        "code": 189,
        "name": "IGP-M variacao mensal (%)",
        "periodicity": "monthly",
    },
    "dolar": {
        "code": 1,
        "name": "Taxa de cambio - Dolar americano (venda)",
        "periodicity": "daily",
    },
    "dolar_compra": {
        "code": 10813,
        "name": "Taxa de cambio - Dolar americano (compra)",
        "periodicity": "daily",
    },
    "euro": {
        "code": 21619,
        "name": "Taxa de cambio - Euro (venda)",
        "periodicity": "daily",
    },
    "pib_mensal": {
        "code": 4380,
        "name": "PIB mensal - valores correntes (R$ milhoes)",
        "periodicity": "monthly",
    },
    "desemprego": {
        "code": 24369,
        "name": "Taxa de desocupacao PNAD Continua (%)",
        "periodicity": "monthly",
    },
    "divida_liquida_pib": {
        "code": 4536,
        "name": "Divida liquida do setor publico (% PIB)",
        "periodicity": "monthly",
    },
    "reservas": {
        "code": 3546,
        "name": "Reservas internacionais (US$ milhoes)",
        "periodicity": "daily",
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


def _get_json(url: str, *, timeout: float = 60.0) -> Any:
    _throttle()
    with httpx.Client(headers=_headers(), timeout=timeout, follow_redirects=True) as client:
        resp = client.get(url)
        if resp.status_code == 404:
            raise ValueError(f"Recurso BCB nao encontrado (404): {url}")
        if resp.status_code == 429:
            raise RuntimeError("BCB rate limit (429). Aguarde e tente novamente.")
        resp.raise_for_status()
        return resp.json()


def _parse_date(value: str | date | None, *, default: date | None = None) -> date | None:
    if value is None or value == "":
        return default
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Data invalida: {value!r} (use YYYY-MM-DD ou DD/MM/YYYY)")


def _fmt_br(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def _fmt_iso(d: date) -> str:
    return d.isoformat()


def resolve_series(code_or_alias: str | int) -> dict[str, Any]:
    """Resolve alias (selic, ipca...) or numeric SGS code."""
    if isinstance(code_or_alias, int) or re.fullmatch(r"\d+", str(code_or_alias).strip()):
        code = int(code_or_alias)
        alias = next(
            (k for k, v in KNOWN_SERIES.items() if int(v["code"]) == code),
            None,
        )
        meta = KNOWN_SERIES.get(alias or "", {})
        return {
            "code": code,
            "alias": alias,
            "name": meta.get("name") or f"SGS {code}",
            "periodicity": meta.get("periodicity"),
        }

    key = str(code_or_alias).strip().lower().replace("-", "_").replace(" ", "_")
    # Common synonyms
    synonyms = {
        "selic_diaria": "selic",
        "meta_selic": "selic_meta",
        "ptax": "dolar",
        "usd": "dolar",
        "cambio": "dolar",
        "fx": "dolar",
    }
    key = synonyms.get(key, key)
    if key not in KNOWN_SERIES:
        known = ", ".join(sorted(KNOWN_SERIES))
        raise ValueError(
            f"Serie desconhecida: {code_or_alias!r}. "
            f"Use um codigo SGS numerico ou alias: {known}"
        )
    meta = KNOWN_SERIES[key]
    return {
        "code": int(meta["code"]),
        "alias": key,
        "name": meta["name"],
        "periodicity": meta.get("periodicity"),
    }


def list_known_series() -> dict[str, Any]:
    rows = [
        {
            "alias": alias,
            "code": meta["code"],
            "name": meta["name"],
            "periodicity": meta.get("periodicity"),
        }
        for alias, meta in sorted(KNOWN_SERIES.items(), key=lambda kv: kv[1]["code"])
    ]
    return {
        "count": len(rows),
        "series": rows,
        "note": (
            "Catalogo local de aliases comuns. Qualquer codigo SGS numerico "
            "tambem funciona em get_sgs_series / serie."
        ),
        "provider": "bcb-mcp/KNOWN_SERIES",
        "portal": "https://www3.bcb.gov.br/sgspub/",
    }


def _date_chunks(start: date, end: date) -> list[tuple[date, date]]:
    if start > end:
        start, end = end, start
    chunks: list[tuple[date, date]] = []
    cur = start
    while cur <= end:
        nxt = min(cur + timedelta(days=_MAX_CHUNK_DAYS - 1), end)
        chunks.append((cur, nxt))
        cur = nxt + timedelta(days=1)
    return chunks


def _normalize_sgs_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    seen: set[str] = set()
    for r in rows:
        raw_date = str(r.get("data") or "")
        d = _parse_date(raw_date)
        if d is None:
            continue
        key = _fmt_iso(d)
        if key in seen:
            continue
        seen.add(key)
        val = r.get("valor")
        try:
            num = float(str(val).replace(",", ".")) if val is not None and val != "" else None
        except ValueError:
            num = None
        out.append({"date": key, "value": num, "raw": val})
    out.sort(key=lambda x: x["date"])
    return out


def get_sgs_series(
    code_or_alias: str | int,
    *,
    date_from: str | date | None = None,
    date_to: str | date | None = None,
    last: int | None = None,
) -> dict[str, Any]:
    """Fetch one SGS time series.

    Prefer ``last`` for recent points, or ``date_from``/``date_to`` for a window.
    Windows longer than ~10 years are fetched in chunks.
    """
    meta = resolve_series(code_or_alias)
    code = meta["code"]

    if last is not None:
        n = max(1, min(int(last), 10000))
        url = f"{SGS_BASE}/bcdata.sgs.{code}/dados/ultimos/{n}?formato=json"
        rows = _get_json(url)
        if not isinstance(rows, list):
            raise RuntimeError(f"Resposta SGS inesperada para codigo {code}")
        series = _normalize_sgs_rows(rows)
        return {
            **meta,
            "mode": "last",
            "last": n,
            "count": len(series),
            "series": series,
            "provider": "api.bcb.gov.br/dados/serie (SGS)",
            "url": url,
        }

    end = _parse_date(date_to, default=date.today())
    assert end is not None
    start = _parse_date(date_from, default=end - timedelta(days=365))
    assert start is not None

    all_rows: list[dict[str, Any]] = []
    urls: list[str] = []
    for a, b in _date_chunks(start, end):
        url = (
            f"{SGS_BASE}/bcdata.sgs.{code}/dados?formato=json"
            f"&dataInicial={_fmt_br(a)}&dataFinal={_fmt_br(b)}"
        )
        urls.append(url)
        chunk = _get_json(url)
        if isinstance(chunk, list):
            all_rows.extend(chunk)

    series = _normalize_sgs_rows(all_rows)
    return {
        **meta,
        "mode": "range",
        "date_from": _fmt_iso(start),
        "date_to": _fmt_iso(end),
        "chunks": len(urls),
        "count": len(series),
        "series": series,
        "provider": "api.bcb.gov.br/dados/serie (SGS)",
        "urls": urls if len(urls) <= 5 else urls[:5] + [f"... (+{len(urls) - 5})"],
    }


def get_ptax(
    *,
    date_from: str | date | None = None,
    date_to: str | date | None = None,
    last_days: int = 30,
) -> dict[str, Any]:
    """USD/BRL PTAX via OLINDA (compra/venda + boletim)."""
    end = _parse_date(date_to, default=date.today())
    assert end is not None
    if date_from:
        start = _parse_date(date_from)
    else:
        start = end - timedelta(days=max(1, int(last_days)))
    assert start is not None
    if start > end:
        start, end = end, start

    # OLINDA PTAX expects MM-DD-YYYY
    d1 = start.strftime("%m-%d-%Y")
    d2 = end.strftime("%m-%d-%Y")
    url = (
        f"{OLINDA_PTAX}/CotacaoDolarPeriodo(dataInicial=@d1,dataFinalCotacao=@d2)"
        f"?@d1='{d1}'&@d2='{d2}'&$format=json&$orderby=dataHoraCotacao"
    )
    payload = _get_json(url)
    values = payload.get("value") if isinstance(payload, dict) else None
    if values is None:
        raise RuntimeError("Resposta OLINDA PTAX inesperada")

    rows = []
    for r in values:
        rows.append(
            {
                "datetime": r.get("dataHoraCotacao"),
                "buy": r.get("cotacaoCompra"),
                "sell": r.get("cotacaoVenda"),
                "type": r.get("tipoBoletim"),
            }
        )

    return {
        "pair": "USD/BRL",
        "date_from": _fmt_iso(start),
        "date_to": _fmt_iso(end),
        "count": len(rows),
        "series": rows,
        "provider": "olinda.bcb.gov.br/PTAX",
        "url": url,
        "note": "Boletins Abertura/Intermediario/Fechamento podem coexistir no mesmo dia.",
    }


def get_expectativas(
    indicator: str = "IPCA",
    *,
    top: int = 20,
    filter_expr: str | None = None,
) -> dict[str, Any]:
    """Focus / market expectations (OLINDA Expectativas).

    indicator examples: IPCA, Selic, Cambio, PIB Total, IGP-M.
    """
    top = max(1, min(int(top), 200))
    ind = indicator.strip()
    # Escape single quotes for OData
    ind_esc = ind.replace("'", "''")
    filt = filter_expr or f"Indicador eq '{ind_esc}'"
    url = (
        f"{OLINDA_EXPECT}/ExpectativasMercadoAnuais"
        f"?$filter={filt}&$orderby=Data desc&$top={top}&$format=json"
    )
    # URL-encode filter lightly via httpx params would be better; keep simple path
    # Rebuild with encoded filter
    from urllib.parse import quote

    url = (
        f"{OLINDA_EXPECT}/ExpectativasMercadoAnuais"
        f"?$filter={quote(filt)}&$orderby=Data%20desc&$top={top}&$format=json"
    )
    payload = _get_json(url)
    values = payload.get("value") if isinstance(payload, dict) else None
    if values is None:
        raise RuntimeError("Resposta OLINDA Expectativas inesperada")

    return {
        "indicator": ind,
        "endpoint": "ExpectativasMercadoAnuais",
        "filter": filt,
        "count": len(values),
        "rows": values,
        "provider": "olinda.bcb.gov.br/Expectativas",
        "url": url,
        "note": (
            "Mediana/media das expectativas de mercado (Focus). "
            "Para mensais use o endpoint ExpectativasMercadoMensais via portal OLINDA."
        ),
    }
