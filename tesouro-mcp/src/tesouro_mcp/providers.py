"""Tesouro Nacional fiscal statistics client (ARIA RTN + Grandes Numeros + CKAN)."""

from __future__ import annotations

import os
import re
import time
from datetime import date, datetime
from typing import Any
from urllib.parse import urlencode

import httpx

ARIA_BASE = "https://apiapex.tesouro.gov.br/aria/v1/series-temporais/custom"
GRANDES_NUMEROS = "https://grandesnumeros.tesouro.gov.br"
CKAN_API = "https://www.tesourotransparente.gov.br/ckan/api/3/action"

DEFAULT_UA = os.getenv(
    "TESOURO_USER_AGENT",
    "SEC-data-analysys-tesouro-mcp/0.1 (cafla19791@gmail.com)",
)

_MIN_INTERVAL = 0.2
_last_request = 0.0
_MAX_PAGES = 50

# RTN themes available on ARIA resultado-fiscal.
TEMAS: dict[str, dict[str, str]] = {
    "10": {
        "alias": "resultado_fiscal",
        "name": "Resultado Fiscal do Governo Central - Valores Mensais (Tabela 1.2)",
    },
    "13": {
        "alias": "investimento",
        "name": "Investimento do Governo Federal (Tabela 1.3)",
    },
    "20": {
        "alias": "custeio",
        "name": "Custeio Administrativo do Governo Central (Tabela 1.4)",
    },
}

# Handy aliases -> (tema, codigo_serie)
KNOWN_SERIES: dict[str, dict[str, str]] = {
    "resultado_primario": {
        "tema": "10",
        "codigo": "10.04.1",
        "name": "Resultado Primario - Governo Central",
    },
    "resultado_primario_tesouro": {
        "tema": "10",
        "codigo": "10.04.1.1",
        "name": "Resultado Primario - Tesouro Nacional",
    },
    "resultado_primario_previdencia": {
        "tema": "10",
        "codigo": "10.04.1.2",
        "name": "Resultado Primario - Previdencia Social",
    },
    "resultado_primario_bc": {
        "tema": "10",
        "codigo": "10.04.1.3",
        "name": "Resultado Primario - Banco Central",
    },
    "resultado_abaixo_linha": {
        "tema": "10",
        "codigo": "10.07.1",
        "name": "Resultado Primario do Governo Central - Abaixo da Linha",
    },
    "receita_total": {
        "tema": "10",
        "codigo": "10.01.1",
        "name": "Receita Total",
    },
    "receita_liquida": {
        "tema": "10",
        "codigo": "10.01.2",
        "name": "Receita Liquida (Receita Total - Transf. por Reparticao)",
    },
    "despesa_total": {
        "tema": "10",
        "codigo": "10.03.1",
        "name": "Despesa total",
    },
    "juros_nominais": {
        "tema": "10",
        "codigo": "10.08.1",
        "name": "Juros Nominais",
    },
    "investimento_total": {
        "tema": "13",
        "codigo": "13.1.1",
        "name": "Investimento Total",
    },
    "custeio_total": {
        "tema": "20",
        "codigo": "20.1.1",
        "name": "Custeio Administrativo - Total",
    },
}

GRANDES_ENDPOINTS: dict[str, str] = {
    "resultado_primario": "/resultado_primario",
    "estoque_dpf": "/estoque_dpf",
    "receita_primaria_liquida": "/receita_primaria_liquida",
    "despesa_primaria_total": "/despesa_primaria_total",
    "emissoes_dpf": "/emissoes_dpf",
    "resgate_dpf": "/resgate_dpf",
    "teto_gasto_atingido": "/teto_gasto_atingido",
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
    # Fix occasional double-slash from ARIA pagination links
    url = url.replace("aria//v1", "aria/v1")
    with httpx.Client(headers=_headers(), timeout=timeout, follow_redirects=True) as client:
        resp = client.get(url)
        if resp.status_code == 404:
            raise ValueError(f"Recurso Tesouro nao encontrado (404): {url}")
        if resp.status_code == 429:
            raise RuntimeError("Tesouro rate limit (429). Aguarde e tente novamente.")
        resp.raise_for_status()
        return resp.json()


def _paginate(url: str, *, max_pages: int = _MAX_PAGES) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    cur = url
    pages = 0
    while cur and pages < max_pages:
        data = _get_json(cur)
        if not isinstance(data, dict):
            raise RuntimeError(f"Resposta ARIA inesperada: {type(data)}")
        chunk = data.get("registros") or []
        if not isinstance(chunk, list):
            raise RuntimeError("Campo 'registros' ausente/invalido na resposta ARIA")
        rows.extend(chunk)
        pages += 1
        nxt = data.get("next")
        if not nxt or not chunk:
            break
        cur = str(nxt).replace("aria//v1", "aria/v1")
    return rows


def _parse_month(value: str | None) -> str | None:
    """Normalize to MM/AAAA required by ARIA."""
    if value is None or value == "":
        return None
    s = str(value).strip()
    m = re.fullmatch(r"(\d{2})/(\d{4})", s)
    if m:
        return f"{m.group(1)}/{m.group(2)}"
    m = re.fullmatch(r"(\d{4})-(\d{2})", s)
    if m:
        return f"{m.group(2)}/{m.group(1)}"
    m = re.fullmatch(r"(\d{4})-(\d{2})-\d{2}", s)
    if m:
        return f"{m.group(2)}/{m.group(1)}"
    m = re.fullmatch(r"(\d{4})/(\d{2})", s)
    if m:
        return f"{m.group(2)}/{m.group(1)}"
    raise ValueError(
        f"Mes invalido: {value!r}. Use MM/AAAA, YYYY-MM ou YYYY-MM-DD."
    )


def _normalize_row(row: dict[str, Any]) -> dict[str, Any]:
    raw_date = row.get("data")
    iso = None
    if isinstance(raw_date, str) and raw_date:
        try:
            iso = datetime.fromisoformat(raw_date.replace("Z", "+00:00")).date().isoformat()
        except ValueError:
            iso = raw_date[:10]
    val = row.get("valor")
    try:
        num = float(val) if val is not None else None
    except (TypeError, ValueError):
        num = None
    return {
        "date": iso,
        "value": num,
        "unit": "R$ milhoes",
        "tema": row.get("codigoTema") or row.get("nomeTema"),
        "subtema": row.get("nomeSubtema"),
        "codigo_serie": row.get("codigoSerie"),
        "nome_serie": row.get("nomeSerie"),
    }


def list_temas() -> dict[str, Any]:
    return {
        "count": len(TEMAS),
        "temas": [
            {"codigo": k, **v}
            for k, v in TEMAS.items()
        ],
        "note": (
            "Temas 10/13/20 correspondem as tabelas 1.2/1.3/1.4 do Boletim "
            "Resultado do Tesouro Nacional (RTN)."
        ),
        "provider": "apiapex.tesouro.gov.br/aria (series-temporais)",
    }


def list_known_aliases() -> dict[str, Any]:
    rows = [
        {
            "alias": alias,
            "tema": meta["tema"],
            "codigo_serie": meta["codigo"],
            "name": meta["name"],
        }
        for alias, meta in sorted(KNOWN_SERIES.items())
    ]
    return {
        "count": len(rows),
        "aliases": rows,
        "provider": "tesouro-mcp/KNOWN_SERIES",
    }


def resolve_serie(alias_or_code: str) -> dict[str, str]:
    key = str(alias_or_code).strip().lower().replace("-", "_").replace(" ", "_")
    synonyms = {
        "rp": "resultado_primario",
        "primario": "resultado_primario",
        "receita": "receita_total",
        "despesa": "despesa_total",
        "investimento": "investimento_total",
        "custeio": "custeio_total",
    }
    key = synonyms.get(key, key)
    if key in KNOWN_SERIES:
        meta = KNOWN_SERIES[key]
        return {
            "alias": key,
            "tema": meta["tema"],
            "codigo_serie": meta["codigo"],
            "name": meta["name"],
        }
    # bare series code like 10.04.1
    if re.fullmatch(r"\d+(\.\d+)+", str(alias_or_code).strip()):
        code = str(alias_or_code).strip()
        tema = code.split(".", 1)[0]
        return {
            "alias": None,
            "tema": tema,
            "codigo_serie": code,
            "name": f"Serie {code}",
        }
    known = ", ".join(sorted(KNOWN_SERIES))
    raise ValueError(
        f"Serie desconhecida: {alias_or_code!r}. "
        f"Use alias ({known}) ou codigo (ex.: 10.04.1)."
    )


def list_series(*, tema: str | None = None) -> dict[str, Any]:
    """Catalog of RTN series from ARIA /custom/series."""
    rows = _paginate(f"{ARIA_BASE}/series?pageSize=1000")
    if tema:
        tema = str(tema).strip()
        # allow alias
        for code, meta in TEMAS.items():
            if tema.lower() in {code, meta["alias"]}:
                tema = code
                break
        rows = [r for r in rows if str(r.get("codigoTema")) == str(tema)]

    out = [
        {
            "tema": r.get("codigoTema"),
            "nome_tema": r.get("nomeTema"),
            "subtema": r.get("nomeSubtema"),
            "codigo_subtema": r.get("codigoSubtema"),
            "codigo_serie": r.get("codigoSerie"),
            "nome_serie": r.get("nomeSerie"),
        }
        for r in rows
    ]
    return {
        "count": len(out),
        "tema_filter": tema,
        "series": out,
        "provider": "apiapex.tesouro.gov.br/aria/custom/series",
    }


def search_series(query: str, *, tema: str | None = None, limit: int = 30) -> dict[str, Any]:
    q = query.strip().lower()
    if not q:
        raise ValueError("query vazia")
    catalog = list_series(tema=tema)
    hits = [
        s
        for s in catalog["series"]
        if q in (s.get("nome_serie") or "").lower()
        or q in (s.get("subtema") or "").lower()
        or q in (s.get("codigo_serie") or "").lower()
    ]
    limit = max(1, min(int(limit), 200))
    return {
        "query": query,
        "count": len(hits[:limit]),
        "total_matches": len(hits),
        "series": hits[:limit],
        "provider": catalog["provider"],
    }


def get_resultado_fiscal(
    *,
    tema: str = "10",
    data_inicio: str | None = None,
    data_fim: str | None = None,
    codigo_serie: str | None = None,
    correcao_ipca: bool = False,
) -> dict[str, Any]:
    """Monthly RTN fiscal series (ARIA resultado-fiscal).

    Values are in R$ milhoes (current prices, unless correcao_ipca=True).
    """
    tema_key = str(tema).strip().lower()
    for code, meta in TEMAS.items():
        if tema_key in {code, meta["alias"]}:
            tema_key = code
            break
    if tema_key not in TEMAS:
        raise ValueError(f"Tema invalido: {tema!r}. Use 10, 13 ou 20.")

    params: dict[str, str] = {"tema": tema_key}
    di = _parse_month(data_inicio)
    df = _parse_month(data_fim)
    if di:
        params["data_inicio"] = di
    if df:
        params["data_fim"] = df
    if codigo_serie:
        params["codigo_da_serie"] = str(codigo_serie).strip()
    if correcao_ipca:
        params["correcao_ipca"] = "true"

    url = f"{ARIA_BASE}/resultado-fiscal?{urlencode(params)}"
    raw = _paginate(url)
    series = [_normalize_row(r) for r in raw]
    series.sort(key=lambda x: (x.get("codigo_serie") or "", x.get("date") or ""))

    return {
        "tema": tema_key,
        "tema_name": TEMAS[tema_key]["name"],
        "data_inicio": di,
        "data_fim": df,
        "codigo_serie": codigo_serie,
        "correcao_ipca": bool(correcao_ipca),
        "count": len(series),
        "unit": "R$ milhoes",
        "series": series,
        "provider": "apiapex.tesouro.gov.br/aria/custom/resultado-fiscal",
        "url": url,
        "note": (
            "Valores mensais do Boletim Resultado do Tesouro Nacional (RTN). "
            "Unidade tipica: R$ milhoes."
        ),
    }


def get_serie(
    alias_or_code: str,
    *,
    data_inicio: str | None = None,
    data_fim: str | None = None,
    correcao_ipca: bool = False,
) -> dict[str, Any]:
    """Convenience wrapper: alias/code -> resultado-fiscal filtered series."""
    meta = resolve_serie(alias_or_code)
    out = get_resultado_fiscal(
        tema=meta["tema"],
        data_inicio=data_inicio,
        data_fim=data_fim,
        codigo_serie=meta["codigo_serie"],
        correcao_ipca=correcao_ipca,
    )
    out["alias"] = meta.get("alias")
    out["name"] = meta.get("name")
    return out


def get_grandes_numeros(metric: str | None = None) -> dict[str, Any]:
    """Headline fiscal figures from Tesouro Grandes Numeros API."""
    if metric:
        key = metric.strip().lower().replace("-", "_")
        if key not in GRANDES_ENDPOINTS:
            known = ", ".join(sorted(GRANDES_ENDPOINTS))
            raise ValueError(f"Metrica desconhecida: {metric!r}. Use: {known}")
        path = GRANDES_ENDPOINTS[key]
        data = _get_json(f"{GRANDES_NUMEROS}{path}")
        return {
            "metric": key,
            "value": data.get("num") if isinstance(data, dict) else data,
            "raw": data,
            "provider": "grandesnumeros.tesouro.gov.br",
            "url": f"{GRANDES_NUMEROS}{path}",
        }

    items = []
    for key, path in GRANDES_ENDPOINTS.items():
        try:
            data = _get_json(f"{GRANDES_NUMEROS}{path}")
            items.append(
                {
                    "metric": key,
                    "value": data.get("num") if isinstance(data, dict) else data,
                }
            )
        except Exception as exc:  # noqa: BLE001
            items.append({"metric": key, "error": str(exc)})
    return {
        "count": len(items),
        "items": items,
        "provider": "grandesnumeros.tesouro.gov.br",
        "note": "Numeros-resumo da capa do Tesouro Transparente (texto formatado).",
    }


def ckan_package_search(query: str = "resultado do tesouro", *, rows: int = 10) -> dict[str, Any]:
    """Search open datasets on Tesouro Transparente (CKAN)."""
    rows = max(1, min(int(rows), 50))
    url = f"{CKAN_API}/package_search?{urlencode({'q': query, 'rows': rows})}"
    data = _get_json(url)
    if not isinstance(data, dict) or not data.get("success"):
        raise RuntimeError("Falha na busca CKAN do Tesouro Transparente")
    result = data.get("result") or {}
    packages = []
    for p in result.get("results") or []:
        resources = [
            {
                "name": r.get("name"),
                "format": r.get("format"),
                "url": r.get("url"),
                "last_modified": r.get("last_modified"),
            }
            for r in (p.get("resources") or [])
        ]
        packages.append(
            {
                "id": p.get("name") or p.get("id"),
                "title": p.get("title"),
                "notes": (p.get("notes") or "")[:400],
                "organization": (p.get("organization") or {}).get("title"),
                "resources": resources,
            }
        )
    return {
        "query": query,
        "count": len(packages),
        "total": result.get("count"),
        "packages": packages,
        "provider": "tesourotransparente.gov.br/ckan",
        "url": url,
    }


def ckan_package_show(package_id: str = "resultado-do-tesouro-nacional") -> dict[str, Any]:
    """Show one CKAN package (resources / download URLs)."""
    url = f"{CKAN_API}/package_show?{urlencode({'id': package_id})}"
    data = _get_json(url)
    if not isinstance(data, dict) or not data.get("success"):
        raise RuntimeError(f"Pacote CKAN nao encontrado: {package_id!r}")
    p = data.get("result") or {}
    resources = [
        {
            "id": r.get("id"),
            "name": r.get("name"),
            "format": r.get("format"),
            "url": r.get("url"),
            "description": r.get("description"),
            "last_modified": r.get("last_modified"),
            "size": r.get("size"),
        }
        for r in (p.get("resources") or [])
    ]
    return {
        "id": p.get("name") or p.get("id"),
        "title": p.get("title"),
        "notes": p.get("notes"),
        "metadata_modified": p.get("metadata_modified"),
        "tags": [t.get("name") for t in (p.get("tags") or [])],
        "resources": resources,
        "provider": "tesourotransparente.gov.br/ckan",
        "url": url,
    }
