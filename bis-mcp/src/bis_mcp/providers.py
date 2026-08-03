"""BIS WS_CBPOL client: SDMX REST + local flat CSV (ContAgil WinPython).

Public sources (no API key):
- SDMX: https://stats.bis.org/api/v1/data/WS_CBPOL/...
- Bulk flat CSV: https://data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip
"""

from __future__ import annotations

import csv
import io
import os
import re
import time
import zipfile
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import quote

import httpx

SDMX_DATA = "https://stats.bis.org/api/v1/data/WS_CBPOL"
BULK_FLAT_ZIP = "https://data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip"
FLAT_CSV_NAME = "WS_CBPOL_csv_flat.csv"

DEFAULT_UA = os.getenv(
    "BIS_USER_AGENT",
    "SEC-data-analysys-bis-mcp/0.1 (cafla19791@gmail.com)",
)

_MIN_INTERVAL = 0.25
_last_request = 0.0

# Common ContAgil / analysis aliases -> ISO REF_AREA.
KNOWN_AREAS: dict[str, dict[str, str]] = {
    "br": {"code": "BR", "name": "Brazil (SELIC meta / policy rate)"},
    "brasil": {"code": "BR", "name": "Brazil (SELIC meta / policy rate)"},
    "brazil": {"code": "BR", "name": "Brazil (SELIC meta / policy rate)"},
    "selic": {"code": "BR", "name": "Brazil (SELIC meta / policy rate)"},
    "us": {"code": "US", "name": "United States"},
    "eua": {"code": "US", "name": "United States"},
    "usa": {"code": "US", "name": "United States"},
    "fed": {"code": "US", "name": "United States"},
    "xm": {"code": "XM", "name": "Euro area"},
    "euro": {"code": "XM", "name": "Euro area"},
    "eurozona": {"code": "XM", "name": "Euro area"},
    "ecb": {"code": "XM", "name": "Euro area"},
    "gb": {"code": "GB", "name": "United Kingdom"},
    "uk": {"code": "GB", "name": "United Kingdom"},
    "jp": {"code": "JP", "name": "Japan"},
    "japao": {"code": "JP", "name": "Japan"},
    "cn": {"code": "CN", "name": "China"},
    "china": {"code": "CN", "name": "China"},
    "ar": {"code": "AR", "name": "Argentina"},
    "mx": {"code": "MX", "name": "Mexico"},
    "cl": {"code": "CL", "name": "Chile"},
    "co": {"code": "CO", "name": "Colombia"},
    "za": {"code": "ZA", "name": "South Africa"},
    "in": {"code": "IN", "name": "India"},
    "tr": {"code": "TR", "name": "Türkiye"},
    "ch": {"code": "CH", "name": "Switzerland"},
    "ca": {"code": "CA", "name": "Canada"},
    "au": {"code": "AU", "name": "Australia"},
    "nz": {"code": "NZ", "name": "New Zealand"},
    "se": {"code": "SE", "name": "Sweden"},
    "no": {"code": "NO", "name": "Norway"},
    "kr": {"code": "KR", "name": "Korea"},
}

# ContAgil WinPython default location (Windows). Also honor env override.
_DEFAULT_LOCAL_CANDIDATES = (
    Path(os.getenv("BIS_CBPOL_CSV", "")),
    Path(FLAT_CSV_NAME),
    Path("..") / FLAT_CSV_NAME,
    Path(r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython")
    / FLAT_CSV_NAME,
)


def _headers() -> dict[str, str]:
    return {
        "User-Agent": DEFAULT_UA,
        "Accept": "text/csv,application/json,*/*",
        "Accept-Encoding": "gzip, deflate",
    }


def _throttle() -> None:
    global _last_request
    now = time.monotonic()
    wait = _MIN_INTERVAL - (now - _last_request)
    if wait > 0:
        time.sleep(wait)
    _last_request = time.monotonic()


def _get_text(url: str, *, timeout: float = 120.0) -> str:
    _throttle()
    with httpx.Client(headers=_headers(), timeout=timeout, follow_redirects=True) as client:
        resp = client.get(url)
        if resp.status_code == 404:
            raise ValueError(f"Recurso BIS nao encontrado (404): {url}")
        if resp.status_code == 429:
            raise RuntimeError("BIS rate limit (429). Aguarde e tente novamente.")
        resp.raise_for_status()
        return resp.text


def _get_bytes(url: str, *, timeout: float = 300.0) -> bytes:
    _throttle()
    with httpx.Client(headers=_headers(), timeout=timeout, follow_redirects=True) as client:
        resp = client.get(url)
        if resp.status_code == 404:
            raise ValueError(f"Recurso BIS nao encontrado (404): {url}")
        resp.raise_for_status()
        return resp.content


def list_known_areas() -> dict[str, Any]:
    """Local aliases for common REF_AREA codes."""
    items = []
    seen: set[str] = set()
    for alias, meta in sorted(KNOWN_AREAS.items()):
        code = meta["code"]
        if code in seen and alias != code.lower():
            # keep one primary alias row; still list alias mapping below
            pass
        items.append({"alias": alias, "code": code, "name": meta["name"]})
        seen.add(code)
    return {
        "dataset": "WS_CBPOL",
        "description": "BIS central bank policy rates",
        "aliases": items,
        "bulk_flat_zip": BULK_FLAT_ZIP,
        "sdmx_base": SDMX_DATA,
        "contagil_default_csv": str(
            Path(r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython")
            / FLAT_CSV_NAME
        ),
    }


def resolve_area(code_or_alias: str) -> dict[str, str]:
    raw = str(code_or_alias).strip()
    if not raw:
        raise ValueError("Informe um codigo/alias de pais (ex.: BR, brasil, US, euro).")
    key = raw.lower()
    if key in KNOWN_AREAS:
        meta = KNOWN_AREAS[key]
        return {"alias": key, "code": meta["code"], "name": meta["name"]}
    if re.fullmatch(r"[A-Za-z]{2}", raw):
        code = raw.upper()
        name = KNOWN_AREAS.get(code.lower(), {}).get("name", code)
        return {"alias": key, "code": code, "name": name}
    raise ValueError(
        f"Area desconhecida: {code_or_alias!r}. Use catalog ou codigos ISO (BR, US, XM)."
    )


def _split_areas(areas: str | Iterable[str]) -> list[dict[str, str]]:
    if isinstance(areas, str):
        parts = re.split(r"[\s,;+|]+", areas.strip())
    else:
        parts = [str(a) for a in areas]
    parts = [p for p in parts if p]
    if not parts:
        raise ValueError("Informe ao menos uma area (ex.: BR,US,XM).")
    resolved = [resolve_area(p) for p in parts]
    # unique preserve order
    out: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in resolved:
        if item["code"] not in seen:
            out.append(item)
            seen.add(item["code"])
    return out


def _normalize_freq(freq: str) -> str:
    f = (freq or "M").strip().upper()
    if f in {"M", "MONTHLY", "MES", "MENSAL"}:
        return "M"
    if f in {"D", "DAILY", "DIA", "DIARIO", "DIÁRIO"}:
        return "D"
    if f in {"A", "Y", "ANNUAL", "ANO", "ANUAL"}:
        return "A"
    raise ValueError(f"Frequencia invalida: {freq!r}. Use M, D ou A.")


def _parse_obs(value: str | None) -> float | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    # Flat CSV sometimes uses "11.75: something" — keep left token.
    if ":" in s and not s.replace(".", "", 1).replace("-", "", 1).isdigit():
        left = s.split(":", 1)[0].strip()
        if left:
            s = left
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _norm_key(key: str | None) -> str:
    """Normalize SDMX/flat headers: 'REF_AREA:Reference area' -> 'REF_AREA'."""
    if not key:
        return ""
    k = key.strip().upper()
    if ":" in k:
        k = k.split(":", 1)[0].strip()
    return k.replace(" ", "_")


def _norm_code(value: str | None) -> str:
    """Normalize coded cells: 'BR: Brazil' / 'M: Monthly' -> 'BR' / 'M'."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    if ":" in s:
        return s.split(":", 1)[0].strip()
    return s


def _row_dict(raw: dict[str, Any]) -> dict[str, str]:
    out: dict[str, str] = {}
    for k, v in raw.items():
        key = _norm_key(k)
        if not key:
            continue
        out[key] = v.strip() if isinstance(v, str) else ("" if v is None else str(v))
    return out


def _rows_from_csv_text(text: str) -> list[dict[str, Any]]:
    # SDMX CSV may embed commas inside quoted compilation fields.
    reader = csv.DictReader(io.StringIO(text))
    rows: list[dict[str, Any]] = []
    for raw in reader:
        if not raw:
            continue
        data = _row_dict(raw)
        period = data.get("TIME_PERIOD") or ""
        area = _norm_code(data.get("REF_AREA"))
        freq = _norm_code(data.get("FREQ"))
        obs = _parse_obs(data.get("OBS_VALUE"))
        if not period or not area:
            continue
        if obs is None:
            continue
        rows.append(
            {
                "freq": freq,
                "ref_area": area,
                "time_period": period,
                "value": obs,
                "title": data.get("TITLE") or "",
                "obs_status": _norm_code(data.get("OBS_STATUS")),
            }
        )
    rows.sort(key=lambda r: (r["ref_area"], r["time_period"]))
    return rows


def find_local_flat_csv(explicit: str | Path | None = None) -> Path | None:
    """Locate WS_CBPOL_csv_flat.csv (ContAgil winpython or CWD)."""
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit))
    candidates.extend(p for p in _DEFAULT_LOCAL_CANDIDATES if str(p))
    for path in candidates:
        try:
            if path.is_file():
                return path.resolve()
        except OSError:
            continue
    return None


def download_flat_csv(
    dest_dir: str | Path,
    *,
    overwrite: bool = True,
) -> dict[str, Any]:
    """Download BIS bulk flat zip and extract WS_CBPOL_csv_flat.csv."""
    dest = Path(dest_dir)
    dest.mkdir(parents=True, exist_ok=True)
    csv_path = dest / FLAT_CSV_NAME
    if csv_path.exists() and not overwrite:
        return {
            "path": str(csv_path.resolve()),
            "bytes": csv_path.stat().st_size,
            "source": "existing",
            "url": BULK_FLAT_ZIP,
        }

    blob = _get_bytes(BULK_FLAT_ZIP)
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        names = zf.namelist()
        target = next((n for n in names if n.endswith(FLAT_CSV_NAME) or n == FLAT_CSV_NAME), None)
        if not target:
            # fallback: first csv
            target = next((n for n in names if n.lower().endswith(".csv")), None)
        if not target:
            raise RuntimeError(f"ZIP sem CSV. Conteudo: {names[:20]}")
        with zf.open(target) as src, open(csv_path, "wb") as out:
            out.write(src.read())

    return {
        "path": str(csv_path.resolve()),
        "bytes": csv_path.stat().st_size,
        "source": "downloaded",
        "url": BULK_FLAT_ZIP,
        "zip_member": target,
    }


def read_local_series(
    areas: str | Iterable[str],
    *,
    csv_path: str | Path | None = None,
    freq: str = "M",
    date_from: str | None = None,
    date_to: str | None = None,
    last: int | None = None,
) -> dict[str, Any]:
    """Read policy rates from a local WS_CBPOL_csv_flat.csv."""
    resolved = _split_areas(areas)
    freq_code = _normalize_freq(freq)
    path = find_local_flat_csv(csv_path)
    if path is None:
        raise FileNotFoundError(
            f"{FLAT_CSV_NAME} nao encontrado. Baixe com: bis-cli download "
            f"ou defina BIS_CBPOL_CSV. ContAgil tipico: "
            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\WS_CBPOL_csv_flat.csv"
        )

    wanted = {a["code"] for a in resolved}
    rows: list[dict[str, Any]] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for raw in reader:
            data = _row_dict(raw)
            area = _norm_code(data.get("REF_AREA"))
            f = _norm_code(data.get("FREQ"))
            if area not in wanted:
                continue
            if f and f != freq_code:
                continue
            period = (data.get("TIME_PERIOD") or "").strip()
            obs = _parse_obs(data.get("OBS_VALUE"))
            if not period or obs is None:
                continue
            if date_from and period < date_from:
                continue
            if date_to and period > date_to:
                continue
            rows.append(
                {
                    "freq": f or freq_code,
                    "ref_area": area,
                    "time_period": period,
                    "value": obs,
                    "title": data.get("TITLE") or "",
                    "obs_status": _norm_code(data.get("OBS_STATUS")),
                }
            )

    rows.sort(key=lambda r: (r["ref_area"], r["time_period"]))
    if last and last > 0:
        # last N per area
        by_area: dict[str, list[dict[str, Any]]] = {}
        for row in rows:
            by_area.setdefault(row["ref_area"], []).append(row)
        trimmed: list[dict[str, Any]] = []
        for code in sorted(by_area):
            trimmed.extend(by_area[code][-last:])
        rows = trimmed

    return {
        "source": "local_csv",
        "path": str(path),
        "dataset": "WS_CBPOL",
        "freq": freq_code,
        "areas": resolved,
        "count": len(rows),
        "series": rows,
    }


def get_policy_rates(
    areas: str | Iterable[str],
    *,
    freq: str = "M",
    date_from: str | None = None,
    date_to: str | None = None,
    last: int | None = None,
    prefer_local: bool = False,
    csv_path: str | Path | None = None,
) -> dict[str, Any]:
    """Fetch policy rates via SDMX (default) or local flat CSV."""
    if prefer_local or csv_path:
        return read_local_series(
            areas,
            csv_path=csv_path,
            freq=freq,
            date_from=date_from,
            date_to=date_to,
            last=last,
        )

    resolved = _split_areas(areas)
    freq_code = _normalize_freq(freq)
    key = f"{freq_code}.{'+'.join(a['code'] for a in resolved)}"
    # detail=dataonly keeps payload small; full includes TITLE/COMPILATION.
    params = ["detail=dataonly", "format=csv"]
    if date_from:
        params.append(f"startPeriod={quote(str(date_from))}")
    if date_to:
        params.append(f"endPeriod={quote(str(date_to))}")
    url = f"{SDMX_DATA}/{key}?{'&'.join(params)}"

    try:
        text = _get_text(url)
        rows = _rows_from_csv_text(text)
    except Exception:
        # Offline / blocked SDMX: fall back to local ContAgil CSV if present.
        local = find_local_flat_csv(csv_path)
        if local is None:
            raise
        return read_local_series(
            areas,
            csv_path=local,
            freq=freq,
            date_from=date_from,
            date_to=date_to,
            last=last,
        )

    if last and last > 0:
        by_area: dict[str, list[dict[str, Any]]] = {}
        for row in rows:
            by_area.setdefault(row["ref_area"], []).append(row)
        trimmed = []
        for code in [a["code"] for a in resolved]:
            trimmed.extend(by_area.get(code, [])[-last:])
        rows = trimmed

    return {
        "source": "sdmx",
        "url": url,
        "dataset": "WS_CBPOL",
        "freq": freq_code,
        "areas": resolved,
        "count": len(rows),
        "series": rows,
    }


def compare_latest(
    areas: str | Iterable[str],
    *,
    freq: str = "M",
    prefer_local: bool = False,
    csv_path: str | Path | None = None,
) -> dict[str, Any]:
    """Latest observation per area (useful ContAgil snapshot)."""
    data = get_policy_rates(
        areas,
        freq=freq,
        last=1,
        prefer_local=prefer_local,
        csv_path=csv_path,
    )
    latest = []
    for area in data["areas"]:
        pts = [r for r in data["series"] if r["ref_area"] == area["code"]]
        if pts:
            p = pts[-1]
            latest.append(
                {
                    "code": area["code"],
                    "name": area["name"],
                    "time_period": p["time_period"],
                    "value": p["value"],
                }
            )
        else:
            latest.append(
                {
                    "code": area["code"],
                    "name": area["name"],
                    "time_period": None,
                    "value": None,
                }
            )
    return {
        "dataset": "WS_CBPOL",
        "freq": data["freq"],
        "source": data["source"],
        "latest": latest,
    }


def extract_areas_csv(
    areas: str | Iterable[str],
    out_path: str | Path,
    *,
    csv_path: str | Path | None = None,
    freq: str | None = "M",
) -> dict[str, Any]:
    """Extract selected countries from the huge flat CSV into a slim file.

    ContAgil tip: extract BR (or BR,US,XM) once, then use --local --csv on the slim file.
    """
    resolved = _split_areas(areas)
    wanted = {a["code"] for a in resolved}
    freq_code = _normalize_freq(freq) if freq else None
    src = find_local_flat_csv(csv_path)
    if src is None:
        raise FileNotFoundError(
            f"{FLAT_CSV_NAME} nao encontrado para extrair. Rode: bis-cli download"
        )

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = ["FREQ", "REF_AREA", "TIME_PERIOD", "OBS_VALUE", "TITLE", "OBS_STATUS"]
    n = 0
    with open(src, "r", encoding="utf-8-sig", newline="") as fh_in, open(
        out, "w", encoding="utf-8", newline=""
    ) as fh_out:
        reader = csv.DictReader(fh_in)
        writer = csv.DictWriter(fh_out, fieldnames=fieldnames)
        writer.writeheader()
        for raw in reader:
            data = _row_dict(raw)
            area = _norm_code(data.get("REF_AREA"))
            f = _norm_code(data.get("FREQ"))
            if area not in wanted:
                continue
            if freq_code and f and f != freq_code:
                continue
            period = (data.get("TIME_PERIOD") or "").strip()
            obs = data.get("OBS_VALUE") or ""
            if not period or not str(obs).strip():
                continue
            writer.writerow(
                {
                    "FREQ": f,
                    "REF_AREA": area,
                    "TIME_PERIOD": period,
                    "OBS_VALUE": _norm_code(obs) if ":" in str(obs) else obs,
                    "TITLE": data.get("TITLE") or "",
                    "OBS_STATUS": _norm_code(data.get("OBS_STATUS")),
                }
            )
            n += 1

    return {
        "source": str(src),
        "path": str(out.resolve()),
        "areas": resolved,
        "freq": freq_code,
        "rows": n,
        "bytes": out.stat().st_size,
    }
