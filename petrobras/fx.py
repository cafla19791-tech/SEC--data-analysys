"""Câmbio médio anual USD/BRL via API do Banco Central (SGS)."""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any

import requests

BCB_SGS_USD_VENDA = 1  # Taxa de câmbio - Dólar americano (venda)


def fetch_annual_avg_usd_brl(
    years: list[int],
    cache_path: str | Path | None = None,
    use_cache: bool = True,
) -> dict[str, dict[str, Any]]:
    """
    Retorna média anual da taxa USD/BRL (venda) para cada ano.

    Fonte: Banco Central do Brasil — SGS série 1.
    """
    path = Path(cache_path) if cache_path else None
    cached: dict[str, dict[str, Any]] = {}
    if use_cache and path and path.exists():
        cached = json.loads(path.read_text(encoding="utf-8"))

    result: dict[str, dict[str, Any]] = {}
    for year in years:
        key = str(year)
        if key in cached and cached[key].get("avg_usd_brl"):
            result[key] = cached[key]
            continue

        url = (
            "https://api.bcb.gov.br/dados/serie/bcdata.sgs."
            f"{BCB_SGS_USD_VENDA}/dados"
            f"?formato=json&dataInicial=01/01/{year}&dataFinal=31/12/{year}"
        )
        data = _get_with_retry(url)
        values = [float(row["valor"].replace(",", ".")) for row in data]
        if not values:
            raise RuntimeError(f"Sem observações de câmbio para {year}")
        result[key] = {
            "avg_usd_brl": sum(values) / len(values),
            "n_obs": len(values),
            "source": "BCB SGS 1 (USD venda)",
        }
        time.sleep(0.15)

    if path:
        path.parent.mkdir(parents=True, exist_ok=True)
        merged = {**cached, **result}
        path.write_text(json.dumps(merged, indent=2), encoding="utf-8")

    return result


def _get_with_retry(url: str, retries: int = 4) -> list[dict[str, Any]]:
    last: Exception | None = None
    for attempt in range(retries):
        try:
            response = requests.get(url, timeout=90)
            response.raise_for_status()
            return response.json()
        except Exception as exc:  # noqa: BLE001 — rede instável do BCB
            last = exc
            time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"Falha ao consultar BCB: {url}") from last
