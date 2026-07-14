"""Cliente mínimo para a API pública SEC EDGAR (CompanyFacts)."""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

import requests

SEC_DATA_BASE = "https://data.sec.gov"
PETROBRAS_CIK = 1119639


class SecClient:
    """Busca fatos financeiros XBRL na SEC EDGAR com cache local opcional."""

    def __init__(
        self,
        user_agent: str = "SEC-Data-Analysis cafla19791@gmail.com",
        min_interval: float = 0.2,
        max_retries: int = 5,
    ) -> None:
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate",
            "Accept": "application/json",
        }
        self.min_interval = min_interval
        self.max_retries = max_retries
        self._last_request = 0.0

    def _throttle(self) -> None:
        elapsed = time.monotonic() - self._last_request
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self._last_request = time.monotonic()

    def _get_json(self, url: str) -> dict[str, Any]:
        last_error: Exception | None = None
        for attempt in range(self.max_retries):
            self._throttle()
            response = requests.get(url, headers=self.headers, timeout=120)
            if response.status_code == 200:
                return response.json()
            if response.status_code in {403, 429, 502, 503}:
                last_error = requests.HTTPError(
                    f"{response.status_code} for {url}",
                    response=response,
                )
                time.sleep(1.5 * (attempt + 1))
                continue
            response.raise_for_status()
        assert last_error is not None
        raise last_error

    def get_company_facts(
        self,
        cik: int | str = PETROBRAS_CIK,
        cache_path: str | Path | None = None,
        use_cache: bool = True,
    ) -> dict[str, Any]:
        """Baixa CompanyFacts; usa cache em disco se disponível."""
        path = Path(cache_path) if cache_path else None
        if use_cache and path and path.exists():
            import json

            return json.loads(path.read_text(encoding="utf-8"))

        cik_padded = str(cik).zfill(10)
        url = f"{SEC_DATA_BASE}/api/xbrl/companyfacts/CIK{cik_padded}.json"
        data = self._get_json(url)

        if path:
            path.parent.mkdir(parents=True, exist_ok=True)
            import json

            path.write_text(json.dumps(data), encoding="utf-8")
        return data
