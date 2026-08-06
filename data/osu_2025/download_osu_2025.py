#!/usr/bin/env python3
"""Download the official OSU 2025 anexos spreadsheet from gov.br."""

from __future__ import annotations

import urllib.request
from pathlib import Path

URL = (
    "https://www.gov.br/planejamento/pt-br/assuntos/avaliacao-de-politicas-publicas/"
    "arquivos/orcamento-de-subsidios-da-uniao/osu_2025-anexos-publicacao.xlsx/"
    "@@download/file"
)
OUT = Path(__file__).resolve().parent / "OSU_2025_anexos_fonte.xlsx"


def main() -> None:
    req = urllib.request.Request(URL, headers={"User-Agent": "SEC-data-analysys/1.0"})
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = resp.read()
    OUT.write_bytes(data)
    print(f"Saved {OUT} ({len(data):,} bytes)")


if __name__ == "__main__":
    main()
