#!/usr/bin/env python3
"""Top 100 diretas BNDES 2019–2022 (IPCA 31/07/2026)."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.top100_diretas_2003_2018_ipca import processar


def main(argv: list[str] | None = None) -> int:
    saida = (
        ROOT
        / "output"
        / "top100_diretas"
        / "TOP100_DIRETAS_2019_2022_IPCA_JUL2026.xlsx"
    )
    fonte = ROOT / "data" / "bndes_naoautomaticas" / "naoautomaticas.xlsx"
    try:
        ranking = processar(
            fonte=fonte,
            saida=saida,
            n=100,
            ano_ini=2019,
            ano_fim=2022,
            data_ref=datetime(2026, 7, 31),
            baixar=True,
        )
        print(ranking.head(10).to_string(index=False))
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
