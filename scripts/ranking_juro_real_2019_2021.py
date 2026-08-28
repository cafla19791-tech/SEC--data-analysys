#!/usr/bin/env python3
"""Atalho: ranking 1/1/2019–31/12/2021."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.ranking_juro_real_acumulado import main as _main


def main(argv: list[str] | None = None) -> int:
    return _main(["--ano-inicio", "2019", "--ano-fim", "2021", *(argv or [])])


if __name__ == "__main__":
    raise SystemExit(main())
