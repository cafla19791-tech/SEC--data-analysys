#!/usr/bin/env python3
"""Entrypoint — espelho de ``scripts/top100_diretas_2003_2018_ipca.py``."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.top100_diretas_2003_2018_ipca import main

if __name__ == "__main__":
    raise SystemExit(main())
