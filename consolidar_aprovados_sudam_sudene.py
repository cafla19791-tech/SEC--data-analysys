#!/usr/bin/env python3
"""Entrypoint — espelho de ``scripts/consolidar_aprovados_sudam_sudene.py``."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.consolidar_aprovados_sudam_sudene import main

if __name__ == "__main__":
    raise SystemExit(main())
