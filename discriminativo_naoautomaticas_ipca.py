#!/usr/bin/env python3
"""Entrypoint ContAgil — espelho de ``scripts/discriminativo_naoautomaticas_ipca.py``."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.discriminativo_naoautomaticas_ipca import main

if __name__ == "__main__":
    raise SystemExit(main())
