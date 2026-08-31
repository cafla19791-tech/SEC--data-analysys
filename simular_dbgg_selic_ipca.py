#!/usr/bin/env python3
"""Entrypoint — espelho de ``scripts/simular_dbgg_selic_ipca.py``."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_SCRIPT = ROOT / "scripts" / "simular_dbgg_selic_ipca.py"
if not _SCRIPT.exists():
    sys.stderr.write("ERRO: scripts/simular_dbgg_selic_ipca.py nao encontrado.\n")
    raise SystemExit(2)

from scripts.simular_dbgg_selic_ipca import main

if __name__ == "__main__":
    raise SystemExit(main())
