#!/usr/bin/env python3
"""Entrypoint — espelho de ``scripts/fatores_condicionantes_base_monetaria.py``."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_SCRIPT = ROOT / "scripts" / "fatores_condicionantes_base_monetaria.py"
if not _SCRIPT.exists():
    sys.stderr.write(
        "ERRO: scripts/fatores_condicionantes_base_monetaria.py nao encontrado.\n"
    )
    raise SystemExit(2)

from scripts.fatores_condicionantes_base_monetaria import main

if __name__ == "__main__":
    raise SystemExit(main())
