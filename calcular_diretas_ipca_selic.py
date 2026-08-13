#!/usr/bin/env python3
"""Entrypoint ContAgil — espelho de ``scripts/calcular_diretas_ipca_selic.py``."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_SCRIPT = ROOT / "scripts" / "calcular_diretas_ipca_selic.py"
if not _SCRIPT.exists():
    sys.stderr.write(
        "ERRO: scripts\\calcular_diretas_ipca_selic.py nao encontrado.\n"
        "Rode deploy_contagil_winpython.bat a partir do repositorio.\n"
    )
    raise SystemExit(2)

from scripts.calcular_diretas_ipca_selic import main

if __name__ == "__main__":
    raise SystemExit(main())
