#!/usr/bin/env python3
"""Entrypoint ContAgil — espelho de ``scripts/discriminativo_ranking_juros_reais.py``."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_SCRIPT = ROOT / "scripts" / "discriminativo_ranking_juros_reais.py"
if not _SCRIPT.exists():
    sys.stderr.write(
        "ERRO: scripts\\discriminativo_ranking_juros_reais.py nao encontrado.\n"
    )
    raise SystemExit(2)

from scripts.discriminativo_ranking_juros_reais import main

if __name__ == "__main__":
    raise SystemExit(main())
