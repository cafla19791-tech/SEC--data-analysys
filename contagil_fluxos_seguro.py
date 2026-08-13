#!/usr/bin/env python3
"""Entrypoint ContAgil/WinPython — espelho de ``scripts/contagil_fluxos_seguro.py``.

Uso (na pasta winpython):
  python contagil_fluxos_seguro.py

  python contagil_fluxos_seguro.py \\
      --massa-dados dados \\
      --pasta-saida saida \\
      --fatores fator_acumulado_SELIC_TJLP_TLP.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.contagil_fluxos_seguro import main

if __name__ == "__main__":
    raise SystemExit(main())
