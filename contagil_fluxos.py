#!/usr/bin/env python3
r"""Entrypoint ContAgil/WinPython — espelho de ``scripts/contagil_fluxos.py``.

Uso (capitalização mensal):
  python contagil_fluxos.py ^
      --massa-dados "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\dados" ^
      --pasta-saida "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida" ^
      --arquivo-fatores "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\fator_acumulado_SELIC_TJLP_TLP.xlsx"
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.contagil_fluxos import main

if __name__ == "__main__":
    raise SystemExit(main())
