#!/usr/bin/env python3
"""Entrypoint ContAgil/WinPython — espelho de ``scripts/resumo_fluxos_polars.py``.

Capitalização mensal SELIC + TJLP + TLP (fator validado 30/06/2026 = 82.79354074).

Uso (na pasta winpython):
  python resumo_fluxos_polars.py \\
      --pasta "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
      --original "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx" \\
      --selic "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\selic_mensal.xlsx" \\
      --tjlp  "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\tjlp_mensal.xlsx" \\
      --tlp   "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\tlp_mensal.xlsx"
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.resumo_fluxos_polars import main

if __name__ == "__main__":
    raise SystemExit(main())
