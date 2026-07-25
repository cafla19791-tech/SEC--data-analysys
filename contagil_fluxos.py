#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Entrypoint ContAgil/WinPython - espelho de scripts/contagil_fluxos.py.

Este arquivo e PYTHON. Nao cole aqui o conteudo de contagil_fluxos_bndes.bat.

Uso (uma linha):
  python contagil_fluxos.py --massa-dados dados --pasta-saida saida --arquivo-fatores fator_acumulado_SELIC_TJLP_TLP.xlsx
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
