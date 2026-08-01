#!/usr/bin/env python3
"""Entrypoint ContAgil/WinPython — espelho de ``scripts/resumo_fluxos_polars.py``.

IMPORTANTE: este arquivo precisa da pasta ``scripts/`` ao lado (repo completo).
Se estiver na WinPython sem o repo:

  1) No repositorio clonado, rode: deploy_contagil_winpython.bat
  2) Ou rode a partir do clone:

       cd C:\\caminho\\SEC--data-analysys
       python scripts\\resumo_fluxos_polars.py --pasta "...\\saida" ...

Uso (apos deploy, na pasta winpython):
  python resumo_fluxos_polars.py \\
      --pasta "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
      --original "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\dados\\BNDES INDIRETAS 2002.xlsx" \\
      --selic "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\selic_mensal.xlsx" \\
      --tjlp "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\tjlp_mensal.xlsx"
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_SCRIPT = ROOT / "scripts" / "resumo_fluxos_polars.py"
if not _SCRIPT.exists():
    sys.stderr.write(
        "ERRO: nao encontrado scripts\\resumo_fluxos_polars.py em:\n"
        f"  {ROOT}\n\n"
        "A pasta ContAgil WinPython nao traz o repositorio.\n"
        "Solucao:\n"
        "  1) Clone/atualize SEC--data-analysys\n"
        "  2) No repo, execute: deploy_contagil_winpython.bat\n"
        "  3) Ou rode o comando a partir da pasta do repo.\n"
    )
    raise SystemExit(2)

from scripts.resumo_fluxos_polars import main

if __name__ == "__main__":
    raise SystemExit(main())
