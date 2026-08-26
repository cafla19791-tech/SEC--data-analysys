#!/usr/bin/env python3
"""Entrypoint ContAgil: baixa Boletins de Urna 2022 (28 UFs).

Evita a pasta WinPython\\Scripts (conflito com 'scripts').
Carrega o código de sec_scripts\\ (ContAgil) ou scripts\\ (repo).
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


def _load_main():
    root = Path(__file__).resolve().parent
    candidates = [
        root / "sec_scripts" / "baixar_boletins_urna_2022.py",
        root / "scripts" / "baixar_boletins_urna_2022.py",
    ]
    for path in candidates:
        if not path.exists():
            continue
        pkg_dir = str(path.parent)
        if pkg_dir not in sys.path:
            sys.path.insert(0, pkg_dir)
        if str(root) not in sys.path:
            sys.path.insert(0, str(root))
        spec = importlib.util.spec_from_file_location(
            "baixar_boletins_urna_2022_main", path
        )
        if spec is None or spec.loader is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.main
    raise FileNotFoundError(
        "No ContAgil NAO use: python scripts\\baixar_boletins_urna_2022.py\n"
        "(essa pasta e o Scripts do pip). Use:\n"
        "  python baixar_boletins_urna_2022.py\n"
        "ou duplo-clique em baixar_boletins_urna_2022.bat\n"
        "Se o arquivo nao existir, rode baixar_boletins_urna_2022.ps1 "
        "ou o bloco curl na pasta winpython."
    )


if __name__ == "__main__":
    raise SystemExit(_load_main()())
