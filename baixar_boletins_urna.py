#!/usr/bin/env python3
"""Entrypoint ContAgil: Boletins de Urna 2014/2018/2022 (1º e 2º turno).

Evita a pasta WinPython\\Scripts (conflito com 'scripts').
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


def _load_main():
    root = Path(__file__).resolve().parent
    candidates = [
        root / "sec_scripts" / "baixar_boletins_urna.py",
        root / "scripts" / "baixar_boletins_urna.py",
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
            "baixar_boletins_urna_main", path
        )
        if spec is None or spec.loader is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.main
    raise FileNotFoundError(
        "Use: python baixar_boletins_urna.py --ano 2022 --turno 1\n"
        "Na RFB: python baixar_boletins_urna.py --somente-resultado-github"
    )


if __name__ == "__main__":
    raise SystemExit(_load_main()())
