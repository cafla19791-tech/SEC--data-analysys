#!/usr/bin/env python3
"""Entrypoint ContAgil-safe: numera contratos indiretas N-AAAA.

Evita a pasta WinPython\\Scripts (conflito case-insensitive com 'scripts').
Carrega código de sec_scripts\\ (ContAgil) ou scripts\\ (repo/dev).
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


def _load_main():
    root = Path(__file__).resolve().parent
    candidates = [
        root / "sec_scripts" / "numerar_contratos_indiretas.py",
        root / "scripts" / "numerar_contratos_indiretas.py",
    ]
    for path in candidates:
        if not path.exists():
            continue
        pkg_dir = str(path.parent)
        if pkg_dir not in sys.path:
            sys.path.insert(0, pkg_dir)
        if str(root) not in sys.path:
            sys.path.insert(0, str(root))
        spec = importlib.util.spec_from_file_location("numerar_contratos_indiretas_main", path)
        if spec is None or spec.loader is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.main
    raise FileNotFoundError(
        "Não achei numerar_contratos_indiretas.py em sec_scripts\\ nem scripts\\. "
        "Rode baixar_numerar_contratos_indiretas.ps1 de novo."
    )


if __name__ == "__main__":
    raise SystemExit(_load_main()())
