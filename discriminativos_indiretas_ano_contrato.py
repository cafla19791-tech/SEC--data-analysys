#!/usr/bin/env python3
"""Entrypoint ContAgil: discriminativos de indiretas por ano do contrato."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


def _load_main():
    root = Path(__file__).resolve().parent
    candidates = [
        root / "sec_scripts" / "discriminativos_indiretas_ano_contrato.py",
        root / "scripts" / "discriminativos_indiretas_ano_contrato.py",
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
            "discriminativos_indiretas_ano_contrato_main", path
        )
        if spec is None or spec.loader is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.main
    raise FileNotFoundError(
        "Nao achei discriminativos_indiretas_ano_contrato.py em "
        "sec_scripts\\ nem scripts\\."
    )


if __name__ == "__main__":
    raise SystemExit(_load_main()())
