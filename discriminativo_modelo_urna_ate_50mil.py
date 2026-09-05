#!/usr/bin/env python3
"""Entrypoint ContAgil: margens UE2020 vs antigas em municípios até 50 mil hab."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


def _load_main():
    root = Path(__file__).resolve().parent
    candidates = [
        root / "sec_scripts" / "discriminativo_modelo_urna_ate_50mil.py",
        root / "scripts" / "discriminativo_modelo_urna_ate_50mil.py",
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
            "discriminativo_modelo_urna_ate_50mil_main", path
        )
        if spec is None or spec.loader is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.main
    raise FileNotFoundError("Use: python discriminativo_modelo_urna_ate_50mil.py")


if __name__ == "__main__":
    raise SystemExit(_load_main()())
