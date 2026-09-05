#!/usr/bin/env python3
"""Entrypoint ContAgil: discriminativo municipal UE2020 vs anteriores."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


def _load_main():
    root = Path(__file__).resolve().parent
    candidates = [
        root / "sec_scripts" / "discriminativo_urnas_municipio.py",
        root / "scripts" / "discriminativo_urnas_municipio.py",
    ]
    for path in candidates:
        if not path.exists():
            continue
        pkg_dir = str(path.parent)
        if pkg_dir not in sys.path:
            sys.path.insert(0, pkg_dir)
        spec = importlib.util.spec_from_file_location(
            "discriminativo_urnas_municipio_main", path
        )
        if spec is None or spec.loader is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.main
    raise FileNotFoundError(
        "Falta discriminativo_urnas_municipio.py em sec_scripts\\ ou scripts\\.\n"
        "Use: python discriminativo_urnas_municipio.py"
    )


if __name__ == "__main__":
    raise SystemExit(_load_main()())
