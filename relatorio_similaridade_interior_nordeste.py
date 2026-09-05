#!/usr/bin/env python3
"""Entrypoint ContAgil: relatório de similaridade no interior do Nordeste."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


def _load_main():
    root = Path(__file__).resolve().parent
    candidates = [
        root / "sec_scripts" / "relatorio_similaridade_interior_nordeste.py",
        root / "scripts" / "relatorio_similaridade_interior_nordeste.py",
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
            "relatorio_similaridade_interior_nordeste_main", path
        )
        if spec is None or spec.loader is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.main
    raise FileNotFoundError("Use: python relatorio_similaridade_interior_nordeste.py")


if __name__ == "__main__":
    raise SystemExit(_load_main()())
