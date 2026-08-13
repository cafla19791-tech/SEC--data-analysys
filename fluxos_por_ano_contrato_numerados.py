#!/usr/bin/env python3
"""Entrypoint ContAgil: fluxos por aba de ano a partir de BNDES_INDIRETAS_NUMERADOS."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


def _load_main():
    root = Path(__file__).resolve().parent
    for path in (
        root / "sec_scripts" / "fluxos_por_ano_contrato_numerados.py",
        root / "scripts" / "fluxos_por_ano_contrato_numerados.py",
    ):
        if not path.exists():
            continue
        for p in (str(path.parent), str(root)):
            if p not in sys.path:
                sys.path.insert(0, p)
        spec = importlib.util.spec_from_file_location(
            "fluxos_por_ano_contrato_numerados_main", path
        )
        if spec is None or spec.loader is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.main
    raise FileNotFoundError(
        "Nao achei fluxos_por_ano_contrato_numerados.py em sec_scripts\\ nem scripts\\."
    )


if __name__ == "__main__":
    raise SystemExit(_load_main()())
