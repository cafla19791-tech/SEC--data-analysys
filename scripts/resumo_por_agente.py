#!/usr/bin/env python3
"""
Resumo por Agente Financeiro — versão CLI / apoio à Web.

Corrige o script de referência:
  - header=5 (não cabeçalho=5)
  - ascending=False (não ascendente=False)
  - vínculo contrato → agente (não merge por índice em CSV de parcelas)

Uso:
  # A partir dos contratos + fluxos já gerados
  python scripts/resumo_por_agente.py \\
      --contratos data/operacoes_indiretas_automaticas_2009-2010.csv \\
      --fluxos output/fluxos_completos_corrigido.csv

  # Ou só imprime o ranking já salvo
  python scripts/resumo_por_agente.py --from-output
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

from scripts.gerar_fluxos import (
    OUTPUT_DIR,
    agregar_por_agente,
    load_from_csv,
    load_from_excel,
    salvar_resumo_por_agente,
)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--from-output",
        action="store_true",
        help="Lê output/resumo_por_agente.csv já gerado e imprime o top N.",
    )
    p.add_argument(
        "--contratos",
        type=Path,
        help="CSV filtrado BNDES (sep=;) com instituicao_financeira_credenciada.",
    )
    p.add_argument(
        "--excel",
        type=Path,
        help="Excel do portal (header=5) com Instituição Financeira Credenciada.",
    )
    p.add_argument(
        "--fluxos",
        type=Path,
        default=OUTPUT_DIR / "fluxos_completos_final.csv",
        help="CSV detalhado de parcelas (default: output/fluxos_completos_final.csv).",
    )
    p.add_argument("--top", type=int, default=20, help="Linhas a imprimir (default 20).")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print("Resumo por Agente Financeiro - Versão Web / CLI")

    out_csv = OUTPUT_DIR / "resumo_por_agente.csv"
    if args.from_output:
        if not out_csv.exists():
            print(f"Arquivo não encontrado: {out_csv}", file=sys.stderr)
            print("Gere com: python scripts/gerar_fluxos.py --download", file=sys.stderr)
            return 1
        resumo = pd.read_csv(out_csv)
        print(resumo.head(args.top).to_string(index=False))
        return 0

    if args.excel:
        contratos = load_from_excel(args.excel)
    elif args.contratos:
        contratos = load_from_csv(args.contratos)
    else:
        print("Informe --contratos, --excel ou --from-output.", file=sys.stderr)
        return 2

    if not args.fluxos.exists():
        print(f"Fluxos não encontrados: {args.fluxos}", file=sys.stderr)
        return 1

    # Leitura em chunks se o CSV for enorme
    parts: list[pd.DataFrame] = []
    for chunk in pd.read_csv(args.fluxos, chunksize=500_000):
        impacto_col = (
            "impacto_fiscal"
            if "impacto_fiscal" in chunk.columns
            else "impacto"
        )
        cols = ["contrato", "subsidio", impacto_col]
        if "Instituição Financeira" in chunk.columns:
            cols.append("Instituição Financeira")
        parts.append(chunk[cols])
    fluxos = pd.concat(parts, ignore_index=True)

    resumo = agregar_por_agente(fluxos, contratos)
    csv_path, xlsx_path = salvar_resumo_por_agente(resumo)

    print(resumo.head(args.top).to_string(index=False))
    print(f"\n✅ {csv_path}")
    print(f"✅ {xlsx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
