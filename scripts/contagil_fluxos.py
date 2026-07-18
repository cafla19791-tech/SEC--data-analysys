#!/usr/bin/env python3
"""
Entrypoint no estilo do script ContAgil/RFB:

  - carrega SELIC STP (col A = data, col E = fator acumulado)
  - capitaliza com calcular_impacto_fiscal_real (dia seguinte → 30/06/2026)
  - gera fluxos e grava output/fluxos_completos_final.xlsx

Uso:
  PYTHONPATH=. python3 scripts/contagil_fluxos.py --input data/sample_operacoes_com_agente.csv
  PYTHONPATH=. python3 scripts/contagil_fluxos.py --excel caminho/operacoes.xlsx
  PYTHONPATH=. python3 scripts/contagil_fluxos.py --arquivo-selic "STP-....xlsx"
  PYTHONPATH=. python3 scripts/contagil_fluxos.py --teste-contrato0
  PYTHONPATH=. python3 scripts/contagil_fluxos.py --download
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from scripts.gerar_fluxos import (
    CONTAGIL_SELIC_DEFAULT,
    DATA_IMPACTO,
    OUTPUT_DIR,
    SelicSerie,
    calcular_impacto_fiscal_real,
    carregar_selic_serie,
    gerar_fluxos,
    load_from_csv,
    load_from_excel,
    main as gerar_fluxos_main,
    resolver_excel_operacoes,
)


def teste_contrato0(serie: SelicSerie) -> float:
    """Validação ContAgil: subsidio=1886.11 em 15/02/2009."""
    subsidio = 1886.11
    data_parcela = datetime(2009, 2, 15)
    impacto = calcular_impacto_fiscal_real(subsidio, data_parcela, serie)
    print(f"Contrato 0 — subsidio={subsidio} data={data_parcela.date()}")
    print(f"Impacto Fiscal (fatores ContAgil): R$ {impacto:,.2f}")
    data_proxima = data_parcela + timedelta(days=1)
    idx_inicio = serie.idx_proximo(data_proxima)
    idx_fim = serie.idx_proximo(DATA_IMPACTO)
    if idx_fim > idx_inicio:
        fator = serie.fatores[idx_fim] / serie.fatores[idx_inicio]
        print(f"  fator = {fator:.6f}  (idx {idx_inicio} → {idx_fim})")
    return impacto


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--excel", type=Path, help="Excel de operações (header=5).")
    p.add_argument("--input", type=Path, help="CSV de operações (sep=';).")
    p.add_argument(
        "--download",
        action="store_true",
        help="Baixa/filtra CSV aberto BNDES 2009–2010 e processa em lotes.",
    )
    p.add_argument(
        "--arquivo-selic",
        type=Path,
        default=None,
        help=f"STP ContAgil (default auto: {CONTAGIL_SELIC_DEFAULT.name} / Bacen).",
    )
    p.add_argument("--stem", default="fluxos_completos_final")
    p.add_argument("--max-contratos", type=int, default=None)
    p.add_argument(
        "--teste-contrato0",
        action="store_true",
        help="Só valida impacto ContAgil do contrato 0 (1886.11 @ 15/02/2009).",
    )
    p.add_argument(
        "--sem-selic-fatores",
        action="store_true",
        help="Usa SELIC 14,5%% composta constante (não ContAgil).",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print("🚀 Gerando fluxos...")

    # Pipeline completo (lotes + resumo) via gerar_fluxos.main
    if args.download or (
        not args.teste_contrato0
        and args.excel is None
        and args.input is None
        and resolver_excel_operacoes() is None
    ):
        cli = ["--stem", args.stem, "--download"]
        if args.arquivo_selic is not None:
            cli += ["--arquivo-selic", str(args.arquivo_selic)]
        if args.sem_selic_fatores:
            cli.append("--sem-selic-fatores")
        if args.max_contratos is not None:
            cli += ["--max-contratos", str(args.max_contratos)]
        return gerar_fluxos_main(cli)

    ns = argparse.Namespace(
        arquivo_selic=args.arquivo_selic,
        baixar_selic=not args.sem_selic_fatores,
        sem_selic_fatores=args.sem_selic_fatores,
    )
    serie = carregar_selic_serie(ns)

    if args.teste_contrato0:
        if serie is None:
            raise RuntimeError("--teste-contrato0 exige fatores SELIC.")
        teste_contrato0(serie)
        print("✅ Concluído!")
        return 0

    if args.excel:
        print(f"Lendo Excel: {args.excel}")
        df = load_from_excel(args.excel)
    elif args.input:
        print(f"Lendo CSV: {args.input}")
        df = load_from_csv(args.input)
    else:
        excel = resolver_excel_operacoes()
        if excel is None:
            raise FileNotFoundError(
                "Excel de operações não encontrado. Passe --excel, --input ou --download."
            )
        print(f"Lendo Excel: {excel}")
        df = load_from_excel(excel)

    if args.max_contratos is not None:
        df = df.head(args.max_contratos).copy()
        df["contrato"] = df.index

    # Equivalente ContAgil: df_fluxos = gerar_fluxos(df, selic)
    if serie is not None:
        selic_df = pd.DataFrame(
            {
                "data": pd.to_datetime(serie.datas),
                "b": pd.NA,
                "c": pd.NA,
                "d": pd.NA,
                "fator_acumulado": serie.fatores,
            }
        )
        df_fluxos = gerar_fluxos(df, selic_df)
    else:
        df_fluxos = gerar_fluxos(df)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_DIR / f"{args.stem}.xlsx"
    df_fluxos.to_excel(out, index=False)
    print(f"✅ Concluído! → {out} ({len(df_fluxos):,} parcelas)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
