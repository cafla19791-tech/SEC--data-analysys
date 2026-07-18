#!/usr/bin/env python3
"""
Impacto fiscal por ano de pagamento — capitalizado até 30/06/2026.

Lê o CSV de parcelas (`fluxos_completos_*.csv`), calcula o impacto individual
de cada parcela e agrega por ano de `data_fluxo`.

Modos de impacto:
  - recalcular (padrão): subsídio × (1 + SELIC_aa/12)^meses  (script de referência)
  - coluna: usa `impacto_fiscal` / `impacto` já gravado no CSV (ContAgil/Bacen)
  - composta: subsídio × (1 + SELIC_m)^meses, com SELIC_m = (1+aa)^(1/12)-1

Uso:
  python3 scripts/impacto_fiscal_por_ano.py
  python3 scripts/impacto_fiscal_por_ano.py --fluxos output/fluxos_completos_corrigido.csv
  python3 scripts/impacto_fiscal_por_ano.py --modo coluna --fluxos output/fluxos_completos_final.csv
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta

from scripts.gerar_fluxos import OUTPUT_DIR, TAXA_SELIC_ANUAL, taxa_mensal_composta

DATA_REFERENCIA = datetime(2026, 6, 30)

CANDIDATOS_FLUXOS = (
    OUTPUT_DIR / "fluxos_completos_corrigido.csv",
    OUTPUT_DIR / "fluxos_completos_final.csv",
    OUTPUT_DIR / "fluxos_amostra.csv",
    Path("/tmp/app-streamlit/output/fluxos_completos_corrigido.csv"),
    Path("/tmp/app/output/fluxos_amostra.csv"),
)

COLUNAS_IMPACTO = ("impacto_fiscal", "impacto")


def resolver_fluxos(explicit: Path | None = None) -> Path:
    if explicit is not None:
        return explicit
    for path in CANDIDATOS_FLUXOS:
        if path.exists() and path.is_file():
            return path
    raise FileNotFoundError(
        "Nenhum CSV de fluxos encontrado. Informe --fluxos ou gere com:\n"
        "  python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv"
    )


def calcular_meses_ate_2026(data) -> int:
    """Meses de `data` até 30/06/2026 (anos×12 + meses; dias ignorados)."""
    ts = pd.Timestamp(data).to_pydatetime()
    delta = relativedelta(DATA_REFERENCIA, ts)
    return delta.years * 12 + delta.months


def _coluna_impacto(columns: pd.Index) -> str | None:
    for name in COLUNAS_IMPACTO:
        if name in columns:
            return name
    return None


def _impacto_recalcular(subsidio: pd.Series, meses: pd.Series, taxa_aa: float) -> pd.Series:
    return subsidio * (1.0 + taxa_aa / 12.0) ** meses


def _impacto_composta(subsidio: pd.Series, meses: pd.Series, taxa_aa: float) -> pd.Series:
    taxa_m = taxa_mensal_composta(taxa_aa)
    return subsidio * (1.0 + taxa_m) ** meses


def agregar_impacto_por_ano(
    df: pd.DataFrame,
    *,
    modo: str = "recalcular",
    taxa_selic_anual: float = TAXA_SELIC_ANUAL,
) -> pd.DataFrame:
    """
    Agrega subsídio e impacto fiscal por ano de pagamento.

    Retorna colunas:
      Ano | Soma Subsídio Nominal (R$) | Impacto Fiscal 2026 (R$) | Quantidade de Parcelas
    """
    if "data_fluxo" not in df.columns or "subsidio" not in df.columns:
        raise ValueError("CSV precisa das colunas data_fluxo e subsidio")

    work = df.copy()
    work["data_fluxo"] = pd.to_datetime(work["data_fluxo"])
    work["ano_pagamento"] = work["data_fluxo"].dt.year
    work["subsidio"] = pd.to_numeric(work["subsidio"], errors="coerce").fillna(0.0)

    if modo == "coluna":
        col = _coluna_impacto(work.columns)
        if col is None:
            raise ValueError(
                "Modo 'coluna' exige impacto_fiscal ou impacto no CSV. "
                "Use --modo recalcular ou --modo composta."
            )
        work["impacto_individual"] = pd.to_numeric(work[col], errors="coerce").fillna(0.0)
    else:
        work["meses_ate_2026"] = work["data_fluxo"].apply(calcular_meses_ate_2026)
        if modo == "composta":
            work["impacto_individual"] = _impacto_composta(
                work["subsidio"], work["meses_ate_2026"], taxa_selic_anual
            )
        elif modo == "recalcular":
            work["impacto_individual"] = _impacto_recalcular(
                work["subsidio"], work["meses_ate_2026"], taxa_selic_anual
            )
        else:
            raise ValueError(f"Modo desconhecido: {modo}")

    count_col = "mes" if "mes" in work.columns else "data_fluxo"
    resumo = (
        work.groupby("ano_pagamento", sort=True)
        .agg(
            subsidio=("subsidio", "sum"),
            impacto_individual=("impacto_individual", "sum"),
            quantidade=(count_col, "count"),
        )
        .reset_index()
    )
    resumo.columns = [
        "Ano",
        "Soma Subsídio Nominal (R$)",
        "Impacto Fiscal 2026 (R$)",
        "Quantidade de Parcelas",
    ]
    resumo["Soma Subsídio Nominal (R$)"] = resumo["Soma Subsídio Nominal (R$)"].round(2)
    resumo["Impacto Fiscal 2026 (R$)"] = resumo["Impacto Fiscal 2026 (R$)"].round(2)
    return resumo


def carregar_fluxos(path: Path, chunksize: int = 500_000) -> pd.DataFrame:
    """Lê CSV (ou Excel) de parcelas; usa chunks se o arquivo for grande."""
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(path)

    # CSV pequeno: leitura direta; grande: chunks
    size = path.stat().st_size
    if size < 50 * 1024 * 1024:
        return pd.read_csv(path)

    parts: list[pd.DataFrame] = []
    usecols_candidates = [
        "data_fluxo",
        "subsidio",
        "mes",
        "impacto_fiscal",
        "impacto",
    ]
    # Primeira passagem: descobrir colunas disponíveis
    header = pd.read_csv(path, nrows=0)
    usecols = [c for c in usecols_candidates if c in header.columns]
    for chunk in pd.read_csv(path, usecols=usecols, chunksize=chunksize):
        parts.append(chunk)
    return pd.concat(parts, ignore_index=True)


def salvar_resumo(resumo: pd.DataFrame, out_xlsx: Path, out_csv: Path | None = None) -> tuple[Path, Path]:
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    resumo.to_excel(out_xlsx, index=False)
    if out_csv is None:
        out_csv = out_xlsx.with_suffix(".csv")
    resumo.to_csv(out_csv, index=False)
    return out_xlsx, out_csv


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--fluxos",
        type=Path,
        default=None,
        help="CSV/XLSX de parcelas (default: auto-detecta em output/).",
    )
    p.add_argument(
        "--modo",
        choices=("recalcular", "coluna", "composta"),
        default="recalcular",
        help=(
            "recalcular = (1+SELIC/12)^meses (padrão); "
            "coluna = impacto já no CSV; "
            "composta = (1+SELIC_m)^meses"
        ),
    )
    p.add_argument(
        "--taxa-selic",
        type=float,
        default=TAXA_SELIC_ANUAL,
        help=f"SELIC anual (default {TAXA_SELIC_ANUAL}).",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_DIR / "impacto_fiscal_por_ano.xlsx",
        help="Excel de saída (default: output/impacto_fiscal_por_ano.xlsx).",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        fluxos_path = resolver_fluxos(args.fluxos)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if not fluxos_path.exists():
        print(f"Arquivo não encontrado: {fluxos_path}", file=sys.stderr)
        return 1

    print("Carregando arquivo de fluxos...")
    print(f"  → {fluxos_path}")
    df = carregar_fluxos(fluxos_path)
    print(f"Total de parcelas carregadas: {len(df):,}")

    modo = args.modo
    if modo == "coluna" and _coluna_impacto(df.columns) is None:
        print(
            "Aviso: coluna de impacto ausente; caindo para --modo recalcular.",
            file=sys.stderr,
        )
        modo = "recalcular"

    resumo = agregar_impacto_por_ano(
        df, modo=modo, taxa_selic_anual=args.taxa_selic
    )

    print("\n" + "=" * 80)
    print("IMPACTO FISCAL POR ANO DE PAGAMENTO")
    print(f"(modo={modo}, SELIC={args.taxa_selic:.1%}, ref={DATA_REFERENCIA:%d/%m/%Y})")
    print("=" * 80)
    print(resumo.to_string(index=False))

    xlsx_path, csv_path = salvar_resumo(resumo, args.output)
    # Também grava na raiz do projeto se o usuário pediu o nome clássico
    root_xlsx = Path("impacto_fiscal_por_ano.xlsx")
    if args.output.resolve() != root_xlsx.resolve():
        resumo.to_excel(root_xlsx, index=False)

    print(f"\nArquivo salvo: {xlsx_path}")
    print(f"Arquivo salvo: {csv_path}")
    if root_xlsx.exists():
        print(f"Arquivo salvo: {root_xlsx}")

    print("\n" + "=" * 80)
    print("TOTAIS GERAIS")
    print("=" * 80)
    print(f"Total Subsídio Nominal: R$ {resumo['Soma Subsídio Nominal (R$)'].sum():,.2f}")
    print(f"Total Impacto Fiscal 2026: R$ {resumo['Impacto Fiscal 2026 (R$)'].sum():,.2f}")
    print(f"Total de Parcelas: {resumo['Quantidade de Parcelas'].sum():,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
