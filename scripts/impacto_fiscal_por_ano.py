#!/usr/bin/env python3
"""
Impacto fiscal por ano de pagamento — capitalizado até 30/06/2026.

Lê o CSV de parcelas (`fluxos_completos_*.csv`), calcula o impacto individual
de cada parcela e agrega por ano de `data_fluxo`.

Metodologia ContAgil (modo padrão quando há STP/Bacen):
  data_proxima = data_parcela + 1 dia
  impacto = subsídio × fator(nearest 30/06/2026) / fator(nearest data_proxima)
  Fatores: coluna E do STP ContAgil, ou Bacen SGS 11 (--baixar-selic).

Outros modos:
  - coluna: usa `impacto_fiscal` / `impacto` já gravado no CSV
  - recalcular: subsídio × (1 + SELIC_aa/12)^meses
  - composta: subsídio × (1 + SELIC_m)^meses

Uso:
  python3 scripts/impacto_fiscal_por_ano.py --baixar-selic
  python3 scripts/impacto_fiscal_por_ano.py --arquivo-selic "STP-....xlsx"
  python3 scripts/impacto_fiscal_por_ano.py --modo coluna --fluxos output/fluxos_completos_final.csv
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

from scripts.gerar_fluxos import (
    DATA_IMPACTO,
    OUTPUT_DIR,
    SELIC_BACEN_CACHE,
    TAXA_SELIC_ANUAL,
    SelicSerie,
    resolver_arquivo_selic,
    taxa_mensal_composta,
)

DATA_REFERENCIA = DATA_IMPACTO  # 30/06/2026

CANDIDATOS_FLUXOS = (
    OUTPUT_DIR / "fluxos_completos_final.csv",
    OUTPUT_DIR / "fluxos_completos_corrigido.csv",
    OUTPUT_DIR / "fluxos_amostra.csv",
    Path("/tmp/app-streamlit/output/fluxos_completos_corrigido.csv"),
    Path("/tmp/app/output/fluxos_amostra.csv"),
)

COLUNAS_IMPACTO = ("impacto_fiscal", "impacto")
MODOS = ("contagil", "coluna", "recalcular", "composta")


def resolver_fluxos(explicit: Path | None = None) -> Path:
    if explicit is not None:
        return explicit
    for path in CANDIDATOS_FLUXOS:
        if path.exists() and path.is_file():
            return path
    raise FileNotFoundError(
        "Nenhum CSV de fluxos encontrado. Informe --fluxos ou gere com:\n"
        "  python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --baixar-selic"
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


def _impacto_contagil(
    subsidio: pd.Series,
    data_fluxo: pd.Series,
    selic_serie: SelicSerie,
) -> pd.Series:
    """ContAgil: capitaliza a partir do dia seguinte à parcela (coluna E / Bacen)."""
    datas_proxima = pd.to_datetime(data_fluxo) + timedelta(days=1)
    idx_inicio = np.fromiter(
        (selic_serie.idx_proximo(d) for d in datas_proxima),
        dtype=np.int64,
        count=len(datas_proxima),
    )
    idx_fim = selic_serie.idx_proximo(DATA_REFERENCIA)
    fator_inicio = selic_serie.fatores[idx_inicio]
    fator_fim = float(selic_serie.fatores[idx_fim])
    subs = subsidio.to_numpy(dtype=float, copy=False)

    out = np.where(subs <= 0, 0.0, subs)
    mask = (subs > 0) & (idx_fim > idx_inicio) & (fator_inicio > 0)
    out = out.astype(float, copy=True)
    out[mask] = np.round(subs[mask] * (fator_fim / fator_inicio[mask]), 2)
    out[~mask & (subs > 0)] = np.round(subs[~mask & (subs > 0)], 2)
    return pd.Series(out, index=subsidio.index)


def carregar_serie_selic(
    arquivo_selic: Path | None = None,
    baixar_selic: bool = False,
) -> SelicSerie | None:
    path = resolver_arquivo_selic(arquivo_selic)
    if path is not None:
        print(f"Lendo SELIC ContAgil (col E / fator_acumulado): {path}")
        serie = SelicSerie.from_excel(path)
        print(f"  {len(serie.datas):,} pontos ({serie.origem})")
        return serie
    if arquivo_selic is not None:
        raise FileNotFoundError(f"Arquivo SELIC não encontrado: {arquivo_selic}")
    if baixar_selic:
        serie = SelicSerie.from_bacen(cache_path=SELIC_BACEN_CACHE)
        print(f"  {len(serie.datas):,} pontos ({serie.origem})")
        return serie
    return None


def agregar_impacto_por_ano(
    df: pd.DataFrame,
    *,
    modo: str = "contagil",
    taxa_selic_anual: float = TAXA_SELIC_ANUAL,
    selic_serie: SelicSerie | None = None,
) -> pd.DataFrame:
    """
    Agrega subsídio e impacto fiscal por ano de pagamento.

    Retorna colunas:
      Ano | Soma Subsídio Nominal (R$) | Impacto Fiscal 2026 (R$) | Quantidade de Parcelas
    """
    if modo not in MODOS:
        raise ValueError(f"Modo desconhecido: {modo}")
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
                "Use --modo contagil, recalcular ou composta."
            )
        work["impacto_individual"] = pd.to_numeric(work[col], errors="coerce").fillna(0.0)
    elif modo == "contagil":
        if selic_serie is None:
            raise ValueError(
                "Modo 'contagil' exige série SELIC. "
                "Passe --arquivo-selic ou --baixar-selic."
            )
        work["impacto_individual"] = _impacto_contagil(
            work["subsidio"], work["data_fluxo"], selic_serie
        )
    else:
        work["meses_ate_2026"] = work["data_fluxo"].apply(calcular_meses_ate_2026)
        if modo == "composta":
            work["impacto_individual"] = _impacto_composta(
                work["subsidio"], work["meses_ate_2026"], taxa_selic_anual
            )
        else:  # recalcular
            work["impacto_individual"] = _impacto_recalcular(
                work["subsidio"], work["meses_ate_2026"], taxa_selic_anual
            )

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
    header = pd.read_csv(path, nrows=0)
    usecols = [c for c in usecols_candidates if c in header.columns]
    for chunk in pd.read_csv(path, usecols=usecols, chunksize=chunksize):
        parts.append(chunk)
    return pd.concat(parts, ignore_index=True)


def salvar_resumo(
    resumo: pd.DataFrame, out_xlsx: Path, out_csv: Path | None = None
) -> tuple[Path, Path]:
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
        choices=MODOS,
        default="contagil",
        help=(
            "contagil = fatores STP/Bacen col E, +1 dia (padrão); "
            "coluna = impacto já no CSV; "
            "recalcular = (1+SELIC/12)^meses; "
            "composta = (1+SELIC_m)^meses"
        ),
    )
    p.add_argument(
        "--arquivo-selic",
        type=Path,
        default=None,
        help=(
            "Excel STP ContAgil (col A=data, col E=fator). "
            "Default: auto-descoberta (caminho ContAgil Windows / data/STP*.xlsx)."
        ),
    )
    p.add_argument(
        "--baixar-selic",
        action="store_true",
        help="Baixa SELIC Bacen (SGS 11) e monta fatores ContAgil se não houver STP.",
    )
    p.add_argument(
        "--taxa-selic",
        type=float,
        default=TAXA_SELIC_ANUAL,
        help=f"SELIC anual nos modos recalcular/composta (default {TAXA_SELIC_ANUAL}).",
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

    print("Gerando impacto fiscal por ano (metodologia ContAgil quando disponível)...")
    print(f"Carregando arquivo de fluxos: {fluxos_path}")
    df = carregar_fluxos(fluxos_path)
    print(f"Total de parcelas carregadas: {len(df):,}")

    modo = args.modo
    selic_serie: SelicSerie | None = None

    if modo == "coluna" and _coluna_impacto(df.columns) is None:
        print(
            "Aviso: coluna de impacto ausente; tentando --modo contagil.",
            file=sys.stderr,
        )
        modo = "contagil"

    if modo == "contagil":
        try:
            selic_serie = carregar_serie_selic(args.arquivo_selic, args.baixar_selic)
        except FileNotFoundError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        if selic_serie is None:
            # Auto: se já há impacto no CSV, usa coluna; senão recalcular
            if _coluna_impacto(df.columns) is not None:
                print(
                    "Aviso: sem STP/Bacen; usando coluna de impacto do CSV.",
                    file=sys.stderr,
                )
                modo = "coluna"
            else:
                print(
                    "Aviso: sem STP ContAgil. Use --arquivo-selic ou --baixar-selic. "
                    "Caindo para --modo recalcular (SELIC 14,5%/12).",
                    file=sys.stderr,
                )
                modo = "recalcular"

    resumo = agregar_impacto_por_ano(
        df,
        modo=modo,
        taxa_selic_anual=args.taxa_selic,
        selic_serie=selic_serie,
    )

    origem = selic_serie.origem if selic_serie is not None else "n/a"
    print("\n" + "=" * 80)
    print("IMPACTO FISCAL POR ANO DE PAGAMENTO")
    print(
        f"(modo={modo}, SELIC={args.taxa_selic:.1%}, "
        f"ref={DATA_REFERENCIA:%d/%m/%Y}, fatores={origem})"
    )
    print("=" * 80)
    print(resumo.to_string(index=False))

    xlsx_path, csv_path = salvar_resumo(resumo, args.output)
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
