#!/usr/bin/env python3
"""
Processamento ultra-rápido ContAgil com Polars.

Porta o script ContAgil/WinPython (lazy CSV + join_asof SELIC) com as
correções de metodologia do repo:

  - fatores SELIC na **coluna D** (índice 3), com fallback na coluna E
  - fator final = FATOR_30_06_2026 (STP ContAgil) ou max até 30/06/2026
  - join na própria ``data_fluxo`` (nearest ContAgil), não data+1
  - aliases de colunas (taxa_selic_mensal / selic_mes, impacto / impacto_fiscal)
  - fallbacks cloud quando os caminhos WinPython não existem

Uso (WinPython ContAgil):
  python scripts/resumo_fluxos_polars.py \\
      --pasta "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
      --original "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx" \\
      --selic "STP-20260716182715078.xlsx"

Uso (repo / cloud):
  python3 scripts/resumo_fluxos_polars.py \\
      --pasta output \\
      --original data/sample_operacoes_com_agente.csv \\
      --selic data/selic_fatores_bacen.xlsx
"""

from __future__ import annotations

import argparse
import glob
import os
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import polars as pl
from xlsxwriter import Workbook

from scripts.gerar_fluxos import FATOR_30_06_2026
from scripts.resumo_fluxos_avancado import (
    listar_arquivos_fluxos,
    resolver_original,
    resolver_pasta,
    resolver_selic,
)

DATA_FINAL_DEFAULT = date(2026, 6, 30)
AGENTE_COL = "Instituição Financeira Credenciada"
WORKBOOK_NAME = "resumo_fluxos_polars.xlsx"


def _normalizar_coluna_data(df: pl.DataFrame, col: str = "data") -> pl.DataFrame:
    """Normaliza coluna de data (Date / Datetime / string BR)."""
    dtype = df[col].dtype
    if dtype == pl.Date:
        return df
    if dtype == pl.Datetime:
        return df.with_columns(pl.col(col).cast(pl.Date))
    return df.with_columns(
        pl.col(col).cast(pl.Utf8).str.to_date("%d/%m/%Y", strict=False).alias(col)
    )


def _escolher_coluna_fator(raw: pl.DataFrame) -> tuple[str, bool]:
    """Escolhe coluna de fator ContAgil.

    Ordem (igual ``SelicSerie.from_dataframe``):
      1) coluna nomeada fator_acumulado / fator → Bacen/cache (sem FATOR ref)
      2) coluna D (índice 3) se parece fator acumulado (mediana > 1) → STP ContAgil
      3) coluna E (índice 4) ou última numérica
    """
    cols = list(raw.columns)
    lower = {str(c).strip().lower(): c for c in cols}
    for key in ("fator_acumulado", "fator"):
        if key in lower:
            return lower[key], False

    def _mediana(col: str) -> float:
        s = raw.get_column(col).cast(pl.Float64, strict=False).drop_nulls()
        if s.is_empty():
            return 0.0
        return float(s.median())

    if len(cols) >= 4:
        med_d = _mediana(cols[3])
        # Taxa diária Bacen (~0.04) ≠ fator acumulado ContAgil (>1 e crescente)
        if med_d > 1.0:
            return cols[3], True
    if len(cols) >= 5:
        return cols[4], False
    return cols[-1], False


def carregar_fatores_selic(caminho_selic: str | Path) -> tuple[pl.DataFrame, float]:
    """Lê STP ContAgil / Bacen: data + fator acumulado.

    Retorna (selic_df, fator_final). STP ContAgil (col D) usa FATOR_30_06_2026;
    Bacen/outros usam o fator em 30/06/2026 (ou max da série).
    """
    path = Path(caminho_selic)
    raw = pl.read_excel(path)
    if raw.width < 2:
        raise ValueError(f"SELIC sem colunas suficientes: {path}")

    data_col = raw.columns[0]
    fator_col, usar_fator_ref = _escolher_coluna_fator(raw)

    selic = raw.select(
        [
            pl.col(data_col).alias("data"),
            pl.col(fator_col).cast(pl.Float64, strict=False).alias("fator_acumulado"),
        ]
    )
    selic = (
        _normalizar_coluna_data(selic, "data")
        .drop_nulls(["data", "fator_acumulado"])
        .filter(pl.col("fator_acumulado") > 0)
        .sort("data")
        .unique(subset=["data"], keep="last")
    )

    if selic.is_empty():
        raise ValueError(f"Nenhum fator SELIC válido em {path}")

    if usar_fator_ref:
        fator_final = float(FATOR_30_06_2026)
    else:
        ate_fim = selic.filter(pl.col("data") <= DATA_FINAL_DEFAULT)
        fator_final = float(
            (ate_fim if not ate_fim.is_empty() else selic)["fator_acumulado"].max()
        )
    return selic, fator_final


def calcular_impacto_fiscal(
    df: pl.DataFrame,
    selic_df: pl.DataFrame,
    fator_final: float,
) -> pl.DataFrame:
    """impacto = subsidio × fator_final / fator(nearest data_fluxo) — ContAgil."""
    work = df.with_columns(
        [
            pl.col("data_fluxo").cast(pl.Date).alias("data_fluxo"),
            pl.col("subsidio").cast(pl.Float64),
        ]
    ).sort("data_fluxo")

    # ContAgil: nearest na própria data da parcela (join_asof backward ≈ último ≤ data)
    joined = work.join_asof(
        selic_df.sort("data"),
        left_on="data_fluxo",
        right_on="data",
        strategy="backward",
    )

    return joined.with_columns(
        [
            pl.when(pl.col("fator_acumulado").is_not_null() & (pl.col("fator_acumulado") > 0))
            .then(
                (pl.col("subsidio") * (fator_final / pl.col("fator_acumulado"))).round(2)
            )
            .otherwise(pl.col("subsidio").round(2))
            .alias("impacto_acumulado_2026")
        ]
    ).drop([c for c in ("data",) if c in joined.columns and c != "data_fluxo"])


def _normalizar_colunas_fluxos(df: pl.DataFrame) -> pl.DataFrame:
    """Aliases ContAgil/repo → nomes canônicos do script Polars."""
    rename: dict[str, str] = {}
    cols = set(df.columns)

    if "impacto" not in cols and "impacto_fiscal" in cols:
        rename["impacto_fiscal"] = "impacto"
    if "taxa_selic_mensal" in cols and "selic_mes" not in cols:
        rename["taxa_selic_mensal"] = "selic_mes"
    if "taxa_contrato_mensal" in cols and "taxa_contrato" not in cols:
        rename["taxa_contrato_mensal"] = "taxa_contrato"
    if "Instituição Financeira" in cols and AGENTE_COL not in cols:
        rename["Instituição Financeira"] = AGENTE_COL

    if rename:
        df = df.rename(rename)

    required = {"contrato", "data_fluxo", "subsidio"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Colunas ausentes nos fluxos: {sorted(missing)}")

    # data_fluxo → Date
    if df["data_fluxo"].dtype != pl.Date:
        if df["data_fluxo"].dtype == pl.Datetime:
            df = df.with_columns(pl.col("data_fluxo").cast(pl.Date))
        else:
            df = df.with_columns(
                pl.col("data_fluxo").cast(pl.Utf8).str.to_date(strict=False)
            )

    df = df.with_columns(pl.col("subsidio").cast(pl.Float64, strict=False).fill_null(0.0))

    if "saldo" not in df.columns:
        if "saldo_fiscal" in df.columns:
            df = df.with_columns(pl.col("saldo_fiscal").alias("saldo"))
        else:
            df = df.with_columns(pl.lit(None).cast(pl.Float64).alias("saldo"))

    return df.drop_nulls(["data_fluxo"])


def _somente_amostra_resumo(path: Path) -> bool:
    """True se o Excel é workbook agregado (só aba Amostra_Parcelas)."""
    if path.suffix.lower() not in {".xlsx", ".xls"}:
        return False
    try:
        import fastexcel

        nomes = set(fastexcel.read_excel(path).sheet_names)
    except Exception:
        return False
    return "Amostra_Parcelas" in nomes and "Sheet1" not in nomes and "Parcelas" not in nomes


def carregar_fluxos_polars(pasta: Path) -> pl.DataFrame:
    """Lazy-load CSV ContAgil; fallback Excel quando não há CSV."""
    csvs = sorted(glob.glob(os.path.join(str(pasta), "fluxos_*.csv")))
    csvs = [p for p in csvs if "diario" not in Path(p).stem.lower()]

    if csvs:
        print(f"Carregando {len(csvs)} CSV(s) com Polars (lazy)...")
        lazy_frames = [pl.scan_csv(arq, try_parse_dates=True) for arq in csvs]
        df = pl.concat(lazy_frames, how="diagonal_relaxed").collect()
        return _normalizar_colunas_fluxos(df)

    # Fallback: Excel de parcelas (repo / ContAgil xlsx)
    xlsxs = [
        p
        for p in listar_arquivos_fluxos(pasta)
        if p.suffix.lower() in {".xlsx", ".xls"}
    ]
    # Ignora workbooks só com Amostra_Parcelas quando há detalhe completo
    if any(not _somente_amostra_resumo(p) for p in xlsxs):
        xlsxs = [p for p in xlsxs if not _somente_amostra_resumo(p)]

    if not xlsxs:
        raise FileNotFoundError(
            f"Nenhum fluxos_*.csv/.xlsx em {pasta}. "
            "Gere com scripts/contagil_fluxos.py ou scripts/gerar_fluxos.py."
        )

    print(f"Carregando {len(xlsxs)} Excel(s) com Polars...")
    partes: list[pl.DataFrame] = []
    for path in xlsxs:
        try:
            try:
                part = pl.read_excel(path, sheet_name="Sheet1")
            except Exception:
                part = pl.read_excel(path)
            if "data_fluxo" not in part.columns or "subsidio" not in part.columns:
                try:
                    part = pl.read_excel(path, sheet_name="Amostra_Parcelas")
                except Exception:
                    print(f"  Ignorando {path.name}: sem colunas de parcelas")
                    continue
            print(f"  Lendo: {path.name} ({part.height:,} linhas)")
            partes.append(_normalizar_colunas_fluxos(part))
        except Exception as exc:
            print(f"  Ignorando {path.name}: {exc}")
    if not partes:
        raise FileNotFoundError(f"Nenhum arquivo de parcelas válido em {pasta}")
    return pl.concat(partes, how="diagonal_relaxed")


def carregar_instituicoes(caminho_original: str | Path) -> pl.DataFrame | None:
    """Lê Excel/CSV original e devolve contrato → Instituição Financeira Credenciada."""
    path = Path(caminho_original)
    suffix = path.suffix.lower()

    if suffix == ".csv":
        # Amostra / CSV BNDES (sep=;)
        try:
            raw = pl.read_csv(path, separator=";", infer_schema_length=5000)
        except Exception:
            raw = pl.read_csv(path, infer_schema_length=5000)
        col_agente = None
        for cand in (
            "instituicao_financeira_credenciada",
            "Instituição Financeira Credenciada",
            "agente",
        ):
            if cand in raw.columns:
                col_agente = cand
                break
        if col_agente is None:
            return None
        return (
            raw.select(pl.col(col_agente).alias(AGENTE_COL))
            .with_row_index("contrato")
        )

    # Excel portal (header na linha 6) ou ContAgil (header=0)
    for skip in (5, 0):
        try:
            raw = pl.read_excel(
                path,
                sheet_name=0,
                read_options={"header_row": skip},
            )
        except TypeError:
            # API fastexcel/polars variante
            try:
                raw = pl.read_excel(path, sheet_name=0)
                if skip == 5 and AGENTE_COL not in raw.columns:
                    continue
            except Exception:
                continue
        except Exception:
            continue

        if AGENTE_COL in raw.columns:
            return (
                raw.select(pl.col(AGENTE_COL))
                .drop_nulls()
                .with_row_index("contrato")
            )
        # aliases
        for cand in (
            "Instituicao Financeira Credenciada",
            "agente",
            "Instituição Financeira",
        ):
            if cand in raw.columns:
                return (
                    raw.select(pl.col(cand).alias(AGENTE_COL))
                    .drop_nulls()
                    .with_row_index("contrato")
                )
    return None


def adicionar_spread(df: pl.DataFrame) -> pl.DataFrame:
    """spread = (1 + selic_mes − taxa_contrato) ** mes_no_contrato (se colunas existirem)."""
    if "selic_mes" not in df.columns or "taxa_contrato" not in df.columns:
        return df.with_columns(pl.lit(None).cast(pl.Float64).alias("spread"))

    return (
        df.sort(["contrato", "data_fluxo"])
        .with_columns(
            [
                pl.col("contrato").cum_count().over("contrato").alias("mes_no_contrato"),
                pl.col("selic_mes").cast(pl.Float64, strict=False),
                pl.col("taxa_contrato").cast(pl.Float64, strict=False),
            ]
        )
        .with_columns(
            [
                (
                    (1 + pl.col("selic_mes") - pl.col("taxa_contrato"))
                    ** pl.col("mes_no_contrato")
                )
                .cast(pl.Float64)
                .alias("spread")
            ]
        )
    )


def montar_resumos(df: pl.DataFrame) -> dict[str, pl.DataFrame]:
    """Gera Contratos / Por_Ano / Por_Agente / Impacto_Por_Ano."""
    agente = AGENTE_COL if AGENTE_COL in df.columns else None
    group_contrato = ["contrato"] + ([agente] if agente else [])

    aggs = [
        pl.sum("subsidio").round(2).alias("subsidio_total"),
        pl.sum("impacto_acumulado_2026").round(2).alias("impacto_2026"),
        pl.col("saldo").last().alias("saldo_final"),
        pl.len().alias("parcelas"),
    ]
    if "spread" in df.columns:
        aggs.insert(-1, pl.col("spread").last().alias("spread_final"))
    else:
        aggs.insert(-1, pl.lit(None).cast(pl.Float64).alias("spread_final"))

    resumo_contratos = df.group_by(group_contrato).agg(aggs).sort("contrato")

    group_ano = ["contrato", "ano"] + ([agente] if agente else [])
    resumo_ano = (
        df.group_by(group_ano)
        .agg(
            [
                pl.sum("subsidio").round(2).alias("subsidio"),
                pl.sum("impacto_acumulado_2026").round(2).alias("impacto_acumulado_2026"),
            ]
        )
        .sort(["contrato", "ano"])
    )

    if agente:
        por_agente = (
            df.group_by(agente)
            .agg(
                [
                    pl.col("contrato").n_unique().alias("qtd_contratos"),
                    pl.sum("subsidio").round(2).alias("subsidio"),
                    pl.sum("impacto_acumulado_2026").round(2).alias("impacto_acumulado_2026"),
                ]
            )
            .sort("subsidio", descending=True)
        )
    else:
        por_agente = pl.DataFrame(
            {
                AGENTE_COL: [],
                "qtd_contratos": [],
                "subsidio": [],
                "impacto_acumulado_2026": [],
            }
        )

    impacto_ano = (
        df.group_by("ano")
        .agg(
            [
                pl.sum("subsidio").round(2).alias("subsidio"),
                pl.sum("impacto_acumulado_2026").round(2).alias("impacto_acumulado_2026"),
                pl.len().alias("parcelas"),
            ]
        )
        .sort("ano")
    )

    return {
        "Contratos": resumo_contratos,
        "Por_Ano": resumo_ano,
        "Por_Agente": por_agente,
        "Impacto_Por_Ano": impacto_ano,
    }


def exportar_excel(resumos: dict[str, pl.DataFrame], output_path: Path) -> Path:
    """Grava workbook multi-aba via xlsxwriter."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with Workbook(str(output_path)) as wb:
        for nome, frame in resumos.items():
            frame.write_excel(workbook=wb, worksheet=nome, autofit=True)
    return output_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--pasta", required=True, type=Path, help="Pasta ContAgil saida/")
    p.add_argument(
        "--original",
        required=True,
        help="Excel/CSV de operações (nome curto ContAgil ok)",
    )
    p.add_argument(
        "--selic",
        required=True,
        help="Excel STP ContAgil (col D) — nome curto ok",
    )
    p.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Pasta de saída (default: --pasta resolvida)",
    )
    p.add_argument(
        "--baixar-selic",
        action="store_true",
        help="Força Bacen se STP ausente (já é automático para nomes STP-*)",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        pasta = resolver_pasta(args.pasta)
        original_path = resolver_original(args.original, pasta)
        selic_path, _serie = resolver_selic(
            args.selic, pasta, baixar_selic=args.baixar_selic
        )
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if selic_path is None:
        print("Arquivo SELIC não encontrado.", file=sys.stderr)
        return 1

    output_dir = args.output_dir if args.output_dir is not None else pasta
    os.makedirs(output_dir, exist_ok=True)

    print("Processamento Ultra-Rápido ContAgil (Polars)")
    print(f"Pasta fluxos : {pasta}")
    print(f"Original     : {original_path}")
    print(f"SELIC        : {selic_path}")

    try:
        df = carregar_fluxos_polars(pasta)
        print(f"Total de linhas: {df.height:,}")

        inst = carregar_instituicoes(original_path)
        if inst is not None and AGENTE_COL not in df.columns:
            df = df.join(inst, on="contrato", how="left")
            print(f"Instituições  : {inst.height:,} contratos mapeados")
        elif AGENTE_COL in df.columns:
            print("Instituições  : já presentes nos fluxos")
        else:
            print("Instituições  : (não encontradas no original)")

        selic_df, fator_final = carregar_fatores_selic(selic_path)
        print(
            f"SELIC pontos  : {selic_df.height:,} | fator_final={fator_final:.5f}"
        )

        df = calcular_impacto_fiscal(df, selic_df, fator_final)
        df = adicionar_spread(df)
        df = df.with_columns(pl.col("data_fluxo").dt.year().alias("ano"))

        resumos = montar_resumos(df)
        out = exportar_excel(resumos, Path(output_dir) / WORKBOOK_NAME)

        # Espelhos CSV leves
        resumos["Contratos"].write_csv(Path(output_dir) / "resumo_contratos_polars.csv")
        resumos["Por_Ano"].write_csv(Path(output_dir) / "resumo_por_ano_polars.csv")
        resumos["Por_Agente"].write_csv(Path(output_dir) / "resumo_por_agente_polars.csv")
        resumos["Impacto_Por_Ano"].write_csv(
            Path(output_dir) / "impacto_fiscal_por_ano_polars.csv"
        )
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print("✅ Processamento Polars concluído com sucesso!")
    print(f"   Workbook        : {out}")
    print(f"   Total contratos : {df['contrato'].n_unique():,}")
    print(f"   Total linhas    : {df.height:,}")
    print(
        f"   Subsídio total  : R$ {df['subsidio'].sum():,.2f}"
    )
    print(
        f"   Impacto 2026    : R$ {df['impacto_acumulado_2026'].sum():,.2f}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
