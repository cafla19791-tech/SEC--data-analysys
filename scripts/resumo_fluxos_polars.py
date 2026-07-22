#!/usr/bin/env python3
"""
Versão FINAL ContAgil — Polars + gráficos + relatório executivo.

Porta o script ContAgil/WinPython (lazy CSV + join_asof SELIC) com:

  - fatores SELIC na **coluna D** (STP) **ou** taxas ContAgil/Bacen (% a.d.)
  - TJLP mensal ContAgil (``--tjlp tjlp_mensal.xlsx``) na taxa efetiva
  - fator final = FATOR_30_06_2026 (STP) ou cumprod das taxas até 30/06/2026
  - join_asof na própria ``data_fluxo`` (ContAgil nearest) — não produto global
  - Excel multi-aba + Plotly/Matplotlib + RELATORIO_EXECUTIVO.md
  - fallbacks cloud quando os caminhos WinPython não existem

Uso (WinPython ContAgil):
  python scripts/resumo_fluxos_polars.py \\
      --pasta "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
      --original "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx" \\
      --selic "STP-20260716182715078.xlsx" \\
      --tjlp "tjlp_mensal.xlsx"

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
import warnings
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import plotly.express as px
import polars as pl
from xlsxwriter import Workbook

from scripts.gerar_fluxos import (
    CONTAGIL_WINPYTHON,
    DATA_DIR,
    FATOR_30_06_2026,
    TJLP_TLP_BASE,
)
from scripts.resumo_fluxos_avancado import (
    _candidatos_nome,
    _parece_contagil,
    listar_arquivos_fluxos,
    resolver_original,
    resolver_pasta,
    resolver_selic,
)

DATA_FINAL_DEFAULT = date(2026, 6, 30)
AGENTE_COL = "Instituição Financeira Credenciada"
WORKBOOK_NAME = "resumo_fluxos_polars_final.xlsx"
WORKBOOK_ALIAS = "resumo_fluxos_polars.xlsx"
RELATORIO_NAME = "RELATORIO_EXECUTIVO.md"
GRAFICO_HTML = "grafico_interativo.html"
GRAFICO_PNG = "grafico_top_subsidio.png"
CONTAGIL_TJLP_DEFAULT = CONTAGIL_WINPYTHON / "tjlp_mensal.xlsx"
NOME_TJLP_DEFAULT = "tjlp_mensal.xlsx"


def _normalizar_coluna_data(df: pl.DataFrame, col: str = "data") -> pl.DataFrame:
    """Normaliza coluna de data (Date / Datetime / string BR dd/mm/yyyy ou dd/mm/yy)."""
    dtype = df[col].dtype
    if dtype == pl.Date:
        return df
    if dtype == pl.Datetime:
        return df.with_columns(pl.col(col).cast(pl.Date))
    texto = pl.col(col).cast(pl.Utf8).str.strip_chars()
    # ContAgil Bacen: "01/01/09" (%d/%m/%y) ou "01/01/2009" (%d/%m/%Y)
    parsed = (
        pl.when(texto.str.len_chars() <= 8)
        .then(texto.str.to_date("%d/%m/%y", strict=False))
        .otherwise(texto.str.to_date("%d/%m/%Y", strict=False))
    )
    return df.with_columns(parsed.alias(col))


def _encontrar_coluna(cols: list[str], *needles: str) -> str | None:
    """Primeira coluna cujo nome contém todos os needles (case-insensitive)."""
    for c in cols:
        low = str(c).strip().lower()
        if all(n.lower() in low for n in needles):
            return c
    return None


def _coluna_data_contagil(cols: list[str]) -> str:
    for cand in ("Data", "data", "DATA"):
        if cand in cols:
            return cand
    return cols[0]


def _para_decimal_taxa(
    series: pl.Series,
    *,
    coluna: str | None = None,
    forcar_percentual: bool | None = None,
) -> pl.Series:
    """Converte taxa em % para decimal.

    ContAgil Bacen usa colunas ``% a.d.`` / ``% a.m.`` com valores como 0,04
    (percentual) → 0,0004 (decimal). Sem indício no nome, usa mediana > 0,3.
    """
    s = series.cast(pl.Float64, strict=False)
    if forcar_percentual is None and coluna is not None:
        low = str(coluna).lower()
        forcar_percentual = ("% a.d" in low) or ("% a.m" in low) or ("%" in low)
    if forcar_percentual is None:
        med = float(s.drop_nulls().median()) if s.drop_nulls().len() else 0.0
        forcar_percentual = med > 0.3
    return s / 100.0 if forcar_percentual else s


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


def _parece_selic_taxas(raw: pl.DataFrame) -> bool:
    """True se o Excel ContAgil tem taxa Selic (% a.d. / % a.m.), não fator col D."""
    cols = [str(c) for c in raw.columns]
    if _encontrar_coluna(cols, "selic") and (
        _encontrar_coluna(cols, "% a.d")
        or _encontrar_coluna(cols, "% a.m")
        or _encontrar_coluna(cols, "a.d")
        or _encontrar_coluna(cols, "a.m")
    ):
        return True
    # fator_acumulado nomeado → fatores
    lower = {str(c).strip().lower() for c in cols}
    if "fator_acumulado" in lower or "fator" in lower:
        return False
    return False


def carregar_selic_mensal(caminho: str | Path) -> pl.DataFrame:
    """Lê Excel ContAgil de taxas Selic (Bacen) — colunas ``Data`` + Selic % a.d./a.m.

    Retorna ``data``, ``selic_taxa`` (decimal) e ``fator_acumulado`` (cumprod).
    """
    path = Path(caminho)
    raw = pl.read_excel(path)
    print(f"Colunas SELIC ({path.name}): {raw.columns}")
    cols = [str(c) for c in raw.columns]
    data_col = _coluna_data_contagil(cols)
    taxa_col = (
        _encontrar_coluna(cols, "selic", "% a.d")
        or _encontrar_coluna(cols, "selic", "a.d")
        or _encontrar_coluna(cols, "selic", "% a.m")
        or _encontrar_coluna(cols, "selic", "a.m")
        or _encontrar_coluna(cols, "selic")
    )
    if taxa_col is None:
        raise ValueError(
            f"Coluna de taxa Selic não encontrada em {path}. Colunas: {cols}"
        )

    taxa = _para_decimal_taxa(raw.get_column(taxa_col), coluna=taxa_col)
    selic = raw.select(
        [
            pl.col(data_col).alias("data"),
            pl.Series("selic_taxa", taxa),
        ]
    )
    selic = (
        _normalizar_coluna_data(selic, "data")
        .drop_nulls(["data", "selic_taxa"])
        .filter(pl.col("selic_taxa") > -0.5)
        .sort("data")
        .unique(subset=["data"], keep="last")
        .with_columns(
            [
                (1.0 + pl.col("selic_taxa"))
                .cum_prod()
                .alias("fator_acumulado")
            ]
        )
    )
    if selic.is_empty():
        raise ValueError(f"Nenhuma taxa SELIC válida em {path}")
    return selic


def carregar_tjlp_mensal(caminho: str | Path) -> pl.DataFrame:
    """Lê Excel ContAgil TJLP mensal (% a.m.).

    Retorna ``data``, ``tjlp_mensal`` (decimal ao mês).
    """
    path = Path(caminho)
    raw = pl.read_excel(path)
    print(f"Colunas TJLP ({path.name}): {raw.columns}")
    cols = [str(c) for c in raw.columns]
    data_col = _coluna_data_contagil(cols)
    taxa_col = (
        _encontrar_coluna(cols, "tjlp", "% a.m")
        or _encontrar_coluna(cols, "tjlp", "a.m")
        or _encontrar_coluna(cols, "tjlp")
    )
    if taxa_col is None:
        raise ValueError(
            f"Coluna TJLP mensal não encontrada em {path}. Colunas: {cols}"
        )

    taxa = _para_decimal_taxa(raw.get_column(taxa_col), coluna=taxa_col)
    tjlp = raw.select(
        [
            pl.col(data_col).alias("data"),
            pl.Series("tjlp_mensal", taxa),
        ]
    )
    tjlp = (
        _normalizar_coluna_data(tjlp, "data")
        .drop_nulls(["data", "tjlp_mensal"])
        .sort("data")
        .unique(subset=["data"], keep="last")
    )
    if tjlp.is_empty():
        raise ValueError(f"Nenhuma TJLP válida em {path}")
    return tjlp


def carregar_selic_auto(caminho: str | Path) -> tuple[pl.DataFrame, float, str]:
    """Detecta layout ContAgil e carrega SELIC.

    Retorna (selic_df com fator_acumulado, fator_final, modo)
    onde modo ∈ {"fatores", "taxas"}.
    """
    path = Path(caminho)
    raw = pl.read_excel(path)
    if _parece_selic_taxas(raw):
        selic = carregar_selic_mensal(path)
        ate_fim = selic.filter(pl.col("data") <= DATA_FINAL_DEFAULT)
        fator_final = float(
            (ate_fim if not ate_fim.is_empty() else selic)["fator_acumulado"].max()
        )
        return selic, fator_final, "taxas"
    selic, fator_final = carregar_fatores_selic(path)
    return selic, fator_final, "fatores"


def resolver_tjlp(nome: str | Path | None, pasta: Path) -> Path | None:
    """Resolve ``tjlp_mensal.xlsx`` ContAgil (opcional)."""
    if nome is None:
        nome = CONTAGIL_TJLP_DEFAULT
    path = Path(nome)
    bases = [
        Path.cwd(),
        pasta,
        pasta.parent,
        CONTAGIL_WINPYTHON,
        DATA_DIR / "contagil_winpython",
        DATA_DIR,
        ROOT / "attachments",
        Path("/home/workdir/attachments"),
    ]
    for cand in _candidatos_nome(path, bases):
        if cand.exists() and cand.is_file():
            return cand
    if CONTAGIL_TJLP_DEFAULT.exists():
        return CONTAGIL_TJLP_DEFAULT
    sample = DATA_DIR / NOME_TJLP_DEFAULT
    if sample.exists():
        return sample
    # Ausente: usa TJLP_TLP_BASE (6%) na taxa efetiva — não é erro
    if _parece_contagil(path) or path.name.lower().startswith("tjlp"):
        print(f"⚠️ TJLP ContAgil ausente: {nome}")
        print(f"   Usando base ContAgil TJLP/TLP = {TJLP_TLP_BASE:.2%} a.a.")
        return None
    return None


def anexar_tjlp_mensal(df: pl.DataFrame, tjlp: pl.DataFrame | None) -> pl.DataFrame:
    """Join_asof TJLP mensal na data da parcela (mês ContAgil)."""
    if tjlp is None or tjlp.is_empty():
        return df.with_columns(
            pl.lit(None).cast(pl.Float64).alias("tjlp_mensal")
        )

    work = df.with_columns(
        [
            pl.col("data_fluxo").cast(pl.Date),
            pl.col("data_fluxo").cast(pl.Date).dt.truncate("1mo").alias("_mes_tjlp"),
        ]
    ).sort("_mes_tjlp")
    tjlp_m = (
        tjlp.with_columns(pl.col("data").dt.truncate("1mo").alias("_mes_tjlp"))
        .sort("_mes_tjlp")
        .unique(subset=["_mes_tjlp"], keep="last")
    )
    joined = work.join_asof(
        tjlp_m.select(["_mes_tjlp", "tjlp_mensal"]),
        on="_mes_tjlp",
        strategy="backward",
    ).drop("_mes_tjlp")
    return joined


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


def adicionar_taxa_contrato_efetiva(df: pl.DataFrame) -> pl.DataFrame:
    """Define ``taxa_contrato_efetiva`` (mensal) ContAgil TJLP/TLP.

    Preferência:
      1) série ``tjlp_mensal`` ContAgil + spread do contrato
         → ``(1+tjlp)×(1+spread)−1`` (ou TJLP_TLP_BASE se série ausente)
      2) ``taxa_contrato`` / ``taxa_contrato_mensal`` já nos fluxos (já efetiva)
      3) recomputa a partir de ``encargo_financeiro`` + ``juros`` (% a.a.)
    """
    work = df
    if "encargo_financeiro" not in work.columns and "custo_financeiro" in work.columns:
        work = work.with_columns(pl.col("custo_financeiro").alias("encargo_financeiro"))
    if "taxa_contrato" not in work.columns and "taxa_contrato_mensal" in work.columns:
        work = work.with_columns(pl.col("taxa_contrato_mensal").alias("taxa_contrato"))

    tem_tjlp_serie = (
        "tjlp_mensal" in work.columns
        and work["tjlp_mensal"].null_count() < work.height
    )

    # Spread do contrato (mensal). Se só há juros % a.a., converte.
    if "taxa_contrato" in work.columns:
        spread_m = pl.col("taxa_contrato").cast(pl.Float64, strict=False).fill_null(0.0)
    elif "juros" in work.columns:
        juros_aa = pl.col("juros").cast(pl.Float64, strict=False).fill_null(0.0) / 100.0
        spread_m = (1.0 + juros_aa) ** (1.0 / 12.0) - 1.0
    else:
        spread_m = pl.lit(0.0)

    base_fix_m = (1.0 + TJLP_TLP_BASE) ** (1.0 / 12.0) - 1.0
    if tem_tjlp_serie:
        tjlp_m = pl.col("tjlp_mensal").cast(pl.Float64, strict=False).fill_null(base_fix_m)
    else:
        tjlp_m = pl.lit(base_fix_m)

    # ContAgil: TJLP/TLP → (1+tjlp_m)×(1+spread_m)−1; demais → spread_m
    if "encargo_financeiro" in work.columns:
        enc = (
            pl.col("encargo_financeiro")
            .cast(pl.Utf8)
            .fill_null("")
            .str.to_uppercase()
        )
        efetiva = (
            pl.when(enc.str.contains("TJLP") | enc.str.contains("TLP"))
            .then((1.0 + tjlp_m) * (1.0 + spread_m) - 1.0)
            .otherwise(spread_m)
        )
        return work.with_columns(efetiva.alias("taxa_contrato_efetiva"))

    # Sem encargo: se já há taxa_contrato nos fluxos gerados, ela já é efetiva
    if "taxa_contrato" in work.columns and not tem_tjlp_serie:
        return work.with_columns(spread_m.alias("taxa_contrato_efetiva"))

    return work.with_columns(spread_m.alias("taxa_contrato_efetiva"))


def adicionar_spread(df: pl.DataFrame) -> pl.DataFrame:
    """spread = (1 + selic_mes − taxa_contrato_efetiva) ** mes_no_contrato."""
    work = adicionar_taxa_contrato_efetiva(df)

    if "selic_mes" not in work.columns:
        if "taxa_selic_mensal" in work.columns:
            work = work.with_columns(pl.col("taxa_selic_mensal").alias("selic_mes"))
        else:
            return work.with_columns(pl.lit(None).cast(pl.Float64).alias("spread"))

    if "taxa_contrato_efetiva" not in work.columns:
        return work.with_columns(pl.lit(None).cast(pl.Float64).alias("spread"))

    return (
        work.sort(["contrato", "data_fluxo"])
        .with_columns(
            [
                pl.col("contrato").cum_count().over("contrato").alias("mes_no_contrato"),
                pl.col("selic_mes").cast(pl.Float64, strict=False),
                pl.col("taxa_contrato_efetiva").cast(pl.Float64, strict=False),
            ]
        )
        .with_columns(
            [
                (
                    (1 + pl.col("selic_mes") - pl.col("taxa_contrato_efetiva"))
                    ** pl.col("mes_no_contrato")
                )
                .cast(pl.Float64)
                .alias("spread")
            ]
        )
    )


def montar_totais(df: pl.DataFrame) -> pl.DataFrame:
    """Aba Totais_Gerais do workbook final."""
    n_agentes = (
        int(df[AGENTE_COL].n_unique()) if AGENTE_COL in df.columns else 0
    )
    return pl.DataFrame(
        {
            "Métrica": [
                "Total Subsídio",
                "Total Impacto 2026",
                "Total Contratos",
                "Total Agentes",
                "Total Parcelas",
            ],
            "Valor": [
                round(float(df["subsidio"].sum()), 2),
                round(float(df["impacto_acumulado_2026"].sum()), 2),
                int(df["contrato"].n_unique()),
                n_agentes,
                int(df.height),
            ],
        }
    )


def montar_resumos(df: pl.DataFrame) -> dict[str, pl.DataFrame]:
    """Gera Contratos / Por_Ano / Por_Agente / Impacto_Por_Ano / Totais_Gerais."""
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

    # Versão FINAL: Por_Ano = ano × instituição (paste ContAgil)
    if agente:
        resumo_ano = (
            df.group_by(["ano", agente])
            .agg(
                [
                    pl.sum("subsidio").round(2).alias("subsidio"),
                    pl.sum("impacto_acumulado_2026")
                    .round(2)
                    .alias("impacto_acumulado_2026"),
                ]
            )
            .sort(["ano", agente])
        )
        por_agente = (
            df.group_by(agente)
            .agg(
                [
                    pl.col("contrato").n_unique().alias("qtd_contratos"),
                    pl.sum("subsidio").round(2).alias("subsidio"),
                    pl.sum("impacto_acumulado_2026")
                    .round(2)
                    .alias("impacto_acumulado_2026"),
                ]
            )
            .sort("subsidio", descending=True)
        )
    else:
        resumo_ano = (
            df.group_by("ano")
            .agg(
                [
                    pl.sum("subsidio").round(2).alias("subsidio"),
                    pl.sum("impacto_acumulado_2026")
                    .round(2)
                    .alias("impacto_acumulado_2026"),
                ]
            )
            .sort("ano")
        )
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
        "Totais_Gerais": montar_totais(df),
    }


def gerar_graficos(df: pl.DataFrame, output_dir: Path) -> tuple[Path, Path]:
    """Gera PNG (Matplotlib top-10) + HTML interativo (Plotly impacto/ano)."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    png_path = output_dir / GRAFICO_PNG
    html_path = output_dir / GRAFICO_HTML

    top = (
        df.group_by("contrato")
        .agg(pl.sum("subsidio").alias("subsidio"))
        .sort("subsidio", descending=True)
        .head(10)
        .to_pandas()
    )
    plt.figure(figsize=(12, 6))
    plt.bar(top["contrato"].astype(str), top["subsidio"], color="#1f4e79")
    plt.title("Top 10 Contratos — Subsídio Nominal")
    plt.xlabel("Contrato")
    plt.ylabel("Subsídio (R$)")
    plt.tight_layout()
    plt.savefig(png_path, dpi=120)
    plt.close()

    por_ano = (
        df.group_by("ano")
        .agg(pl.sum("impacto_acumulado_2026").alias("impacto_acumulado_2026"))
        .sort("ano")
        .to_pandas()
    )
    fig = px.bar(
        por_ano,
        x="ano",
        y="impacto_acumulado_2026",
        title="Impacto Fiscal acumulado até 30/06/2026 por Ano",
        labels={
            "ano": "Ano",
            "impacto_acumulado_2026": "Impacto Fiscal 2026 (R$)",
        },
    )
    fig.write_html(str(html_path), include_plotlyjs="cdn")
    return png_path, html_path


def gerar_relatorio_executivo(df: pl.DataFrame, output_dir: Path) -> Path:
    """Escreve RELATORIO_EXECUTIVO.md com totais e top agentes."""
    output_dir = Path(output_dir)
    path = output_dir / RELATORIO_NAME

    n_contratos = int(df["contrato"].n_unique())
    total_sub = float(df["subsidio"].sum())
    total_imp = float(df["impacto_acumulado_2026"].sum())

    if AGENTE_COL in df.columns:
        top_agentes = (
            df.group_by(AGENTE_COL)
            .agg(
                [
                    pl.col("contrato").n_unique().alias("Contratos"),
                    pl.sum("subsidio").round(2).alias("Subsídio (R$)"),
                    pl.sum("impacto_acumulado_2026")
                    .round(2)
                    .alias("Impacto 2026 (R$)"),
                ]
            )
            .sort("Impacto 2026 (R$)", descending=True)
            .head(5)
            .to_pandas()
        )
        try:
            tabela = top_agentes.to_markdown(index=False)
        except Exception:
            tabela = top_agentes.to_string(index=False)
    else:
        tabela = "_Instituição financeira não disponível nos fluxos._"

    relatorio = f"""# Relatório Executivo - Subsídios BNDES

**Data:** {datetime.now().strftime("%d/%m/%Y %H:%M")}
**Total de Contratos:** {n_contratos:,}
**Total de Parcelas:** {df.height:,}
**Total Subsídio Nominal:** R$ {total_sub:,.2f}
**Total Impacto Fiscal 2026:** R$ {total_imp:,.2f}

## Principais Agentes

{tabela}

## Arquivos gerados

- `{WORKBOOK_NAME}` — Contratos, Por_Ano, Por_Agente, Impacto_Por_Ano, Totais_Gerais
- `{GRAFICO_HTML}` — gráfico interativo (Plotly)
- `{GRAFICO_PNG}` — top 10 contratos (Matplotlib)
- `{RELATORIO_NAME}` — este relatório

Metodologia ContAgil: impacto = subsídio × fator_final / fator(data_fluxo),
com fator na coluna D do STP (FATOR_30_06_2026 = {FATOR_30_06_2026}) ou
série Bacen equivalente.
"""
    path.write_text(relatorio, encoding="utf-8")
    return path


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
        help="Excel STP ContAgil (col D) ou taxas Selic ContAgil/Bacen",
    )
    p.add_argument(
        "--tjlp",
        default=str(CONTAGIL_TJLP_DEFAULT),
        help=(
            "Excel TJLP mensal ContAgil "
            f"(default: {CONTAGIL_TJLP_DEFAULT})"
        ),
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
    p.add_argument(
        "--sem-graficos",
        action="store_true",
        help="Não gera PNG/HTML (útil em testes headless sem plotly).",
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
        tjlp_path = resolver_tjlp(args.tjlp, pasta)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if selic_path is None:
        print("Arquivo SELIC não encontrado.", file=sys.stderr)
        return 1

    output_dir = Path(args.output_dir if args.output_dir is not None else pasta)
    os.makedirs(output_dir, exist_ok=True)

    print("Versão FINAL ContAgil — Polars + SELIC/TJLP + gráficos + relatório")
    print(f"Pasta fluxos : {pasta}")
    print(f"Original     : {original_path}")
    print(f"SELIC        : {selic_path}")
    print(f"TJLP         : {tjlp_path if tjlp_path else f'(base {TJLP_TLP_BASE:.2%} a.a.)'}")

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

        selic_df, fator_final, modo_selic = carregar_selic_auto(selic_path)
        print(
            f"SELIC pontos  : {selic_df.height:,} | modo={modo_selic} | "
            f"fator_final={fator_final:.5f}"
        )

        tjlp_df = carregar_tjlp_mensal(tjlp_path) if tjlp_path is not None else None
        if tjlp_df is not None:
            print(f"TJLP pontos   : {tjlp_df.height:,}")
        df = anexar_tjlp_mensal(df, tjlp_df)

        df = calcular_impacto_fiscal(df, selic_df, fator_final)
        df = adicionar_spread(df)
        df = df.with_columns(pl.col("data_fluxo").dt.year().alias("ano"))

        resumos = montar_resumos(df)
        out = exportar_excel(resumos, output_dir / WORKBOOK_NAME)
        # Alias de compatibilidade com a versão anterior
        exportar_excel(resumos, output_dir / WORKBOOK_ALIAS)

        resumos["Contratos"].write_csv(output_dir / "resumo_contratos_polars.csv")
        resumos["Por_Ano"].write_csv(output_dir / "resumo_por_ano_polars.csv")
        resumos["Por_Agente"].write_csv(output_dir / "resumo_por_agente_polars.csv")
        resumos["Impacto_Por_Ano"].write_csv(
            output_dir / "impacto_fiscal_por_ano_polars.csv"
        )
        resumos["Totais_Gerais"].write_csv(output_dir / "totais_gerais_polars.csv")

        relatorio = gerar_relatorio_executivo(df, output_dir)
        png_path = html_path = None
        if not args.sem_graficos:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                png_path, html_path = gerar_graficos(df, output_dir)
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print("✅ Versão FINAL com Polars gerada com sucesso!")
    print(f"   Workbook        : {out}")
    print(f"   Relatório       : {relatorio}")
    if png_path is not None:
        print(f"   Gráfico PNG     : {png_path}")
        print(f"   Gráfico HTML    : {html_path}")
    print(f"   Total contratos : {df['contrato'].n_unique():,}")
    print(f"   Total linhas    : {df.height:,}")
    print(f"   Subsídio total  : R$ {df['subsidio'].sum():,.2f}")
    print(f"   Impacto 2026    : R$ {df['impacto_acumulado_2026'].sum():,.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
