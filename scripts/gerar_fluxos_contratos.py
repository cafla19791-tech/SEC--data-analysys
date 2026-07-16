#!/usr/bin/env python3
"""
Gera fluxos mensais de contratos (SAC + carência) e subsídio implícito
(Selic de referência vs taxa do contrato), a partir do script ContAgil.

Metodologia (igual ao script de referência):
  - taxa_mensal composta: (1 + juros_aa)^(1/12) - 1
  - TJLP/TLP: juros_aa efetivo = 6% + spread do contrato
  - Fluxos em TODOS os meses (carência + amortização)
  - Amortização constante só após a carência
  - subsídio = saldo × (Selic_aa/12 − taxa_mensal)

Fontes aceitas:
  - Excel ContAgil / portal de transparência (header=5, colunas em PT)
  - CSV do Portal de Dados Abertos do BNDES
  - Download automático do CSV aberto (2009–2010 por padrão)
"""

from __future__ import annotations

import argparse
import glob
import os
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

# =========================================
# CONFIGURAÇÕES
# =========================================
SELIC_AA = 0.145
TAMANHO_LOTE = 50_000
TJLP_TLP_BASE = 0.06

BNDES_CSV_URL = (
    "https://dadosabertos.bndes.gov.br/dataset/"
    "10e21ad1-568e-45e5-a8af-43f2c05ef1a2/resource/"
    "612faa0b-b6be-4b2c-9317-da5dc2c0b901/download/"
    "operacoes-financiamento-operacoes-indiretas-automaticas.csv"
)

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
FILTERED_CSV = DATA_DIR / "operacoes_indiretas_automaticas_2009-2010.csv"
OUTPUT_DIR = ROOT / "output"


# =========================================
# CARREGAR SELIC (série temporal opcional)
# =========================================
class SelicSerie:
    """Lookup rápido de fatores Selic por data (Excel STP do Bacen)."""

    def __init__(self, datas: np.ndarray, fatores: np.ndarray):
        self.datas = datas
        self.fatores = fatores

    @classmethod
    def from_excel(cls, path: Path) -> "SelicSerie":
        selic = pd.read_excel(path)
        datas = pd.to_datetime(selic.iloc[:, 0], dayfirst=True, errors="coerce").values.astype(
            "datetime64[ns]"
        )
        fatores = pd.to_numeric(selic.iloc[:, min(3, selic.shape[1] - 1)], errors="coerce").values
        mask = ~pd.isna(datas) & ~pd.isna(fatores)
        return cls(datas[mask], fatores[mask])

    def fator_rapido(self, datas) -> np.ndarray:
        datas = np.array(datas, dtype="datetime64[ns]")
        idx = np.searchsorted(self.datas, datas, side="right") - 1
        idx[idx < 0] = 0
        return self.fatores[idx]


# =========================================
# FUNÇÕES AUXILIARES
# =========================================
def limpar_valor(series: pd.Series) -> pd.Series:
    """Converte valores BR (R$ 1.234,56) ou já numéricos."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    cleaned = (
        series.astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
        .replace({"": np.nan, "nan": np.nan, "None": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def _find_col(columns, *needles: str) -> str | None:
    upper = {c: str(c).upper() for c in columns}
    for col, name in upper.items():
        if all(n.upper() in name for n in needles):
            return col
    return None


def taxa_mensal_from_row(custo: str, juros_pct) -> float:
    """
    Converte juros anuais do contrato em taxa mensal composta.
    Para TJLP/TLP, soma 6 p.p. (base) ao spread, como no script ContAgil.
    """
    try:
        juros = float(str(juros_pct).replace("%", "").replace(",", ".")) / 100.0
    except (TypeError, ValueError):
        juros = 0.0

    custo_u = str(custo or "").upper()
    if "TJLP" in custo_u or "TLP" in custo_u:
        juros_aa = TJLP_TLP_BASE + juros
    else:
        # TAXA FIXA e demais: usa o juros do contrato
        juros_aa = juros

    return (1.0 + juros_aa) ** (1.0 / 12.0) - 1.0


def download_and_filter_csv(
    url: str = BNDES_CSV_URL,
    start: str = "2009-01-01",
    end: str = "2010-12-31",
    dest: Path = FILTERED_CSV,
) -> Path:
    """Baixa o CSV aberto do BNDES em streaming e grava só o período pedido."""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    print(f"Baixando e filtrando {start} .. {end} ...")
    print(f"URL: {url}")

    reader = pd.read_csv(
        url,
        sep=";",
        encoding="cp1252",
        dtype=str,
        chunksize=100_000,
        low_memory=False,
    )

    start_ts = pd.Timestamp(start)
    end_ts = pd.Timestamp(end)
    parts: list[pd.DataFrame] = []
    total_rows = 0
    kept_rows = 0

    for i, chunk in enumerate(reader, start=1):
        total_rows += len(chunk)
        dates = pd.to_datetime(chunk["data_da_contratacao"], errors="coerce")
        mask = (dates >= start_ts) & (dates <= end_ts)
        filtered = chunk.loc[mask].copy()
        kept_rows += len(filtered)
        if not filtered.empty:
            parts.append(filtered)
        if i % 5 == 0:
            print(f"  chunks={i:,} lidas={total_rows:,} mantidas={kept_rows:,}")

    if not parts:
        raise RuntimeError("Nenhum contrato encontrado no período solicitado.")

    df = pd.concat(parts, ignore_index=True)
    dest.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(dest, index=False, sep=";", encoding="utf-8")
    print(f"Arquivo filtrado salvo: {dest} ({len(df):,} contratos)")
    return dest


def load_contracts_excel(path: Path, header: int = 5) -> pd.DataFrame:
    """Carrega Excel ContAgil / transparência e padroniza colunas."""
    try:
        df = pd.read_excel(path, sheet_name="operacoes_indiretas_automaticas", header=header)
    except ValueError:
        df = pd.read_excel(path, sheet_name=0, header=header)

    col_valor = _find_col(df.columns, "VALOR", "DESEMBOLS") or _find_col(df.columns, "VALOR")
    col_data = _find_col(df.columns, "DATA", "CONTRATA")
    col_juros = _find_col(df.columns, "JUROS")
    col_carencia = _find_col(df.columns, "CARÊNCIA") or _find_col(df.columns, "CARENCIA")
    col_amort = _find_col(df.columns, "AMORTIZA")
    col_encargo = (
        _find_col(df.columns, "ENCARGO")
        or _find_col(df.columns, "CUSTO", "FINANC")
        or _find_col(df.columns, "CUSTO")
    )

    if not col_valor or not col_data:
        raise ValueError(f"Colunas principais não encontradas em {path.name}")

    out = pd.DataFrame(
        {
            "data_contratacao": pd.to_datetime(df[col_data], dayfirst=True, errors="coerce"),
            "valor_desembolsado": limpar_valor(df[col_valor]),
            "juros": limpar_valor(df[col_juros]) if col_juros else 0.0,
            "prazo_carencia": limpar_valor(df[col_carencia]).fillna(0) if col_carencia else 0,
            "prazo_amortizacao": limpar_valor(df[col_amort]) if col_amort else np.nan,
            "custo_financeiro": df[col_encargo].astype(str) if col_encargo else "",
        }
    )
    return _filter_valid(out)


def load_contracts_csv(path: Path) -> pd.DataFrame:
    """Carrega CSV do portal de dados abertos (separador ;)."""
    df = pd.read_csv(path, sep=";", encoding="utf-8", dtype=str, low_memory=False)
    # fallback encoding se necessário
    if "data_da_contratacao" not in df.columns and "data_contratacao" not in df.columns:
        df = pd.read_csv(path, sep=";", encoding="cp1252", dtype=str, low_memory=False)

    rename = {
        "data_da_contratacao": "data_contratacao",
        "data_contratacao": "data_contratacao",
        "valor_desembolsado_reais": "valor_desembolsado",
        "valor_desembolsado": "valor_desembolsado",
        "juros": "juros",
        "prazo_carencia_meses": "prazo_carencia",
        "prazo_carencia": "prazo_carencia",
        "prazo_amortizacao_meses": "prazo_amortizacao",
        "prazo_amortizacao": "prazo_amortizacao",
        "custo_financeiro": "custo_financeiro",
    }
    present = {k: v for k, v in rename.items() if k in df.columns}
    df = df.rename(columns=present)

    out = pd.DataFrame(
        {
            "data_contratacao": pd.to_datetime(df["data_contratacao"], dayfirst=True, errors="coerce"),
            "valor_desembolsado": limpar_valor(df["valor_desembolsado"]),
            "juros": limpar_valor(df["juros"]) if "juros" in df.columns else 0.0,
            "prazo_carencia": limpar_valor(df["prazo_carencia"]).fillna(0)
            if "prazo_carencia" in df.columns
            else 0,
            "prazo_amortizacao": limpar_valor(df["prazo_amortizacao"])
            if "prazo_amortizacao" in df.columns
            else np.nan,
            "custo_financeiro": df["custo_financeiro"].astype(str)
            if "custo_financeiro" in df.columns
            else "",
        }
    )
    return _filter_valid(out)


def _filter_valid(df: pd.DataFrame) -> pd.DataFrame:
    before = len(df)
    out = df.dropna(subset=["data_contratacao", "valor_desembolsado", "prazo_amortizacao"]).copy()
    out = out[(out["valor_desembolsado"] > 0) & (out["prazo_amortizacao"] > 0)]
    out = out.reset_index(drop=True)
    out["contrato"] = out.index
    print(f"Contratos na entrada: {before:,}")
    print(f"Contratos válidos: {len(out):,}")
    return out


def gerar_fluxos(
    df: pd.DataFrame,
    pasta_saida: Path,
    selic_aa: float = SELIC_AA,
    tamanho_lote: int = TAMANHO_LOTE,
    stem: str = "fluxos",
) -> tuple[int, list[Path]]:
    """
    Gera fluxos mês a mês (carência + amortização) e grava em lotes CSV.
    Retorna (total_fluxos, caminhos_dos_lotes).
    """
    pasta_saida.mkdir(parents=True, exist_ok=True)
    # limpa lotes anteriores com o mesmo stem
    for old in pasta_saida.glob(f"{stem}_*.csv"):
        old.unlink()

    buffer: list[dict] = []
    lote = 0
    total_fluxos = 0
    caminhos: list[Path] = []
    selic_m = selic_aa / 12.0

    print(f"Gerando fluxos (Selic {selic_aa * 100:.2f}% a.a., lote={tamanho_lote:,})...")

    for row in df.itertuples(index=False):
        saldo = float(row.valor_desembolsado)
        data_inicio = pd.Timestamp(row.data_contratacao)
        if pd.isnull(data_inicio):
            continue

        taxa = taxa_mensal_from_row(getattr(row, "custo_financeiro", ""), row.juros)
        carencia = int(float(row.prazo_carencia or 0))
        amort = int(float(row.prazo_amortizacao))
        if amort <= 0:
            continue

        amortizacao = saldo / amort
        contrato_id = int(row.contrato)

        for t in range(carencia + amort):
            data_fluxo = data_inicio + relativedelta(months=t + 1)
            subsidio = saldo * (selic_m - taxa)

            buffer.append(
                {
                    "contrato": contrato_id,
                    "mes": t + 1,
                    "data_fluxo": data_fluxo.date(),
                    "saldo": round(saldo, 2),
                    "amortizacao": round(amortizacao if t >= carencia else 0.0, 2),
                    "taxa_mensal": round(taxa, 8),
                    "subsidio": round(subsidio, 2),
                    "em_carencia": t < carencia,
                }
            )
            total_fluxos += 1

            if t >= carencia:
                saldo -= amortizacao
            if saldo <= 1e-9:
                break

            if len(buffer) >= tamanho_lote:
                path = pasta_saida / f"{stem}_{lote}.csv"
                pd.DataFrame(buffer).to_csv(path, index=False)
                print(f"   Lote {lote} salvo ({len(buffer):,} linhas) → {path.name}")
                caminhos.append(path)
                buffer = []
                lote += 1

    if buffer:
        path = pasta_saida / f"{stem}_{lote}.csv"
        pd.DataFrame(buffer).to_csv(path, index=False)
        print(f"   Lote {lote} salvo ({len(buffer):,} linhas) → {path.name}")
        caminhos.append(path)

    return total_fluxos, caminhos


def consolidar_resumo(caminhos: list[Path], pasta_saida: Path, stem: str = "fluxos") -> Path | None:
    """Agrega totais por contrato a partir dos lotes CSV."""
    if not caminhos:
        return None

    parts = []
    for path in caminhos:
        chunk = pd.read_csv(path, usecols=["contrato", "subsidio", "amortizacao", "data_fluxo", "mes"])
        parts.append(chunk)

    df = pd.concat(parts, ignore_index=True)
    resumo = (
        df.groupby("contrato", as_index=False)
        .agg(
            n_meses=("mes", "max"),
            primeira_data=("data_fluxo", "min"),
            ultima_data=("data_fluxo", "max"),
            total_amortizacao=("amortizacao", "sum"),
            total_subsidio=("subsidio", "sum"),
        )
    )
    out = pasta_saida / f"{stem}_resumo.xlsx"
    resumo.to_excel(out, index=False)
    print(f"Resumo por contrato: {out} ({len(resumo):,} contratos)")
    return out


def listar_excels(pasta: Path) -> list[Path]:
    return sorted(Path(p) for p in glob.glob(os.path.join(str(pasta), "*.xlsx")))


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Gera fluxos mensais de contratos (SAC + carência + subsídio Selic)."
    )
    p.add_argument(
        "--pasta-dados",
        type=Path,
        default=None,
        help="Pasta com vários .xlsx ContAgil (equivale a pasta_dados do script original).",
    )
    p.add_argument(
        "--input",
        type=Path,
        default=None,
        help="CSV filtrado ou Excel único de contratos.",
    )
    p.add_argument(
        "--download",
        action="store_true",
        help="Baixa/filtra CSV aberto do BNDES (2009–2010 por padrão).",
    )
    p.add_argument("--start", default="2009-01-01", help="Data inicial do download (YYYY-MM-DD).")
    p.add_argument("--end", default="2010-12-31", help="Data final do download (YYYY-MM-DD).")
    p.add_argument(
        "--arquivo-selic",
        type=Path,
        default=None,
        help="Excel STP da Selic (opcional; usado para lookup de fatores).",
    )
    p.add_argument(
        "--selic",
        type=float,
        default=SELIC_AA * 100,
        help="Selic anual %% no cálculo do subsídio (default: 14.5).",
    )
    p.add_argument(
        "--pasta-saida",
        type=Path,
        default=OUTPUT_DIR,
        help="Pasta de saída dos lotes CSV (default: output/).",
    )
    p.add_argument(
        "--tamanho-lote",
        type=int,
        default=TAMANHO_LOTE,
        help="Linhas por arquivo CSV (default: 50000).",
    )
    p.add_argument(
        "--limit-contracts",
        type=int,
        default=None,
        help="Processa apenas os N primeiros contratos válidos (teste).",
    )
    p.add_argument("--output-stem", default="fluxos", help="Prefixo dos arquivos de saída.")
    p.add_argument(
        "--excel-header",
        type=int,
        default=5,
        help="Linha do header no Excel de transparência (default: 5 = linha 6).",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print("Iniciando geração de fluxos de contratos...")

    selic_serie = None
    if args.arquivo_selic is not None:
        if not args.arquivo_selic.exists():
            print(f"Aviso: arquivo Selic não encontrado: {args.arquivo_selic}", file=sys.stderr)
        else:
            selic_serie = SelicSerie.from_excel(args.arquivo_selic)
            print(f"Série Selic carregada: {len(selic_serie.datas):,} pontos")
            # Mantida para uso futuro (fator acumulado). O subsídio usa --selic, como no script original.

    frames: list[pd.DataFrame] = []

    if args.pasta_dados is not None:
        arquivos = listar_excels(args.pasta_dados)
        if not arquivos:
            print(f"Nenhum .xlsx em {args.pasta_dados}", file=sys.stderr)
            return 1
        for arquivo in arquivos:
            print(f"\nProcessando: {arquivo}")
            try:
                frames.append(load_contracts_excel(arquivo, header=args.excel_header))
            except Exception as exc:
                print(f"   Ignorado ({exc})")
                continue
    elif args.download or (args.input is None and not FILTERED_CSV.exists()):
        download_and_filter_csv(start=args.start, end=args.end, dest=FILTERED_CSV)
        frames.append(load_contracts_csv(FILTERED_CSV))
    else:
        input_path = args.input or FILTERED_CSV
        suffix = input_path.suffix.lower()
        print(f"Entrada: {input_path}")
        if suffix in {".xlsx", ".xls"}:
            frames.append(load_contracts_excel(input_path, header=args.excel_header))
        else:
            frames.append(load_contracts_csv(input_path))

    if not frames:
        print("Nenhum contrato carregado.", file=sys.stderr)
        return 1

    df = pd.concat(frames, ignore_index=True)
    df["contrato"] = df.index

    if args.limit_contracts is not None:
        df = df.head(args.limit_contracts).copy()
        df["contrato"] = df.index
        print(f"Limitando a {len(df):,} contratos (--limit-contracts)")

    total, caminhos = gerar_fluxos(
        df,
        pasta_saida=args.pasta_saida,
        selic_aa=args.selic / 100.0,
        tamanho_lote=args.tamanho_lote,
        stem=args.output_stem,
    )

    if total == 0:
        print("Nenhum fluxo gerado.", file=sys.stderr)
        return 1

    resumo = consolidar_resumo(caminhos, args.pasta_saida, stem=args.output_stem)

    print(f"\nFINALIZADO! Total de fluxos gerados: {total:,}")
    print(f"Lotes CSV: {len(caminhos)}")
    if resumo:
        print(f"Resumo: {resumo}")
    if selic_serie is not None:
        print("(Série Selic carregada; subsídio usou a taxa --selic, como no script original.)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
