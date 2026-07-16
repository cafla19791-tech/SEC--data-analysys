"""
Gera fluxos de pagamento (SAC) e subsídio implícito (Selic vs taxa do contrato)
a partir de operações indiretas automáticas do BNDES (2009-2010).

Fonte padrão: Portal de Dados Abertos do BNDES (CSV).
Também aceita planilha Excel no formato do portal de transparência (header=5).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

# Selic de referência usada no cálculo do subsídio implícito (a.a.)
SELIC_AA = 0.145

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


def _to_float_br(series: pd.Series) -> pd.Series:
    """Converte números no formato BR (vírgula decimal) ou já numéricos."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    cleaned = (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.", "", regex=True)
        .str.replace(",", ".", regex=False)
        .replace({"": np.nan, "nan": np.nan, "None": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def download_and_filter_csv(
    url: str = BNDES_CSV_URL,
    start: str = "2009-01-01",
    end: str = "2010-12-31",
    dest: Path = FILTERED_CSV,
) -> Path:
    """Baixa o CSV completo em streaming e grava apenas o período solicitado."""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    print(f"Baixando e filtrando {start} .. {end} ...")
    print(f"URL: {url}")

    # chunksize mantém uso de memória baixo (~1.1 GB no total)
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


def load_from_open_data_csv(path: Path) -> pd.DataFrame:
    """Carrega CSV do portal de dados abertos e padroniza colunas."""
    df = pd.read_csv(path, sep=";", encoding="utf-8", dtype=str, low_memory=False)
    rename = {
        "data_da_contratacao": "data_contratacao",
        "valor_desembolsado_reais": "valor_desembolsado",
        "juros": "taxa_juros",
        "prazo_carencia_meses": "prazo_carencia",
        "prazo_amortizacao_meses": "prazo_amortizacao",
    }
    df = df.rename(columns=rename)
    return _prepare_contracts(df)


def load_from_excel_transparencia(path: Path, sheet_name: str | int = 0) -> pd.DataFrame:
    """
    Carrega planilha no formato usado pelo script original
    (header na linha 6 / header=5, nomes em português).
    """
    df = pd.read_excel(path, sheet_name=sheet_name, header=5)
    df = df.rename(
        columns={
            "Data da contratação": "data_contratacao",
            "Valor Desembolsado R$ (*)": "valor_desembolsado",
            "Juros": "taxa_juros",
            "Prazo - Carência (meses)": "prazo_carencia",
            "Prazo - Amortização (meses)": "prazo_amortizacao",
        }
    )
    return _prepare_contracts(df)


def _prepare_contracts(df: pd.DataFrame) -> pd.DataFrame:
    required = [
        "data_contratacao",
        "valor_desembolsado",
        "taxa_juros",
        "prazo_carencia",
        "prazo_amortizacao",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas ausentes: {missing}")

    out = df.copy()
    out["data_contratacao"] = pd.to_datetime(out["data_contratacao"], dayfirst=True, errors="coerce")
    out["valor_desembolsado"] = _to_float_br(out["valor_desembolsado"])
    out["taxa_juros"] = _to_float_br(out["taxa_juros"])
    out["prazo_carencia"] = _to_float_br(out["prazo_carencia"]).fillna(0)
    out["prazo_amortizacao"] = _to_float_br(out["prazo_amortizacao"])

    before = len(out)
    out = out.dropna(subset=["data_contratacao", "valor_desembolsado", "taxa_juros", "prazo_amortizacao"])
    out = out[(out["valor_desembolsado"] > 0) & (out["prazo_amortizacao"] > 0)]
    out = out.reset_index(drop=True)
    out["Contrato_ID"] = out.index

    print(f"Contratos carregados: {before:,}")
    print(f"Contratos válidos: {len(out):,}")
    return out


def gerar_fluxos(df: pd.DataFrame, selic_aa: float = SELIC_AA) -> pd.DataFrame:
    """
    Gera parcelas no sistema SAC:
      amortização constante = valor / n
      juros_parcela = saldo * taxa_mensal
      subsídio = (selic_mensal - taxa_mensal) * saldo
    """
    print("Gerando fluxos de pagamento...")
    records: list[dict] = []
    skipped = 0

    for row in df.itertuples(index=False):
        try:
            data_contr = pd.Timestamp(row.data_contratacao)
            valor = float(row.valor_desembolsado)
            taxa_juros = float(row.taxa_juros) / 100.0
            carencia = int(row.prazo_carencia)
            n = int(row.prazo_amortizacao)
            contrato_id = int(row.Contrato_ID)

            if n <= 0 or valor <= 0:
                skipped += 1
                continue

            amort_mensal = valor / n
            saldo = valor
            taxa_m = taxa_juros / 12.0
            selic_m = selic_aa / 12.0

            for p in range(1, n + 1):
                data_pag = data_contr + relativedelta(months=carencia + p)
                juros_parcela = saldo * taxa_m
                subsidio = (selic_m - taxa_m) * saldo
                records.append(
                    {
                        "Contrato_ID": contrato_id,
                        "Parcela": p,
                        "Data_Pagamento": data_pag.date(),
                        "Valor_Amortizacao": round(amort_mensal, 2),
                        "Juros_Parcela": round(juros_parcela, 2),
                        "Saldo_Devedor": round(saldo, 2),
                        "Subsídio": round(subsidio, 2),
                    }
                )
                saldo -= amort_mensal
        except Exception:
            skipped += 1
            continue

    fluxos = pd.DataFrame.from_records(records)
    print(f"Parcelas geradas: {len(fluxos):,}")
    if skipped:
        print(f"Contratos ignorados por erro: {skipped:,}")
    return fluxos


def salvar_saidas(df_fluxos: pd.DataFrame, stem: str = "fluxos_gerados_corrigido") -> dict[str, Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    paths: dict[str, Path] = {}

    parquet_path = OUTPUT_DIR / f"{stem}.parquet"
    csv_path = OUTPUT_DIR / f"{stem}.csv"
    xlsx_path = OUTPUT_DIR / f"{stem}.xlsx"
    resumo_path = OUTPUT_DIR / f"{stem}_resumo.xlsx"

    df_fluxos.to_parquet(parquet_path, index=False)
    paths["parquet"] = parquet_path
    print(f"Parquet: {parquet_path}")

    df_fluxos.to_csv(csv_path, index=False)
    paths["csv"] = csv_path
    print(f"CSV: {csv_path}")

    # Excel tem limite prático ~1M linhas; acima disso só amostra + resumo
    if len(df_fluxos) <= 1_000_000:
        df_fluxos.to_excel(xlsx_path, index=False)
        paths["xlsx"] = xlsx_path
        print(f"Excel: {xlsx_path}")
    else:
        amostra = df_fluxos.head(100_000)
        amostra_path = OUTPUT_DIR / f"{stem}_amostra_100k.xlsx"
        amostra.to_excel(amostra_path, index=False)
        paths["xlsx_amostra"] = amostra_path
        print(f"Excel amostra (100k): {amostra_path}")

    resumo = (
        df_fluxos.groupby("Contrato_ID", as_index=False)
        .agg(
            n_parcelas=("Parcela", "max"),
            primeira_parcela=("Data_Pagamento", "min"),
            ultima_parcela=("Data_Pagamento", "max"),
            total_amortizacao=("Valor_Amortizacao", "sum"),
            total_juros=("Juros_Parcela", "sum"),
            total_subsidio=("Subsídio", "sum"),
        )
    )
    resumo.to_excel(resumo_path, index=False)
    paths["resumo"] = resumo_path
    print(f"Resumo por contrato: {resumo_path}")
    return paths


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gera fluxos de pagamento BNDES (SAC + subsídio Selic).")
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="CSV filtrado ou Excel de transparência. Se omitido, usa data/ filtrado ou baixa da API.",
    )
    parser.add_argument(
        "--download",
        action="store_true",
        help="Força download/filtro do CSV aberto do BNDES (2009-2010).",
    )
    parser.add_argument("--start", default="2009-01-01", help="Data inicial (YYYY-MM-DD) para filtro no download.")
    parser.add_argument("--end", default="2010-12-31", help="Data final (YYYY-MM-DD) para filtro no download.")
    parser.add_argument(
        "--selic",
        type=float,
        default=SELIC_AA * 100,
        help="Selic anual %% usada no subsídio (default: 14.5).",
    )
    parser.add_argument(
        "--limit-contracts",
        type=int,
        default=None,
        help="Opcional: processa apenas os N primeiros contratos válidos (útil para testes).",
    )
    parser.add_argument(
        "--output-stem",
        default="fluxos_gerados_corrigido",
        help="Prefixo dos arquivos em output/.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print("Iniciando geração de fluxos...")

    input_path = args.input
    if args.download or input_path is None:
        if args.download or not FILTERED_CSV.exists():
            download_and_filter_csv(start=args.start, end=args.end, dest=FILTERED_CSV)
        input_path = FILTERED_CSV

    suffix = input_path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        # tenta aba padrão do script original; cai para a primeira aba se não existir
        try:
            df = load_from_excel_transparencia(input_path, sheet_name="operacoes_indiretas_automaticas")
        except ValueError:
            df = load_from_excel_transparencia(input_path, sheet_name=0)
    else:
        df = load_from_open_data_csv(input_path)

    if args.limit_contracts is not None:
        df = df.head(args.limit_contracts).copy()
        print(f"Limitando a {len(df):,} contratos (--limit-contracts)")

    df_fluxos = gerar_fluxos(df, selic_aa=args.selic / 100.0)
    if df_fluxos.empty:
        print("Nenhuma parcela gerada.", file=sys.stderr)
        return 1

    paths = salvar_saidas(df_fluxos, stem=args.output_stem)
    print("\nConcluído!")
    print(f"Total de parcelas: {len(df_fluxos):,}")
    for kind, path in paths.items():
        print(f"  [{kind}] {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
