#!/usr/bin/env python3
"""
Gera fluxos mensais completos (carência + amortização SAC) e impacto fiscal
a valor de 30/06/2026, a partir de operações indiretas automáticas do BNDES.

Baseado no script de referência (SELIC 14,5% a.a.), com carência corrigida:
  - Fluxos em TODOS os meses (carência + amortização)
  - Amortização constante só após a carência
  - subsídio = saldo × (SELIC/12 − juros/12)
  - impacto = subsídio × (1 + SELIC/12)^(meses até 30/06/2026)

Entrada:
  - Excel do portal (header=5), ou
  - CSV aberto do BNDES (download automático 2009–2010)

Saídas:
  - output/fluxos_completos_corrigido.csv     (detalhe por parcela)
  - output/fluxos_completos_corrigido.xlsx    (resumo + amostra)
  - output/fluxos_completos_corrigido.parquet (opcional, se pyarrow ok)
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

# ===================== CONFIGURAÇÕES =====================
TAXA_SELIC_ANUAL = 0.145  # 14,5% a.a.
DATA_IMPACTO = datetime(2026, 6, 30)
SELIC_MENSAL = TAXA_SELIC_ANUAL / 12.0

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

EXCEL_COLUMNS = {
    "Data da contratação": "data_contratacao",
    "Valor Desembolsado R$ (*)": "valor_desembolsado",
    "Juros": "juros",
    "Prazo - Carência (meses)": "prazo_carencia",
    "Prazo - Amortização (meses)": "prazo_amortizacao",
}

CSV_COLUMNS = {
    "data_da_contratacao": "data_contratacao",
    "valor_desembolsado_reais": "valor_desembolsado",
    "juros": "juros",
    "prazo_carencia_meses": "prazo_carencia",
    "prazo_amortizacao_meses": "prazo_amortizacao",
}


def limpar_valor(series: pd.Series) -> pd.Series:
    """Converte BR (1.234,56), US (1234.56) ou já numérico."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    s = (
        series.astype(str)
        .str.replace("R$", "", regex=False)
        .str.strip()
        .replace({"": np.nan, "nan": np.nan, "None": np.nan})
    )

    def _one(v: str):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return np.nan
        text = str(v).strip()
        if not text or text.lower() in {"nan", "none"}:
            return np.nan
        if "," in text and "." in text:
            text = text.replace(".", "").replace(",", ".")
        elif "," in text:
            text = text.replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return np.nan

    return pd.to_numeric(s.map(_one), errors="coerce")


def parse_datas(series: pd.Series) -> pd.Series:
    """Parseia ISO (YYYY-MM-DD) ou BR (DD/MM/YYYY) sem misturar."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    s = series.astype(str).str.strip()
    iso_mask = s.str.match(r"^\d{4}-\d{2}-\d{2}", na=False).fillna(False)

    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if iso_mask.any():
        out.loc[iso_mask] = pd.to_datetime(
            s[iso_mask], errors="coerce", format="ISO8601"
        ).values
    if (~iso_mask).any():
        out.loc[~iso_mask] = pd.to_datetime(
            s[~iso_mask], dayfirst=True, errors="coerce"
        ).values
    return out


def meses_ate_impacto(data_fluxo: datetime, data_impacto: datetime = DATA_IMPACTO) -> int:
    """Meses de data_fluxo até data_impacto (pode ser negativo se fluxo for futuro)."""
    return (data_impacto.year - data_fluxo.year) * 12 + (data_impacto.month - data_fluxo.month)


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


def load_from_excel(path: Path, sheet_name: str | int = "operacoes_indiretas_automaticas") -> pd.DataFrame:
    """Carrega planilha do portal de transparência (header=5)."""
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=5)
    except ValueError:
        df = pd.read_excel(path, sheet_name=0, header=5)
    rename = {k: v for k, v in EXCEL_COLUMNS.items() if k in df.columns}
    df = df.rename(columns=rename)
    return _prepare_contracts(df)


def load_from_csv(path: Path) -> pd.DataFrame:
    """Carrega CSV do portal de dados abertos."""
    df = pd.read_csv(path, sep=";", encoding="utf-8", dtype=str, low_memory=False)
    rename = {k: v for k, v in CSV_COLUMNS.items() if k in df.columns}
    df = df.rename(columns=rename)
    return _prepare_contracts(df)


def _prepare_contracts(df: pd.DataFrame) -> pd.DataFrame:
    required = [
        "data_contratacao",
        "valor_desembolsado",
        "juros",
        "prazo_carencia",
        "prazo_amortizacao",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas ausentes: {missing}. Disponíveis: {list(df.columns)}")

    out = pd.DataFrame(
        {
            "data_contratacao": parse_datas(df["data_contratacao"]),
            "valor_desembolsado": limpar_valor(df["valor_desembolsado"]),
            "juros": limpar_valor(df["juros"]),
            "prazo_carencia": limpar_valor(df["prazo_carencia"]).fillna(0),
            "prazo_amortizacao": limpar_valor(df["prazo_amortizacao"]),
        }
    )

    before = len(out)
    out = out.dropna(
        subset=["data_contratacao", "valor_desembolsado", "juros", "prazo_amortizacao"]
    )
    out = out[(out["valor_desembolsado"] > 0) & (out["prazo_amortizacao"] > 0)]
    out = out.reset_index(drop=True)
    out["contrato"] = out.index

    print(f"Contratos na entrada: {before:,}")
    print(f"Contratos válidos: {len(out):,}")
    return out


def gerar_fluxos_contrato(
    data_contr: pd.Timestamp,
    valor: float,
    taxa_juros_aa: float,
    carencia: int,
    n: int,
    contrato_id: int,
    selic_aa: float = TAXA_SELIC_ANUAL,
    data_impacto: datetime = DATA_IMPACTO,
) -> list[dict]:
    """
    Gera fluxos de UM contrato (carência + amortização).

    Correção vs script com bug:
      O original fazia `data = contr + (carencia+p)` E `em_carencia = p <= carencia`
      no loop `p=1..n`, o que zera amortização nas primeiras parcelas pós-carência
      e deixa saldo residual. Aqui o cronograma cobre carência+n meses.
    """
    if n <= 0 or valor <= 0:
        return []

    amort_mensal = valor / n
    saldo = valor
    taxa_mensal = taxa_juros_aa / 12.0
    selic_mensal = selic_aa / 12.0
    fluxos: list[dict] = []

    total_meses = carencia + n
    for p in range(1, total_meses + 1):
        data_fluxo = data_contr + relativedelta(months=p)
        em_carencia = p <= carencia
        amort = 0.0 if em_carencia else amort_mensal
        subsidio = saldo * (selic_mensal - taxa_mensal)
        meses = meses_ate_impacto(data_fluxo.to_pydatetime(), data_impacto)
        impacto = subsidio * ((1.0 + selic_mensal) ** meses)

        fluxos.append(
            {
                "contrato": contrato_id,
                "mes": p,
                "data_fluxo": data_fluxo.date(),
                "saldo": round(saldo, 2),
                "amortizacao": round(amort, 2),
                "taxa_mensal": round(taxa_mensal, 8),
                "subsidio": round(subsidio, 2),
                "em_carencia": em_carencia,
                "impacto": round(impacto, 2),
            }
        )

        if not em_carencia:
            saldo -= amort_mensal
        if saldo <= 1e-9:
            break

    return fluxos


def gerar_fluxos(
    df: pd.DataFrame,
    selic_aa: float = TAXA_SELIC_ANUAL,
    data_impacto: datetime = DATA_IMPACTO,
) -> pd.DataFrame:
    """Gera DataFrame completo de fluxos (adequado para volumes menores / testes)."""
    records: list[dict] = []
    skipped = 0

    for row in df.itertuples(index=False):
        try:
            data_contr = pd.Timestamp(row.data_contratacao)
            if pd.isna(data_contr):
                skipped += 1
                continue
            records.extend(
                gerar_fluxos_contrato(
                    data_contr=data_contr,
                    valor=float(row.valor_desembolsado),
                    taxa_juros_aa=float(row.juros) / 100.0,
                    carencia=int(float(row.prazo_carencia or 0)),
                    n=int(float(row.prazo_amortizacao)),
                    contrato_id=int(row.contrato),
                    selic_aa=selic_aa,
                    data_impacto=data_impacto,
                )
            )
        except (TypeError, ValueError, OverflowError):
            skipped += 1
            continue

    if skipped:
        print(f"Contratos ignorados por erro: {skipped:,}")
    return pd.DataFrame(records)


def processar_em_lotes(
    df: pd.DataFrame,
    csv_path: Path,
    lote: int = 2000,
    selic_aa: float = TAXA_SELIC_ANUAL,
    data_impacto: datetime = DATA_IMPACTO,
) -> dict:
    """Processa em lotes, grava CSV detalhado e acumula estatísticas."""
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    if csv_path.exists():
        csv_path.unlink()

    monthly: dict[str, float] = {}
    total_impacto = 0.0
    total_subsidio = 0.0
    total_amort = 0.0
    n_parcelas = 0
    n_contratos_ok = 0
    n_em_carencia = 0
    wrote_header = False

    n = len(df)
    print(f"Processando {n:,} contratos (lote={lote:,})...")

    for start in range(0, n, lote):
        chunk = df.iloc[start : start + lote]
        records: list[dict] = []

        for row in chunk.itertuples(index=False):
            try:
                data_contr = pd.Timestamp(row.data_contratacao)
                if pd.isna(data_contr):
                    continue
                fluxos = gerar_fluxos_contrato(
                    data_contr=data_contr,
                    valor=float(row.valor_desembolsado),
                    taxa_juros_aa=float(row.juros) / 100.0,
                    carencia=int(float(row.prazo_carencia or 0)),
                    n=int(float(row.prazo_amortizacao)),
                    contrato_id=int(row.contrato),
                    selic_aa=selic_aa,
                    data_impacto=data_impacto,
                )
                if fluxos:
                    n_contratos_ok += 1
                    records.extend(fluxos)
            except (TypeError, ValueError, OverflowError):
                continue

        if not records:
            print(f"  lote {start:,}-{start + len(chunk):,}: 0 fluxos")
            continue

        fluxos_df = pd.DataFrame(records)
        n_parcelas += len(fluxos_df)
        total_impacto += float(fluxos_df["impacto"].sum())
        total_subsidio += float(fluxos_df["subsidio"].sum())
        total_amort += float(fluxos_df["amortizacao"].sum())
        n_em_carencia += int(fluxos_df["em_carencia"].sum())

        keys = pd.to_datetime(fluxos_df["data_fluxo"]).dt.to_period("M").astype(str)
        for k, v in fluxos_df.groupby(keys, sort=False)["impacto"].sum().items():
            monthly[k] = monthly.get(k, 0.0) + float(v)

        fluxos_df.to_csv(csv_path, mode="a", index=False, header=not wrote_header)
        wrote_header = True
        print(
            f"  lote {start:,}-{start + len(chunk):,} → +{len(fluxos_df):,} "
            f"(acum {n_parcelas:,})"
        )

    return {
        "n_contratos_entrada": n,
        "n_contratos_ok": n_contratos_ok,
        "n_parcelas": n_parcelas,
        "n_parcelas_em_carencia": n_em_carencia,
        "total_amortizacao": round(total_amort, 2),
        "total_subsidio": round(total_subsidio, 2),
        "total_impacto_fiscal_2026": round(total_impacto, 2),
        "monthly": monthly,
    }


def salvar_excel_resumo(
    stats: dict,
    xlsx_path: Path,
    sample_csv: Path | None = None,
    sample_rows: int = 50_000,
) -> None:
    """Excel legível: resumo + impacto mensal + amostra de parcelas."""
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    resumo = pd.DataFrame(
        [
            {"Indicador": "Taxa SELIC anual (config)", "Valor": TAXA_SELIC_ANUAL},
            {"Indicador": "Data de impacto", "Valor": DATA_IMPACTO.date().isoformat()},
            {"Indicador": "Contratos na entrada", "Valor": stats["n_contratos_entrada"]},
            {"Indicador": "Contratos processados", "Valor": stats["n_contratos_ok"]},
            {"Indicador": "Parcelas geradas", "Valor": stats["n_parcelas"]},
            {
                "Indicador": "Parcelas em carência",
                "Valor": stats["n_parcelas_em_carencia"],
            },
            {"Indicador": "Soma Amortização", "Valor": stats["total_amortizacao"]},
            {"Indicador": "Soma Subsídio (nominal)", "Valor": stats["total_subsidio"]},
            {
                "Indicador": "Soma Impacto Fiscal 2026",
                "Valor": stats["total_impacto_fiscal_2026"],
            },
            {
                "Indicador": "Arquivo detalhado",
                "Valor": "output/fluxos_completos_corrigido.csv",
            },
        ]
    )

    mensal = (
        pd.DataFrame(
            [
                {"Ano_Mes": k, "Impacto_Fiscal_2026": round(v, 2)}
                for k, v in sorted(stats["monthly"].items())
            ]
        )
        if stats["monthly"]
        else pd.DataFrame(columns=["Ano_Mes", "Impacto_Fiscal_2026"])
    )

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        mensal.to_excel(writer, sheet_name="Impacto_Mensal", index=False)
        if sample_csv and sample_csv.exists():
            sample = pd.read_csv(sample_csv, nrows=sample_rows)
            sample.to_excel(writer, sheet_name="Amostra_Parcelas", index=False)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--excel",
        type=Path,
        help="Excel local (header=5), ex.: operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx",
    )
    p.add_argument(
        "--input",
        type=Path,
        help="CSV filtrado já baixado (sep=';).",
    )
    p.add_argument(
        "--download",
        action="store_true",
        help="Baixa/filtra CSV aberto do BNDES (2009–2010 por padrão).",
    )
    p.add_argument("--start", default="2009-01-01")
    p.add_argument("--end", default="2010-12-31")
    p.add_argument("--lote", type=int, default=2000)
    p.add_argument(
        "--max-contratos",
        type=int,
        default=None,
        help="Limita contratos (útil para testes).",
    )
    p.add_argument(
        "--stem",
        default="fluxos_completos_corrigido",
        help="Prefixo dos arquivos de saída.",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if args.excel:
        print(f"Lendo Excel: {args.excel}")
        df = load_from_excel(args.excel)
    elif args.input:
        print(f"Lendo CSV: {args.input}")
        df = load_from_csv(args.input)
    elif FILTERED_CSV.exists() and not args.download:
        print(f"Lendo cache: {FILTERED_CSV}")
        df = load_from_csv(FILTERED_CSV)
    else:
        path = download_and_filter_csv(start=args.start, end=args.end)
        df = load_from_csv(path)

    if args.max_contratos is not None:
        df = df.head(args.max_contratos).copy()
        df["contrato"] = df.index
        print(f"Limitado a {len(df):,} contratos (--max-contratos)")

    csv_path = OUTPUT_DIR / f"{args.stem}.csv"
    xlsx_path = OUTPUT_DIR / f"{args.stem}.xlsx"
    stats_path = OUTPUT_DIR / f"{args.stem}_stats.json"

    stats = processar_em_lotes(df, csv_path, lote=args.lote)
    print(json.dumps({k: v for k, v in stats.items() if k != "monthly"}, indent=2))

    salvar_excel_resumo(stats, xlsx_path, sample_csv=csv_path)
    with stats_path.open("w", encoding="utf-8") as f:
        json.dump({k: v for k, v in stats.items() if k != "monthly"}, f, indent=2)

    print(f"✅ CSV detalhado: {csv_path}")
    print(f"✅ Excel resumo:  {xlsx_path}")
    print(f"✅ Stats JSON:    {stats_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
