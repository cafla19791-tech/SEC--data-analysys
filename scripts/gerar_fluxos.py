#!/usr/bin/env python3
"""Gera fluxos de amortização e impacto fiscal (a valor de 30/06/2026)
a partir das operações indiretas automáticas do BNDES (2009–2010).

Saídas:
  - output/fluxos_gerados.csv          (detalhe por parcela; ~milhões de linhas)
  - output/fluxos_gerados.xlsx         (resumo mensal + totais; legível no Excel)
  - data/operacoes_2009_2010.parquet   (cache das operações baixadas)
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from dateutil.relativedelta import relativedelta

# Configurações
TAXA_SELIC_ANUAL = 0.145
DATA_IMPACTO = datetime(2026, 6, 30)
RESOURCE_ID = "612faa0b-b6be-4b2c-9317-da5dc2c0b901"
DATASTORE_URL = "https://dadosabertos.bndes.gov.br/api/3/action/datastore_search"

# Intervalo 2009–2010 no datastore ordenado por data_da_contratacao
# (obtido por busca binária; pode ser recalculado com --discover-offsets)
DEFAULT_OFFSET_START = 581801
DEFAULT_OFFSET_END = 931388  # exclusivo

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"

COLUMN_MAP = {
    "data_da_contratacao": "Data da contratação",
    "valor_desembolsado_reais": "Valor Desembolsado R$ (*)",
    "juros": "Juros",
    "prazo_carencia_meses": "Prazo - Carência (meses)",
    "prazo_amortizacao_meses": "Prazo - Amortização (meses)",
}

NEEDED_FIELDS = list(COLUMN_MAP.keys()) + ["_id", "cliente", "uf", "produto", "instrumento_financeiro"]


def discover_offsets(session: requests.Session) -> tuple[int, int]:
    """Localiza offsets [start, end) para contratos 2009-01-01 .. 2010-12-31."""

    def date_at(offset: int) -> str:
        payload = {
            "resource_id": RESOURCE_ID,
            "limit": 1,
            "offset": offset,
            "sort": "data_da_contratacao asc",
            "fields": ["data_da_contratacao"],
        }
        r = session.post(DATASTORE_URL, json=payload, timeout=120)
        r.raise_for_status()
        recs = r.json()["result"]["records"]
        return recs[0]["data_da_contratacao"][:10]

    total = session.post(
        DATASTORE_URL,
        json={"resource_id": RESOURCE_ID, "limit": 0},
        timeout=60,
    ).json()["result"]["total"]

    def first_ge(target: str, lo: int, hi: int) -> int:
        while lo < hi:
            mid = (lo + hi) // 2
            d = date_at(mid)
            print(f"  offset={mid} date={d}", flush=True)
            if d < target:
                lo = mid + 1
            else:
                hi = mid
        return lo

    print("Buscando início de 2009...", flush=True)
    start = first_ge("2009-01-01", 0, total - 1)
    print("Buscando início de 2011...", flush=True)
    end = first_ge("2011-01-01", start, total - 1)
    print(f"Offsets: start={start} end={end} (n≈{end - start})", flush=True)
    return start, end


def fetch_operacoes(
    session: requests.Session,
    offset_start: int,
    offset_end: int,
    page_size: int = 32000,
) -> pd.DataFrame:
    """Baixa páginas do datastore no intervalo de offsets."""
    rows: list[dict] = []
    offset = offset_start
    while offset < offset_end:
        limit = min(page_size, offset_end - offset)
        payload = {
            "resource_id": RESOURCE_ID,
            "limit": limit,
            "offset": offset,
            "sort": "data_da_contratacao asc",
            "fields": NEEDED_FIELDS,
        }
        for attempt in range(5):
            try:
                r = session.post(DATASTORE_URL, json=payload, timeout=180)
                r.raise_for_status()
                data = r.json()
                if not data.get("success"):
                    raise RuntimeError(data)
                batch = data["result"]["records"]
                break
            except Exception as exc:
                wait = 2 ** attempt
                print(f"  retry offset={offset} ({exc}); wait {wait}s", flush=True)
                time.sleep(wait)
        else:
            raise RuntimeError(f"Falha ao baixar offset={offset}")

        if not batch:
            break
        rows.extend(batch)
        offset += len(batch)
        print(f"  baixados {len(rows):,} / {offset_end - offset_start:,}", flush=True)
        if len(batch) < limit:
            break

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df.rename(columns=COLUMN_MAP)
    return df


def load_from_excel(path: Path) -> pd.DataFrame:
    """Compatível com o download Excel do portal (header na linha 6)."""
    return pd.read_excel(path, sheet_name="operacoes_indiretas_automaticas", header=5)


def _payment_dates(data_contr: pd.Timestamp, carencia: int, n: int) -> list:
    """Datas de pagamento equivalentes a data_contr + relativedelta(months=carencia+p)."""
    # relativedelta clampa o dia ao último dia do mês-alvo (ex.: 31/jan → 28/fev).
    return [
        (data_contr + relativedelta(months=carencia + p)).date() for p in range(1, n + 1)
    ]


def gerar_fluxos_chunk(df: pd.DataFrame) -> pd.DataFrame:
    """Gera linhas de fluxo para um lote de contratos (lógica do script original).

    Vetoriza saldos/subsídios/impactos com NumPy por contrato.
    """
    selic_m = TAXA_SELIC_ANUAL / 12
    out_contratos: list = []
    out_parcelas: list[np.ndarray] = []
    out_datas: list = []
    out_amort: list[np.ndarray] = []
    out_saldo: list[np.ndarray] = []
    out_sub: list[np.ndarray] = []
    out_imp: list[np.ndarray] = []

    datas = pd.to_datetime(df["Data da contratação"], dayfirst=True, errors="coerce")
    valores = pd.to_numeric(df["Valor Desembolsado R$ (*)"], errors="coerce").to_numpy()
    juros = (pd.to_numeric(df["Juros"], errors="coerce") / 100.0).to_numpy()
    carencias = pd.to_numeric(df["Prazo - Carência (meses)"], errors="coerce").to_numpy()
    ns = pd.to_numeric(df["Prazo - Amortização (meses)"], errors="coerce").to_numpy()
    contratos = (
        df["_id"].to_numpy() if "_id" in df.columns else df.index.to_numpy()
    )
    data_vals = datas.to_numpy()

    for i in range(len(df)):
        data_contr = data_vals[i]
        valor = valores[i]
        taxa_juros = juros[i]
        carencia = carencias[i]
        n = ns[i]
        if (
            pd.isna(data_contr)
            or pd.isna(valor)
            or pd.isna(taxa_juros)
            or pd.isna(carencia)
            or pd.isna(n)
            or valor <= 0
            or n <= 0
        ):
            continue

        n_i = int(n)
        car_i = int(carencia)
        data_ts = pd.Timestamp(data_contr)
        amort = valor / n_i
        parcelas = np.arange(1, n_i + 1, dtype=np.int32)
        saldos = valor - amort * (parcelas - 1)
        subsidios = (selic_m - taxa_juros / 12.0) * saldos

        y0, m0 = data_ts.year, data_ts.month
        abs_month = m0 + car_i + parcelas
        y_pag = y0 + (abs_month - 1) // 12
        m_pag = (abs_month - 1) % 12 + 1
        meses_ate = (DATA_IMPACTO.year - y_pag) * 12 + (DATA_IMPACTO.month - m_pag)
        impactos = subsidios * np.power(1.0 + selic_m, meses_ate)

        out_contratos.append(np.full(n_i, contratos[i]))
        out_parcelas.append(parcelas)
        out_datas.extend(_payment_dates(data_ts, car_i, n_i))
        out_amort.append(np.full(n_i, round(amort, 2)))
        out_saldo.append(np.round(saldos, 2))
        out_sub.append(np.round(subsidios, 2))
        out_imp.append(np.round(impactos, 2))

    if not out_parcelas:
        return pd.DataFrame(
            columns=[
                "Contrato",
                "Parcela",
                "Data_Pagamento",
                "Amortizacao",
                "Saldo_Devedor",
                "Subsídio",
                "Impacto_Fiscal_2026",
            ]
        )

    return pd.DataFrame(
        {
            "Contrato": np.concatenate(out_contratos),
            "Parcela": np.concatenate(out_parcelas),
            "Data_Pagamento": out_datas,
            "Amortizacao": np.concatenate(out_amort),
            "Saldo_Devedor": np.concatenate(out_saldo),
            "Subsídio": np.concatenate(out_sub),
            "Impacto_Fiscal_2026": np.concatenate(out_imp),
        }
    )


def processar_em_lotes(
    df: pd.DataFrame,
    csv_path: Path,
    lote: int = 2000,
) -> dict:
    """Processa contratos em lotes, grava CSV detalhado e acumula resumo mensal."""
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    if csv_path.exists():
        csv_path.unlink()

    monthly: dict[str, float] = {}
    total_impacto = 0.0
    total_subsidio = 0.0
    n_parcelas = 0
    n_contratos_ok = 0
    wrote_header = False

    n = len(df)
    for start in range(0, n, lote):
        chunk = df.iloc[start : start + lote]
        fluxos = gerar_fluxos_chunk(chunk)
        if fluxos.empty:
            print(f"  lote {start:,}-{start + len(chunk):,}: 0 fluxos", flush=True)
            continue

        n_contratos_ok += fluxos["Contrato"].nunique()
        n_parcelas += len(fluxos)
        total_impacto += float(fluxos["Impacto_Fiscal_2026"].sum())
        total_subsidio += float(fluxos["Subsídio"].sum())

        # Agrega por mês de pagamento
        keys = pd.to_datetime(fluxos["Data_Pagamento"]).dt.to_period("M").astype(str)
        for k, v in fluxos.groupby(keys, sort=False)["Impacto_Fiscal_2026"].sum().items():
            monthly[k] = monthly.get(k, 0.0) + float(v)

        fluxos.to_csv(
            csv_path,
            mode="a",
            index=False,
            header=not wrote_header,
        )
        wrote_header = True
        print(
            f"  lote {start:,}-{start + len(chunk):,} → +{len(fluxos):,} parcelas "
            f"(acum {n_parcelas:,})",
            flush=True,
        )

    return {
        "n_contratos_entrada": n,
        "n_contratos_ok": n_contratos_ok,
        "n_parcelas": n_parcelas,
        "total_subsidio": round(total_subsidio, 2),
        "total_impacto_fiscal_2026": round(total_impacto, 2),
        "monthly": monthly,
    }


def salvar_excel_resumo(stats: dict, xlsx_path: Path, sample_csv: Path | None = None) -> None:
    """Grava Excel com totais e série mensal (cabe no limite do Excel)."""
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    resumo = pd.DataFrame(
        [
            {"Indicador": "Taxa SELIC anual (config)", "Valor": TAXA_SELIC_ANUAL},
            {"Indicador": "Data de impacto", "Valor": DATA_IMPACTO.date().isoformat()},
            {"Indicador": "Contratos na entrada", "Valor": stats["n_contratos_entrada"]},
            {"Indicador": "Contratos processados", "Valor": stats["n_contratos_ok"]},
            {"Indicador": "Parcelas geradas", "Valor": stats["n_parcelas"]},
            {"Indicador": "Soma Subsídio (nominal)", "Valor": stats["total_subsidio"]},
            {
                "Indicador": "Soma Impacto Fiscal 2026",
                "Valor": stats["total_impacto_fiscal_2026"],
            },
            {
                "Indicador": "Arquivo detalhado",
                "Valor": "output/fluxos_gerados.csv",
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

        # Amostra das primeiras linhas do CSV (se existir) para inspeção
        if sample_csv and sample_csv.exists():
            sample = pd.read_csv(sample_csv, nrows=50_000)
            sample.to_excel(writer, sheet_name="Amostra_Parcelas", index=False)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--excel",
        type=Path,
        help="Caminho de um .xlsx local (header=5), em vez da API BNDES",
    )
    parser.add_argument("--discover-offsets", action="store_true")
    parser.add_argument("--offset-start", type=int, default=DEFAULT_OFFSET_START)
    parser.add_argument("--offset-end", type=int, default=DEFAULT_OFFSET_END)
    parser.add_argument("--lote", type=int, default=2000)
    parser.add_argument(
        "--max-contratos",
        type=int,
        default=None,
        help="Limita o número de contratos (útil para testes)",
    )
    args = parser.parse_args(argv)

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    cache_path = DATA_DIR / "operacoes_2009_2010.parquet"
    csv_path = OUTPUT_DIR / "fluxos_gerados.csv"
    xlsx_path = OUTPUT_DIR / "fluxos_gerados.xlsx"

    if args.excel:
        print(f"Lendo Excel: {args.excel}", flush=True)
        df = load_from_excel(args.excel)
    elif cache_path.exists() and not args.discover_offsets:
        print(f"Lendo cache: {cache_path}", flush=True)
        df = pd.read_parquet(cache_path)
    else:
        session = requests.Session()
        session.headers.update({"User-Agent": "SEC-data-analysys/1.0"})
        start, end = args.offset_start, args.offset_end
        if args.discover_offsets:
            start, end = discover_offsets(session)
        print(f"Baixando operações offsets [{start}, {end}) ...", flush=True)
        df = fetch_operacoes(session, start, end)
        if df.empty:
            print("Nenhuma operação retornada.", file=sys.stderr)
            return 1
        # Garante filtro de data mesmo se offsets desatualizarem
        datas = pd.to_datetime(df["Data da contratação"])
        mask = (datas >= "2009-01-01") & (datas < "2011-01-01")
        df = df.loc[mask].copy()
        df.to_parquet(cache_path, index=False)
        print(f"Cache salvo: {cache_path} ({len(df):,} linhas)", flush=True)

    if args.max_contratos is not None:
        df = df.head(args.max_contratos).copy()
        print(f"Limitado a {len(df):,} contratos (--max-contratos)", flush=True)

    print(f"Processando {len(df):,} contratos...", flush=True)
    stats = processar_em_lotes(df, csv_path, lote=args.lote)
    print(json.dumps({k: v for k, v in stats.items() if k != "monthly"}, indent=2), flush=True)

    salvar_excel_resumo(stats, xlsx_path, sample_csv=csv_path)
    print(f"✅ CSV detalhado: {csv_path}", flush=True)
    print(f"✅ Excel resumo:  {xlsx_path}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
