#!/usr/bin/env python3
"""
Gera fluxos mensais detalhados (carência + amortização SAC) e impacto fiscal
a valor de 30/06/2026, a partir de operações indiretas automáticas do BNDES.

Metodologia ContAgil (taxas compostas) + carência corrigida:
  - taxa_mensal = (1 + taxa_aa)^(1/12) - 1
  - Fluxos em TODOS os meses (carência + amortização)
  - Amortização constante só após a carência
  - spread = (1 + (SELIC_m − taxa_contrato_m))^n
  - subsídio = saldo × (SELIC_m − taxa_contrato_m)
  - impacto_fiscal:
      * com Excel STP ContAgil: subsídio × fator(30/06/2026) / fator(data_fluxo)
      * sem STP: subsídio × (1 + SELIC_m)^(meses até 30/06/2026)

Entrada:
  - Excel do portal (header=5), ou
  - CSV aberto do BNDES (download automático 2009–2010)
  - (opcional) Excel SELIC ContAgil STP-*.xlsx

Saídas:
  - output/fluxos_completos_final.csv/.xlsx   (detalhe por parcela + colunas extras)
  - output/resumo_por_agente.csv|.xlsx       (agregado por instituição financeira)
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

# ===================== CONFIGURAÇÕES =====================
TAXA_SELIC_ANUAL = 0.145  # 14,5% a.a.
DATA_IMPACTO = datetime(2026, 6, 30)
AGENTE_NAO_INFORMADO = "Não informado"


def taxa_mensal_composta(taxa_aa: float) -> float:
    """Converte taxa anual em taxa mensal composta: (1+r)^(1/12)-1."""
    return (1.0 + float(taxa_aa)) ** (1.0 / 12.0) - 1.0


TAXA_SELIC_MENSAL = taxa_mensal_composta(TAXA_SELIC_ANUAL)


class SelicSerie:
    """Lookup de fatores Selic acumulados (Excel STP ContAgil / Bacen)."""

    def __init__(self, datas: np.ndarray, fatores: np.ndarray):
        self.datas = datas
        self.fatores = fatores

    @classmethod
    def from_excel(cls, path: Path) -> "SelicSerie":
        selic = pd.read_excel(path)
        datas = pd.to_datetime(selic.iloc[:, 0], dayfirst=True, errors="coerce").values.astype(
            "datetime64[ns]"
        )
        col_fator = min(3, selic.shape[1] - 1)
        fatores = pd.to_numeric(selic.iloc[:, col_fator], errors="coerce").values
        mask = ~pd.isna(datas) & ~pd.isna(fatores) & (fatores > 0)
        return cls(datas[mask], fatores[mask].astype(float))

    def fator_rapido(self, datas) -> np.ndarray:
        datas_arr = np.array(datas, dtype="datetime64[ns]")
        idx = np.searchsorted(self.datas, datas_arr, side="right") - 1
        idx = np.clip(idx, 0, len(self.fatores) - 1)
        return self.fatores[idx]

    def capitalizar(
        self,
        valor: float,
        data_fluxo: datetime,
        data_impacto: datetime = DATA_IMPACTO,
    ) -> float:
        """impacto = valor × fator(impacto) / fator(fluxo), como no ContAgil."""
        if valor <= 0:
            return 0.0
        f_fluxo = float(
            self.fator_rapido([np.datetime64(pd.Timestamp(data_fluxo), "ns")])[0]
        )
        f_fim = float(
            self.fator_rapido([np.datetime64(pd.Timestamp(data_impacto), "ns")])[0]
        )
        if f_fluxo <= 0 or np.isnan(f_fluxo) or np.isnan(f_fim):
            return float(valor)
        if f_fim > f_fluxo:
            return round(valor * (f_fim / f_fluxo), 2)
        return round(valor, 2)

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
    "Instituição Financeira Credenciada": "agente",
    "Instituicao Financeira Credenciada": "agente",
}

CSV_COLUMNS = {
    "data_da_contratacao": "data_contratacao",
    "valor_desembolsado_reais": "valor_desembolsado",
    "juros": "juros",
    "prazo_carencia_meses": "prazo_carencia",
    "prazo_amortizacao_meses": "prazo_amortizacao",
    "instituicao_financeira_credenciada": "agente",
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


def _stream_download(url: str, dest: Path, retries: int = 4) -> Path:
    """Baixa arquivo grande via HTTP streaming com retries (pandas URL buffer falha em ~1GB)."""
    import requests

    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".part")
    last_err: Exception | None = None

    for attempt in range(1, retries + 1):
        try:
            resume_from = tmp.stat().st_size if tmp.exists() else 0
            headers = {"Range": f"bytes={resume_from}-"} if resume_from else {}
            mode = "ab" if resume_from else "wb"

            print(
                f"Download tentativa {attempt}/{retries}"
                + (f" (retomando de {resume_from:,} bytes)" if resume_from else "")
                + "..."
            )
            with requests.get(url, stream=True, timeout=120, headers=headers) as resp:
                if resp.status_code not in (200, 206):
                    resp.raise_for_status()
                if resp.status_code == 200 and resume_from:
                    mode = "wb"
                    resume_from = 0

                total = resp.headers.get("Content-Length")
                expected = (
                    int(total) + resume_from
                    if total and resp.status_code == 206
                    else (int(total) if total else None)
                )

                downloaded = resume_from
                with tmp.open(mode) as f:
                    for chunk in resp.iter_content(chunk_size=8 * 1024 * 1024):
                        if not chunk:
                            continue
                        f.write(chunk)
                        downloaded += len(chunk)
                        if expected and downloaded % (64 * 1024 * 1024) < 8 * 1024 * 1024:
                            pct = 100.0 * downloaded / expected
                            print(f"  baixados {downloaded:,}/{expected:,} ({pct:.1f}%)")

            if expected is not None and tmp.stat().st_size < expected:
                raise IOError(
                    f"Download incompleto: {tmp.stat().st_size:,} < {expected:,} bytes"
                )

            tmp.replace(dest)
            print(f"CSV bruto salvo: {dest} ({dest.stat().st_size:,} bytes)")
            return dest
        except Exception as exc:  # noqa: BLE001 — retries em rede
            last_err = exc
            print(f"  falha: {exc}")
            time.sleep(min(2**attempt, 32))

    raise RuntimeError(f"Falha ao baixar {url}: {last_err}")


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

    raw_path = RAW_DIR / "operacoes-indiretas-automaticas.csv"
    if not raw_path.exists() or raw_path.stat().st_size < 100_000_000:
        _stream_download(url, raw_path)
    else:
        print(f"Usando cache local: {raw_path} ({raw_path.stat().st_size:,} bytes)")

    encoding = "utf-8"
    try:
        with raw_path.open("r", encoding="utf-8") as f:
            f.read(2048)
    except UnicodeDecodeError:
        encoding = "cp1252"

    reader = pd.read_csv(
        raw_path,
        sep=";",
        encoding=encoding,
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


def _normalizar_agente(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace(
        {
            "": AGENTE_NAO_INFORMADO,
            "nan": AGENTE_NAO_INFORMADO,
            "None": AGENTE_NAO_INFORMADO,
            "NaT": AGENTE_NAO_INFORMADO,
        }
    )
    return s.fillna(AGENTE_NAO_INFORMADO)


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

    if "agente" in df.columns:
        agente = _normalizar_agente(df["agente"])
    else:
        agente = pd.Series([AGENTE_NAO_INFORMADO] * len(df), index=df.index)

    out = pd.DataFrame(
        {
            "data_contratacao": parse_datas(df["data_contratacao"]),
            "valor_desembolsado": limpar_valor(df["valor_desembolsado"]),
            "juros": limpar_valor(df["juros"]),
            "prazo_carencia": limpar_valor(df["prazo_carencia"]).fillna(0),
            "prazo_amortizacao": limpar_valor(df["prazo_amortizacao"]),
            "agente": agente.values,
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
    print(f"Agentes distintos: {out['agente'].nunique():,}")
    return out


def _coluna_impacto(df: pd.DataFrame) -> str:
    if "impacto_fiscal" in df.columns:
        return "impacto_fiscal"
    if "impacto" in df.columns:
        return "impacto"
    raise ValueError("fluxos precisam de 'impacto_fiscal' ou 'impacto'")


def agregar_por_agente(df_fluxos: pd.DataFrame, contratos: pd.DataFrame) -> pd.DataFrame:
    """
    Resume fluxos por Agente Financeiro (Instituição Financeira Credenciada).

    Correção vs script com merge por índice: o CSV de fluxos é por parcela;
    o vínculo correto é contrato → agente (não left_index/right_index).
    """
    df = df_fluxos.copy()
    impacto_col = _coluna_impacto(df)

    if "Instituição Financeira" in df.columns:
        df["Agente"] = df["Instituição Financeira"].fillna(AGENTE_NAO_INFORMADO)
    elif "agente" in contratos.columns:
        mapa = contratos.set_index("contrato")["agente"]
        df["Agente"] = df["contrato"].map(mapa).fillna(AGENTE_NAO_INFORMADO)
    else:
        raise ValueError("contratos precisa da coluna 'agente' ou fluxos de 'Instituição Financeira'")

    resumo = (
        df.groupby("Agente", dropna=False)
        .agg(
            qtd_contratos=("contrato", "nunique"),
            subsidio=("subsidio", "sum"),
            impacto=(impacto_col, "sum"),
        )
        .round(2)
        .reset_index()
    )
    resumo.columns = [
        "Agente",
        "Qtd Contratos",
        "Total Subsídio (R$)",
        "Impacto Fiscal 2026 (R$)",
    ]
    return resumo.sort_values("Total Subsídio (R$)", ascending=False).reset_index(drop=True)


def resumo_from_agent_agg(agent_agg: dict) -> pd.DataFrame:
    """Converte acumulador interno {agente: {contratos, subsidio, impacto}} em DataFrame."""
    if not agent_agg:
        return pd.DataFrame(
            columns=[
                "Agente",
                "Qtd Contratos",
                "Total Subsídio (R$)",
                "Impacto Fiscal 2026 (R$)",
            ]
        )

    rows = [
        {
            "Agente": agente,
            "Qtd Contratos": vals["contratos"],
            "Total Subsídio (R$)": round(vals["subsidio"], 2),
            "Impacto Fiscal 2026 (R$)": round(vals["impacto"], 2),
        }
        for agente, vals in agent_agg.items()
    ]
    resumo = pd.DataFrame(rows)
    return resumo.sort_values("Total Subsídio (R$)", ascending=False).reset_index(drop=True)


def gerar_fluxos_contrato(
    data_contr: pd.Timestamp,
    valor: float,
    taxa_juros_aa: float,
    carencia: int,
    n: int,
    contrato_id: int,
    instituicao: str = AGENTE_NAO_INFORMADO,
    selic_aa: float = TAXA_SELIC_ANUAL,
    data_impacto: datetime = DATA_IMPACTO,
    selic_serie: SelicSerie | None = None,
) -> list[dict]:
    """
    Gera fluxos detalhados de UM contrato (carência + amortização).

    Correção vs script ContAgil com bug:
      O original fazia `data = contr+(carencia+p)` E `em_carencia = p <= carencia`
      no loop `p=1..n`, o que zera amortização nas primeiras parcelas pós-carência
      e deixa saldo residual. Aqui o cronograma cobre carência+n meses.

    Colunas extras (pedido ContAgil): Instituição Financeira, taxas compostas,
    spread e impacto_fiscal.
    """
    if n <= 0 or valor <= 0:
        return []

    amort_mensal = valor / n
    saldo = valor
    taxa_contrato_mensal = taxa_mensal_composta(taxa_juros_aa)
    taxa_selic_mensal = taxa_mensal_composta(selic_aa)
    spread = (1.0 + (taxa_selic_mensal - taxa_contrato_mensal)) ** n
    fluxos: list[dict] = []

    # Dia 15 como no ContAgil (evita deslocamento de fim de mês)
    try:
        data_base = data_contr.replace(day=15)
    except ValueError:
        data_base = data_contr

    total_meses = carencia + n
    for p in range(1, total_meses + 1):
        data_fluxo = data_base + relativedelta(months=p - 1)
        em_carencia = p <= carencia
        amort = 0.0 if em_carencia else amort_mensal
        subsidio = saldo * (taxa_selic_mensal - taxa_contrato_mensal)

        if selic_serie is not None:
            impacto = selic_serie.capitalizar(
                subsidio, data_fluxo.to_pydatetime(), data_impacto
            )
        else:
            meses = meses_ate_impacto(data_fluxo.to_pydatetime(), data_impacto)
            impacto = round(subsidio * ((1.0 + taxa_selic_mensal) ** meses), 2)

        fluxos.append(
            {
                "contrato": contrato_id,
                "Instituição Financeira": instituicao,
                "mes": p,
                "data_fluxo": data_fluxo.date(),
                "saldo": round(saldo, 2),
                "amortizacao": round(amort, 2),
                "taxa_selic_mensal": round(taxa_selic_mensal, 8),
                "taxa_contrato_mensal": round(taxa_contrato_mensal, 8),
                "spread": round(spread, 6),
                "subsidio": round(subsidio, 2),
                "impacto_fiscal": impacto,
                "em_carencia": em_carencia,
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
    selic_serie: SelicSerie | None = None,
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
            instituicao = str(
                getattr(row, "agente", AGENTE_NAO_INFORMADO) or AGENTE_NAO_INFORMADO
            )
            records.extend(
                gerar_fluxos_contrato(
                    data_contr=data_contr,
                    valor=float(row.valor_desembolsado),
                    taxa_juros_aa=float(row.juros) / 100.0,
                    carencia=int(float(row.prazo_carencia or 0)),
                    n=int(float(row.prazo_amortizacao)),
                    contrato_id=int(row.contrato),
                    instituicao=instituicao,
                    selic_aa=selic_aa,
                    data_impacto=data_impacto,
                    selic_serie=selic_serie,
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
    selic_serie: SelicSerie | None = None,
) -> dict:
    """Processa em lotes, grava CSV detalhado e acumula estatísticas (+ por agente)."""
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    if csv_path.exists():
        csv_path.unlink()

    monthly: dict[str, float] = {}
    agent_agg: dict[str, dict[str, float]] = defaultdict(
        lambda: {"contratos": 0, "subsidio": 0.0, "impacto": 0.0}
    )
    total_impacto = 0.0
    total_subsidio = 0.0
    total_amort = 0.0
    n_parcelas = 0
    n_contratos_ok = 0
    n_em_carencia = 0
    wrote_header = False

    n = len(df)
    modo = "fator STP ContAgil" if selic_serie is not None else "SELIC composta constante"
    print(f"Processando {n:,} contratos (lote={lote:,}, impacto={modo})...")

    for start in range(0, n, lote):
        chunk = df.iloc[start : start + lote]
        records: list[dict] = []

        for row in chunk.itertuples(index=False):
            try:
                data_contr = pd.Timestamp(row.data_contratacao)
                if pd.isna(data_contr):
                    continue
                instituicao = str(
                    getattr(row, "agente", AGENTE_NAO_INFORMADO) or AGENTE_NAO_INFORMADO
                )
                fluxos = gerar_fluxos_contrato(
                    data_contr=data_contr,
                    valor=float(row.valor_desembolsado),
                    taxa_juros_aa=float(row.juros) / 100.0,
                    carencia=int(float(row.prazo_carencia or 0)),
                    n=int(float(row.prazo_amortizacao)),
                    contrato_id=int(row.contrato),
                    instituicao=instituicao,
                    selic_aa=selic_aa,
                    data_impacto=data_impacto,
                    selic_serie=selic_serie,
                )
                if fluxos:
                    n_contratos_ok += 1
                    records.extend(fluxos)
                    agent_agg[instituicao]["contratos"] += 1
                    agent_agg[instituicao]["subsidio"] += float(
                        sum(f["subsidio"] for f in fluxos)
                    )
                    agent_agg[instituicao]["impacto"] += float(
                        sum(f["impacto_fiscal"] for f in fluxos)
                    )
            except (TypeError, ValueError, OverflowError):
                continue

        if not records:
            print(f"  lote {start:,}-{start + len(chunk):,}: 0 fluxos")
            continue

        fluxos_df = pd.DataFrame(records)
        n_parcelas += len(fluxos_df)
        total_impacto += float(fluxos_df["impacto_fiscal"].sum())
        total_subsidio += float(fluxos_df["subsidio"].sum())
        total_amort += float(fluxos_df["amortizacao"].sum())
        n_em_carencia += int(fluxos_df["em_carencia"].sum())

        keys = pd.to_datetime(fluxos_df["data_fluxo"]).dt.to_period("M").astype(str)
        for k, v in fluxos_df.groupby(keys, sort=False)["impacto_fiscal"].sum().items():
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
        "n_agentes": len(agent_agg),
        "metodologia_impacto": modo,
        "taxa_selic_anual": selic_aa,
        "taxa_selic_mensal_composta": round(taxa_mensal_composta(selic_aa), 8),
        "monthly": monthly,
        "por_agente": dict(agent_agg),
    }


def salvar_resumo_por_agente(resumo: pd.DataFrame, stem: str = "resumo_por_agente") -> tuple[Path, Path]:
    """Grava CSV + Excel do ranking por agente financeiro."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    csv_path = OUTPUT_DIR / f"{stem}.csv"
    xlsx_path = OUTPUT_DIR / f"{stem}.xlsx"
    resumo.to_csv(csv_path, index=False)
    resumo.to_excel(xlsx_path, index=False, sheet_name="Por_Agente")
    return csv_path, xlsx_path


def salvar_excel_resumo(
    stats: dict,
    xlsx_path: Path,
    sample_csv: Path | None = None,
    sample_rows: int = 50_000,
    resumo_agente: pd.DataFrame | None = None,
) -> None:
    """Excel legível: resumo + por agente + impacto mensal + amostra de parcelas."""
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
            {"Indicador": "Agentes financeiros", "Valor": stats.get("n_agentes", 0)},
            {
                "Indicador": "Metodologia impacto",
                "Valor": stats.get("metodologia_impacto", "SELIC composta constante"),
            },
            {
                "Indicador": "Arquivo detalhado",
                "Valor": str(xlsx_path.with_suffix(".csv").name),
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
        if resumo_agente is not None and not resumo_agente.empty:
            resumo_agente.to_excel(writer, sheet_name="Por_Agente", index=False)
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
        default="fluxos_completos_final",
        help="Prefixo dos arquivos de saída.",
    )
    p.add_argument(
        "--arquivo-selic",
        type=Path,
        default=None,
        help="Excel STP ContAgil/Bacen (fatores acumulados). Se omitido, usa SELIC composta constante.",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    selic_serie: SelicSerie | None = None
    if args.arquivo_selic is not None:
        print(f"Lendo SELIC STP: {args.arquivo_selic}")
        selic_serie = SelicSerie.from_excel(args.arquivo_selic)
        print(f"  {len(selic_serie.datas):,} pontos de fator Selic carregados")

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

    stats = processar_em_lotes(
        df, csv_path, lote=args.lote, selic_serie=selic_serie
    )
    resumo_agente = resumo_from_agent_agg(stats.get("por_agente", {}))
    agente_csv, agente_xlsx = salvar_resumo_por_agente(resumo_agente)

    printable = {
        k: v for k, v in stats.items() if k not in {"monthly", "por_agente"}
    }
    print(json.dumps(printable, indent=2))
    print("\nResumo por Agente Financeiro (top 20):")
    print(resumo_agente.head(20).to_string(index=False))

    salvar_excel_resumo(
        stats, xlsx_path, sample_csv=csv_path, resumo_agente=resumo_agente
    )
    with stats_path.open("w", encoding="utf-8") as f:
        json.dump(printable, f, indent=2)

    print(f"✅ CSV detalhado: {csv_path}")
    print(f"✅ Excel resumo:  {xlsx_path}")
    print(f"✅ Resumo agente: {agente_csv}")
    print(f"✅ Resumo agente: {agente_xlsx}")
    print(f"✅ Stats JSON:    {stats_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
