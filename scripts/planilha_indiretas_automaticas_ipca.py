#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Planilha das operações indiretas AUTOMÁTICAS do BNDES — uma aba por ano.

Colunas (por operação):
  - Nome Agente Financeiro
  - Nome Cliente
  - Data da Contratação
  - Prazo Carência
  - Prazo Amortização
  - Valor Desembolsado
  - Valor Desembolsado atualizado pelo IPCA até julho/2026

Ordenação em cada aba: agente financeiro + data da contratação.
Ao final de cada aba: totais de desembolso corrente e IPCA por agente.

Fontes (BNDES — operações automáticas):
  https://www.bndes.gov.br/arquivos/central-downloads/operacoes_financiamento/automaticas/

Uso::

  python scripts/planilha_indiretas_automaticas_ipca.py
  python scripts/planilha_indiretas_automaticas_ipca.py --pasta-cache data/bndes_automaticas
"""

from __future__ import annotations

import argparse
import re
import sys
import time
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.calcular_diretas_ipca_selic import (  # noqa: E402
    _baixar_sgs,
    _idx_mes,
    carregar_ipca,
)

MARKER = "indiretas-automaticas-ipca-20260816a"
DATA_REF_DEFAULT = datetime(2026, 7, 31)

BASE_URL = (
    "https://www.bndes.gov.br/arquivos/central-downloads/"
    "operacoes_financiamento/automaticas"
)

ARQUIVOS = (
    "operacoes_indiretas_automaticas_2002-01-01_ate_2008-12-31.xlsx",
    "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx",
    "operacoes_indiretas_automaticas_2011-01-01_ate_2011-12-31.xlsx",
    "operacoes_indiretas_automaticas_2012-01-01_ate_2012-12-31.xlsx",
    "operacoes_indiretas_automaticas_2013-01-01_ate_2013-12-31.xlsx",
    "operacoes_indiretas_automaticas_2014-01-01_ate_2014-12-31.xlsx",
    "operacoes_indiretas_automaticas_2015-01-01_ate_2016-12-31.xlsx",
    "operacoes_indiretas_automaticas_2017-01-01_ate_2026-06-30.xlsx",
)

COL_AGENTE = "Nome Agente Financeiro"
COL_CLIENTE = "Nome Cliente"
COL_DATA = "Data da Contratação"
COL_CAR = "Prazo Carência"
COL_AMORT = "Prazo Amortização"
COL_VALOR = "Valor Desembolsado"
COL_IPCA = "Valor Desembolsado atualizado pelo IPCA até julho/2026"

COLS_OUT = (
    COL_AGENTE,
    COL_CLIENTE,
    COL_DATA,
    COL_CAR,
    COL_AMORT,
    COL_VALOR,
    COL_IPCA,
)


def _norm(s: str) -> str:
    s = str(s).replace("\n", " ").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def detectar_header(path: Path, max_linhas: int = 15) -> int:
    """Localiza a linha do cabeçalho (contém Cliente + Data da contratação)."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        raw = pd.read_excel(path, sheet_name=0, header=None, nrows=max_linhas)
    for i in range(len(raw)):
        cells = [_norm(x) for x in raw.iloc[i].tolist() if pd.notna(x)]
        tem_cliente = any(c == "cliente" for c in cells)
        tem_data = any("data da contrata" in c for c in cells)
        if tem_cliente and tem_data:
            return i
    raise ValueError(f"Cabeçalho não encontrado em {path.name}")


def _achar_coluna(columns: list, *candidatos: str) -> str:
    """Resolve coluna por nome canônico (exato primeiro, depois substring)."""
    norms = {_norm(c): c for c in columns}
    for cand in candidatos:
        alvo = _norm(cand)
        if alvo in norms:
            return norms[alvo]
    for cand in candidatos:
        alvo = _norm(cand)
        for n, orig in norms.items():
            if n.startswith(alvo) or alvo == n:
                return orig
    for cand in candidatos:
        alvo = _norm(cand)
        hits = [(n, orig) for n, orig in norms.items() if alvo in n]
        # evita "porte do cliente" quando se pede "cliente"
        hits = [(n, o) for n, o in hits if not n.startswith("porte ")]
        if len(hits) == 1:
            return hits[0][1]
        if hits:
            hits.sort(key=lambda x: len(x[0]))
            return hits[0][1]
    raise ValueError(
        f"Coluna não encontrada ({candidatos}). Disponíveis: {columns}"
    )


def baixar_arquivo(nome: str, pasta: Path, timeout: int = 300) -> Path:
    pasta.mkdir(parents=True, exist_ok=True)
    dest = pasta / nome
    if dest.exists() and dest.stat().st_size > 1_000_000:
        print(f"[CACHE] {dest.name} ({dest.stat().st_size / 1e6:.1f} MB)")
        return dest
    url = f"{BASE_URL}/{nome}"
    print(f"[DOWNLOAD] {url}")
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    dest.write_bytes(r.content)
    print(f"[OK] {dest.name} ({dest.stat().st_size / 1e6:.1f} MB)")
    return dest


def carregar_operacoes(path: Path) -> pd.DataFrame:
    """Lê um Excel BNDES e devolve colunas normalizadas (sem IPCA)."""
    hdr = detectar_header(path)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(path, sheet_name=0, header=hdr)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

    c_agente = _achar_coluna(
        list(df.columns),
        "instituição financeira credenciada",
        "instituicao financeira credenciada",
        "agente financeiro",
    )
    c_cliente = _achar_coluna(list(df.columns), "cliente")
    c_data = _achar_coluna(list(df.columns), "data da contratação", "data da contratacao")
    c_car = _achar_coluna(list(df.columns), "prazo - carência", "prazo - carencia", "prazo carência")
    c_amort = _achar_coluna(
        list(df.columns), "prazo - amortização", "prazo - amortizacao", "prazo amortização"
    )
    c_valor = _achar_coluna(
        list(df.columns),
        "valor desembolsado",
        "valor desembolsado r$",
    )

    out = pd.DataFrame(
        {
            COL_AGENTE: df[c_agente].astype(str).str.strip(),
            COL_CLIENTE: df[c_cliente].astype(str).str.strip(),
            COL_DATA: pd.to_datetime(df[c_data], dayfirst=True, errors="coerce"),
            COL_CAR: pd.to_numeric(df[c_car], errors="coerce"),
            COL_AMORT: pd.to_numeric(df[c_amort], errors="coerce"),
            COL_VALOR: pd.to_numeric(df[c_valor], errors="coerce"),
        }
    )
    out = out.dropna(subset=[COL_DATA]).copy()
    out[COL_AGENTE] = out[COL_AGENTE].replace({"nan": "", "None": ""}).fillna("")
    out.loc[out[COL_AGENTE] == "", COL_AGENTE] = "(sem agente)"
    out[COL_CLIENTE] = out[COL_CLIENTE].replace({"nan": ""}).fillna("")
    out["_ano"] = out[COL_DATA].dt.year.astype(int)
    return out.reset_index(drop=True)


def aplicar_ipca_vetorizado(
    df: pd.DataFrame,
    ipca: pd.DataFrame,
    data_ref: datetime | pd.Timestamp,
) -> pd.DataFrame:
    """valor_ipca = valor × fator(ref) / fator(mês da contratação)."""
    out = df.copy()
    ref = pd.Timestamp(data_ref).to_period("M").to_timestamp()
    i_ref = _idx_mes(ipca, ref)
    f_ref = float(ipca.loc[i_ref, "fator"])
    mes_ref_usado = pd.Timestamp(ipca.loc[i_ref, "mes"])

    tmp = pd.DataFrame(
        {
            "_i": range(len(out)),
            "mes": out[COL_DATA].dt.to_period("M").dt.to_timestamp(),
        }
    ).sort_values("mes")
    ipca_m = ipca[["mes", "fator"]].sort_values("mes")
    merged = pd.merge_asof(tmp, ipca_m, on="mes", direction="backward")
    merged = merged.sort_values("_i")
    f0 = merged["fator"].to_numpy(dtype=float)
    valor = out[COL_VALOR].to_numpy(dtype=float)
    fator = f_ref / f0
    fator = pd.Series(fator).replace([float("inf"), float("-inf")], pd.NA).to_numpy(
        dtype=float
    )
    ipca_vals = valor * fator
    out[COL_IPCA] = pd.Series(ipca_vals, index=out.index).round(2)
    out.attrs["mes_ipca_ref"] = mes_ref_usado
    out.attrs["fator_ref"] = f_ref
    return out


def totais_por_agente(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby(COL_AGENTE, sort=True)
        .agg(
            **{
                COL_VALOR: (COL_VALOR, "sum"),
                COL_IPCA: (COL_IPCA, "sum"),
                "Qtd Operações": (COL_CLIENTE, "size"),
            }
        )
        .reset_index()
    )
    g[COL_VALOR] = g[COL_VALOR].round(2)
    g[COL_IPCA] = g[COL_IPCA].round(2)
    return g


def ordenar(df: pd.DataFrame) -> pd.DataFrame:
    return df.sort_values([COL_AGENTE, COL_DATA, COL_CLIENTE], kind="mergesort").reset_index(
        drop=True
    )


def _escrever_aba_ano_xlsx(wb, ano: int, df: pd.DataFrame, fmt_data, fmt_num) -> dict:
    detalhe = ordenar(df)[list(COLS_OUT)]
    totais = totais_por_agente(detalhe)
    ws = wb.add_worksheet(str(ano))
    # cabeçalho + detalhe
    for j, c in enumerate(COLS_OUT):
        ws.write(0, j, c)
    for i, tup in enumerate(detalhe.itertuples(index=False, name=None), start=1):
        agente, cliente, data, car, amort, valor, valor_ipca = tup
        ws.write(i, 0, agente if isinstance(agente, str) else str(agente))
        ws.write(i, 1, cliente if isinstance(cliente, str) else str(cliente))
        if isinstance(data, pd.Timestamp) and not pd.isna(data):
            ws.write_datetime(i, 2, data.to_pydatetime(), fmt_data)
        if pd.notna(car):
            ws.write_number(i, 3, float(car))
        if pd.notna(amort):
            ws.write_number(i, 4, float(amort))
        if pd.notna(valor):
            ws.write_number(i, 5, float(valor), fmt_num)
        if pd.notna(valor_ipca):
            ws.write_number(i, 6, float(valor_ipca), fmt_num)

    r = len(detalhe) + 2
    ws.write(r, 0, f"TOTAIS POR AGENTE FINANCEIRO — {ano}")
    r += 1
    ws.write(r, 0, COL_AGENTE)
    ws.write(r, 1, "Qtd Operações")
    ws.write(r, 5, COL_VALOR)
    ws.write(r, 6, COL_IPCA)
    r += 1
    for tup in totais.itertuples(index=False, name=None):
        ag, v, vipca, qtd = tup
        ws.write(r, 0, ag)
        ws.write_number(r, 1, int(qtd))
        ws.write_number(r, 5, float(v), fmt_num)
        ws.write_number(r, 6, float(vipca), fmt_num)
        r += 1
    ws.write(r, 0, "TOTAL GERAL")
    ws.write_number(r, 1, int(totais["Qtd Operações"].sum()))
    ws.write_number(r, 5, float(totais[COL_VALOR].sum()), fmt_num)
    ws.write_number(r, 6, float(totais[COL_IPCA].sum()), fmt_num)

    return {
        "ano": ano,
        "operacoes": len(detalhe),
        "valor": float(detalhe[COL_VALOR].sum()),
        "valor_ipca": float(detalhe[COL_IPCA].sum()),
        "agentes": int(totais.shape[0]),
    }


def escrever_planilha(
    por_ano: dict[int, pd.DataFrame],
    saida: Path,
    *,
    data_ref: datetime,
    mes_ipca_ref: pd.Timestamp | None = None,
) -> Path:
    """Grava workbook com capa, resumo e uma aba por ano (xlsxwriter streaming)."""
    import xlsxwriter

    saida.parent.mkdir(parents=True, exist_ok=True)
    anos = sorted(por_ano.keys())
    wb = xlsxwriter.Workbook(
        str(saida),
        {
            "constant_memory": True,
            "strings_to_urls": False,
            "nan_inf_to_errors": True,
        },
    )
    fmt_data = wb.add_format({"num_format": "dd/mm/yyyy"})
    fmt_num = wb.add_format({"num_format": "#,##0.00"})

    # Capa
    ws = wb.add_worksheet("Capa")
    capa_rows = [
        ("Título", "Operações Indiretas Automáticas BNDES — desembolso atualizado IPCA"),
        ("Fonte", BASE_URL),
        ("Referência IPCA", pd.Timestamp(data_ref).strftime("%Y-%m-%d")),
        (
            "Mês IPCA efetivo",
            pd.Timestamp(mes_ipca_ref).strftime("%Y-%m") if mes_ipca_ref is not None else "",
        ),
        ("Marker", MARKER),
        ("Gerado em", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Anos", f"{anos[0]}–{anos[-1]}" if anos else ""),
        (
            "Observação",
            "Valores de desembolso conforme publicação BNDES. "
            "IPCA: Bacen SGS 433 (% a.m.). Ordenação: agente + data. "
            "Totais por agente ao final de cada aba.",
        ),
    ]
    ws.write(0, 0, "Campo")
    ws.write(0, 1, "Valor")
    for i, (k, v) in enumerate(capa_rows, start=1):
        ws.write(i, 0, k)
        ws.write(i, 1, v)

    resumo_rows = []
    for ano in anos:
        print(f"[ABA] {ano}: {len(por_ano[ano]):,} operações ...")
        sys.stdout.flush()
        info = _escrever_aba_ano_xlsx(wb, ano, por_ano[ano], fmt_data, fmt_num)
        resumo_rows.append(info)

    ws = wb.add_worksheet("Resumo_Anual")
    headers = ["Ano", "Qtd Operações", COL_VALOR, COL_IPCA, "Qtd Agentes"]
    for j, h in enumerate(headers):
        ws.write(0, j, h)
    for i, info in enumerate(resumo_rows, start=1):
        ws.write(i, 0, info["ano"])
        ws.write_number(i, 1, info["operacoes"])
        ws.write_number(i, 2, info["valor"], fmt_num)
        ws.write_number(i, 3, info["valor_ipca"], fmt_num)
        ws.write_number(i, 4, info["agentes"])
    if resumo_rows:
        r = len(resumo_rows) + 1
        ws.write(r, 0, "TOTAL")
        ws.write_number(r, 1, sum(x["operacoes"] for x in resumo_rows))
        ws.write_number(r, 2, sum(x["valor"] for x in resumo_rows), fmt_num)
        ws.write_number(r, 3, sum(x["valor_ipca"] for x in resumo_rows), fmt_num)

    wb.close()
    print(f"[OK] Planilha: {saida} ({saida.stat().st_size / 1e6:.1f} MB)")
    return saida


def carregar_todos(
    pasta_cache: Path,
    *,
    baixar: bool = True,
    anos: set[int] | None = None,
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for nome in ARQUIVOS:
        path = baixar_arquivo(nome, pasta_cache) if baixar else pasta_cache / nome
        if not path.exists():
            raise FileNotFoundError(path)
        print(f"[LER] {path.name} ...")
        sys.stdout.flush()
        t0 = time.time()
        df = carregar_operacoes(path)
        if anos is not None:
            df = df[df["_ano"].isin(anos)]
        print(f"  → {len(df):,} linhas em {time.time() - t0:.1f}s")
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=list(COLS_OUT) + ["_ano"])
    return pd.concat(frames, ignore_index=True)


def processar(
    pasta_cache: Path,
    saida: Path,
    *,
    ipca_path: Path | None = None,
    data_ref: datetime = DATA_REF_DEFAULT,
    baixar: bool = True,
    anos: set[int] | None = None,
) -> Path:
    print(f"[{MARKER}]")
    df = carregar_todos(pasta_cache, baixar=baixar, anos=anos)
    print(f"[INFO] Total operações: {len(df):,}")

    print("[INFO] Carregando IPCA (Bacen SGS 433)...")
    if ipca_path is not None and ipca_path.exists():
        ipca = carregar_ipca(ipca_path)
        if ipca["mes"].min() > pd.Timestamp("2002-01-01"):
            print("[AVISO] IPCA local começa depois de 2002; completando via Bacen...")
            raw = _baixar_sgs(433, inicio="01/01/2002")
            raw["fator"] = (1.0 + raw["valor"] / 100.0).cumprod()
            ipca = raw
    else:
        raw = _baixar_sgs(433, inicio="01/01/2002")
        raw = raw.sort_values("mes").drop_duplicates("mes").copy()
        raw["fator"] = (1.0 + raw["valor"] / 100.0).cumprod()
        ipca = raw.reset_index(drop=True)

    df = aplicar_ipca_vetorizado(df, ipca, data_ref)
    mes_ref = df.attrs.get("mes_ipca_ref")
    print(f"[INFO] IPCA ref pedido={pd.Timestamp(data_ref):%Y-%m} efetivo={mes_ref:%Y-%m}")

    por_ano: dict[int, pd.DataFrame] = {
        int(ano): g.drop(columns=["_ano"]) for ano, g in df.groupby("_ano", sort=True)
    }
    return escrever_planilha(por_ano, saida, data_ref=data_ref, mes_ipca_ref=mes_ref)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--pasta-cache",
        type=Path,
        default=ROOT / "data" / "bndes_automaticas",
        help="Pasta dos Excels BNDES (download/cache)",
    )
    p.add_argument(
        "--saida",
        type=Path,
        default=ROOT / "output" / "indiretas_automaticas_ipca" / "INDIRETAS_AUTOMATICAS_IPCA_JUL2026.xlsx",
    )
    p.add_argument("--ipca", type=Path, default=None)
    p.add_argument(
        "--data-ref",
        type=str,
        default="2026-07-31",
        help="Data de referência do IPCA (YYYY-MM-DD)",
    )
    p.add_argument("--sem-baixar", action="store_true")
    p.add_argument(
        "--anos",
        type=str,
        default=None,
        help="Lista de anos (ex.: 2009,2010) — útil para testes",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    data_ref = datetime.strptime(args.data_ref, "%Y-%m-%d")
    anos = None
    if args.anos:
        anos = {int(x.strip()) for x in args.anos.split(",") if x.strip()}
    try:
        processar(
            args.pasta_cache,
            args.saida,
            ipca_path=args.ipca,
            data_ref=data_ref,
            baixar=not args.sem_baixar,
            anos=anos,
        )
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
