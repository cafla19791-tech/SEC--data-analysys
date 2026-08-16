#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Discriminativo das operações NÃO AUTOMÁTICAS do BNDES.

Fonte:
  https://www.bndes.gov.br/arquivos/central-downloads/operacoes_financiamento/naoautomaticas/naoautomaticas.xlsx

Abas:
  1) Contratos de 2002
  2) 01/01/2003 a 11/05/2016
  3) 12/05/2016 a 31/12/2018
  4) 01/01/2019 a 31/12/2022
  5) 01/01/2023 até hoje
  6) Todos (01/01/2002–hoje) por cliente (ordem decrescente do total
     de empréstimos), com contratos e totais correntes + IPCA
  (+ Capa)

Nas abas 1–5: coluna de desembolso atualizado pelo IPCA até 31/07/2026;
ordenação por ano com totais anuais.

Uso::

  python scripts/discriminativo_naoautomaticas_ipca.py
"""

from __future__ import annotations

import argparse
import re
import sys
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

MARKER = "naoautomaticas-discriminativo-20260816a"
URL_BNDES = (
    "https://www.bndes.gov.br/arquivos/central-downloads/"
    "operacoes_financiamento/naoautomaticas/naoautomaticas.xlsx"
)
DATA_REF_IPCA = datetime(2026, 7, 31)
COL_IPCA = "Valor Desembolsado atualizado pelo IPCA até 31/07/2026"
COL_DATA = "Data da contratação"
COL_CLIENTE = "Cliente"
COL_DESEMBOLSO = "Valor desembolsado R$"
COL_CONTRATADO = "Valor contratado  R$"

PERIODOS = (
    {
        "aba": "1_Ano_2002",
        "titulo": "Contratos do ano 2002",
        "inicio": datetime(2002, 1, 1),
        "fim": datetime(2002, 12, 31),
    },
    {
        "aba": "2_2003_a_2016-05-11",
        "titulo": "Contratos 01/01/2003 a 11/05/2016",
        "inicio": datetime(2003, 1, 1),
        "fim": datetime(2016, 5, 11),
    },
    {
        "aba": "3_2016-05-12_a_2018",
        "titulo": "Contratos 12/05/2016 a 31/12/2018",
        "inicio": datetime(2016, 5, 12),
        "fim": datetime(2018, 12, 31),
    },
    {
        "aba": "4_2019_a_2022",
        "titulo": "Contratos 01/01/2019 a 31/12/2022",
        "inicio": datetime(2019, 1, 1),
        "fim": datetime(2022, 12, 31),
    },
    {
        "aba": "5_2023_ate_hoje",
        "titulo": "Contratos 01/01/2023 até hoje",
        "inicio": datetime(2023, 1, 1),
        "fim": None,  # até hoje
    },
)


def _norm(s: str) -> str:
    s = str(s).replace("\n", " ").strip().lower()
    return re.sub(r"\s+", " ", s)


def baixar_fonte(destino: Path, url: str = URL_BNDES) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists() and destino.stat().st_size > 100_000:
        print(f"[CACHE] {destino} ({destino.stat().st_size / 1e6:.1f} MB)")
        return destino
    print(f"[DOWNLOAD] {url}")
    r = requests.get(url, timeout=300)
    r.raise_for_status()
    destino.write_bytes(r.content)
    print(f"[OK] {destino} ({destino.stat().st_size / 1e6:.1f} MB)")
    return destino


def detectar_header(path: Path, sheet: str = "SITE", max_linhas: int = 20) -> int:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=max_linhas)
    for i in range(len(raw)):
        cells = [_norm(x) for x in raw.iloc[i].tolist() if pd.notna(x)]
        if "cliente" in cells and any("data da contrata" in c for c in cells):
            return i
    raise ValueError(f"Cabeçalho não encontrado em {path.name}")


def carregar_contratos(path: Path, hoje: datetime | None = None) -> pd.DataFrame:
    hdr = detectar_header(path)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(path, sheet_name="SITE", header=hdr)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    if COL_DATA not in df.columns:
        raise ValueError(f"Sem coluna '{COL_DATA}'. Colunas: {list(df.columns)}")
    if COL_DESEMBOLSO not in df.columns:
        # tolerar variação de espaços
        cand = [c for c in df.columns if "desembolso" in _norm(c)]
        if not cand:
            raise ValueError(f"Sem coluna de desembolso. Colunas: {list(df.columns)}")
        df = df.rename(columns={cand[0]: COL_DESEMBOLSO})

    out = df.copy()
    out[COL_DATA] = pd.to_datetime(out[COL_DATA], errors="coerce")
    out = out.dropna(subset=[COL_DATA]).copy()
    out[COL_DESEMBOLSO] = pd.to_numeric(out[COL_DESEMBOLSO], errors="coerce").fillna(0.0)
    if COL_CONTRATADO in out.columns:
        out[COL_CONTRATADO] = pd.to_numeric(out[COL_CONTRATADO], errors="coerce").fillna(0.0)
    out[COL_CLIENTE] = out[COL_CLIENTE].astype(str).str.strip()
    out["_ano"] = out[COL_DATA].dt.year.astype(int)

    limite = pd.Timestamp(hoje or datetime.now()).normalize()
    out = out[out[COL_DATA] <= limite].copy()
    # janela mínima pedida: a partir de 2002
    out = out[out[COL_DATA] >= pd.Timestamp("2002-01-01")].copy()
    return out.reset_index(drop=True)


def aplicar_ipca(
    df: pd.DataFrame,
    ipca: pd.DataFrame,
    data_ref: datetime = DATA_REF_IPCA,
) -> pd.DataFrame:
    out = df.copy()
    ref = pd.Timestamp(data_ref).to_period("M").to_timestamp()
    i_ref = _idx_mes(ipca, ref)
    f_ref = float(ipca.loc[i_ref, "fator"])
    mes_ref = pd.Timestamp(ipca.loc[i_ref, "mes"])

    tmp = pd.DataFrame(
        {
            "_i": range(len(out)),
            "mes": out[COL_DATA].dt.to_period("M").dt.to_timestamp(),
        }
    ).sort_values("mes")
    merged = pd.merge_asof(
        tmp, ipca[["mes", "fator"]].sort_values("mes"), on="mes", direction="backward"
    ).sort_values("_i")
    f0 = merged["fator"].to_numpy(dtype=float)
    valor = out[COL_DESEMBOLSO].to_numpy(dtype=float)
    fator = f_ref / f0
    out[COL_IPCA] = pd.Series(valor * fator, index=out.index).round(2)
    out.attrs["mes_ipca_ref"] = mes_ref
    return out


def filtrar_periodo(
    df: pd.DataFrame,
    inicio: datetime,
    fim: datetime | None,
    hoje: datetime,
) -> pd.DataFrame:
    ini = pd.Timestamp(inicio)
    fim_ts = pd.Timestamp(fim) if fim is not None else pd.Timestamp(hoje).normalize()
    m = (df[COL_DATA] >= ini) & (df[COL_DATA] <= fim_ts)
    return df.loc[m].copy()


def _totais_por_ano(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby("_ano", sort=True)
        .agg(
            **{
                "Qtd Contratos": (COL_CLIENTE, "size"),
                COL_DESEMBOLSO: (COL_DESEMBOLSO, "sum"),
                COL_IPCA: (COL_IPCA, "sum"),
            }
        )
        .reset_index()
        .rename(columns={"_ano": "Ano"})
    )
    g[COL_DESEMBOLSO] = g[COL_DESEMBOLSO].round(2)
    g[COL_IPCA] = g[COL_IPCA].round(2)
    return g


def montar_aba_periodo(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Ordena por ano/data e devolve (detalhe com colunas úteis, totais anuais)."""
    detalhe = df.sort_values(["_ano", COL_DATA, COL_CLIENTE], kind="mergesort").copy()
    # coloca coluna IPCA no final (já existe)
    cols = [c for c in detalhe.columns if c not in (COL_IPCA, "_ano")] + [COL_IPCA]
    detalhe = detalhe[cols]
    totais = _totais_por_ano(df)
    return detalhe, totais


def montar_aba_por_cliente(df: pd.DataFrame) -> pd.DataFrame:
    """
    Todos os contratos, agrupados por cliente.
    Clientes em ordem decrescente do total de empréstimos (desembolso corrente;
    se zero, usa valor contratado).
    """
    work = df.copy()
    if COL_CONTRATADO in work.columns:
        base_emp = work[COL_DESEMBOLSO].where(work[COL_DESEMBOLSO] > 0, work[COL_CONTRATADO])
    else:
        base_emp = work[COL_DESEMBOLSO]
    work["_emp"] = base_emp

    totais_cli = (
        work.groupby(COL_CLIENTE, sort=False)
        .agg(
            total_emprestimos=("_emp", "sum"),
            total_corrente=(COL_DESEMBOLSO, "sum"),
            total_ipca=(COL_IPCA, "sum"),
            qtd=(COL_CLIENTE, "size"),
        )
        .reset_index()
    )
    totais_cli = totais_cli.sort_values(
        ["total_emprestimos", COL_CLIENTE], ascending=[False, True], kind="mergesort"
    )

    cols_base = [
        c
        for c in [
            COL_CLIENTE,
            "CNPJ",
            "Número do contrato",
            COL_DATA,
            COL_CONTRATADO,
            COL_DESEMBOLSO,
            COL_IPCA,
            "UF",
            "Situação do contrato",
            "Forma de apoio",
            "Produto",
        ]
        if c in work.columns or c == COL_IPCA
    ]
    # garantir IPCA
    if COL_IPCA not in cols_base:
        cols_base.append(COL_IPCA)

    blocos: list[pd.DataFrame] = []
    for _, tr in totais_cli.iterrows():
        cli = tr[COL_CLIENTE]
        sub = work[work[COL_CLIENTE] == cli].sort_values(COL_DATA, kind="mergesort")
        detalhe = sub[cols_base].copy()
        blocos.append(detalhe)
        tot_row = {c: "" for c in cols_base}
        tot_row[COL_CLIENTE] = f"TOTAL — {cli}"
        if "CNPJ" in tot_row:
            tot_row["CNPJ"] = int(tr["qtd"])
        tot_row[COL_DESEMBOLSO] = round(float(tr["total_corrente"]), 2)
        tot_row[COL_IPCA] = round(float(tr["total_ipca"]), 2)
        if COL_CONTRATADO in tot_row:
            tot_row[COL_CONTRATADO] = round(float(tr["total_emprestimos"]), 2)
        blocos.append(pd.DataFrame([tot_row]))
        # linha em branco
        blocos.append(pd.DataFrame([{c: "" for c in cols_base}]))

    if not blocos:
        return pd.DataFrame(columns=cols_base)
    return pd.concat(blocos, ignore_index=True)


def _escrever_df(ws, df: pd.DataFrame, fmt_data, fmt_num, start_row: int = 0) -> int:
    cols = list(df.columns)
    for j, c in enumerate(cols):
        ws.write(start_row, j, str(c))
    r = start_row + 1
    for tup in df.itertuples(index=False, name=None):
        for j, val in enumerate(tup):
            if val is None or val == "":
                continue
            if isinstance(val, pd.Timestamp):
                if pd.isna(val):
                    continue
                ws.write_datetime(r, j, val.to_pydatetime(), fmt_data)
                continue
            try:
                if pd.isna(val):
                    continue
            except (TypeError, ValueError):
                pass
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                # datas excel serial already handled; money-ish floats
                if isinstance(val, float) or (
                    isinstance(val, int) and cols[j] in (COL_DESEMBOLSO, COL_IPCA, COL_CONTRATADO)
                ):
                    ws.write_number(r, j, float(val), fmt_num)
                else:
                    ws.write(r, j, val)
            else:
                ws.write(r, j, str(val))
        r += 1
    return r


def escrever_planilha(
    df: pd.DataFrame,
    saida: Path,
    *,
    hoje: datetime,
    data_ref: datetime,
    mes_ipca_ref: pd.Timestamp | None,
) -> Path:
    import xlsxwriter

    saida.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(
        str(saida),
        {"constant_memory": False, "strings_to_urls": False, "nan_inf_to_errors": True},
    )
    fmt_data = wb.add_format({"num_format": "dd/mm/yyyy"})
    fmt_num = wb.add_format({"num_format": "#,##0.00"})
    fmt_titulo = wb.add_format({"bold": True, "font_size": 12})
    fmt_sec = wb.add_format({"bold": True})

    # Capa
    ws = wb.add_worksheet("Capa")
    capa = [
        ("Título", "Discriminativo — Operações Não Automáticas BNDES"),
        ("Fonte", URL_BNDES),
        ("Referência IPCA", pd.Timestamp(data_ref).strftime("%d/%m/%Y")),
        (
            "Mês IPCA efetivo",
            pd.Timestamp(mes_ipca_ref).strftime("%Y-%m") if mes_ipca_ref is not None else "",
        ),
        ("Hoje (corte aba 5/6)", pd.Timestamp(hoje).strftime("%d/%m/%Y")),
        ("Marker", MARKER),
        ("Gerado em", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Total contratos", f"{len(df):,}".replace(",", ".")),
        (
            "Abas",
            "1–5 períodos + IPCA e totais anuais; "
            "6 todos por cliente (ordem decrescente do total de empréstimos)",
        ),
    ]
    ws.write(0, 0, "Campo", fmt_sec)
    ws.write(0, 1, "Valor", fmt_sec)
    for i, (k, v) in enumerate(capa, start=1):
        ws.write(i, 0, k)
        ws.write(i, 1, v)

    resumo_periodos = []
    for per in PERIODOS:
        parte = filtrar_periodo(df, per["inicio"], per["fim"], hoje)
        detalhe, totais_ano = montar_aba_periodo(parte)
        print(f"[ABA] {per['aba']}: {len(parte):,} contratos")
        sys.stdout.flush()
        ws = wb.add_worksheet(per["aba"][:31])
        ws.write(0, 0, per["titulo"], fmt_titulo)
        ws.write(1, 0, f"Contratos: {len(parte):,} | Ordenado por ano | Totais anuais abaixo")
        _escrever_df(ws, detalhe, fmt_data, fmt_num, start_row=3)

        # totais anuais após o detalhe
        start = 3 + len(detalhe) + 2
        ws.write(start, 0, "TOTAIS POR ANO", fmt_sec)
        # cabeçalho totais
        headers_t = ["Ano", "Qtd Contratos", COL_DESEMBOLSO, COL_IPCA]
        for j, h in enumerate(headers_t):
            ws.write(start + 1, j, h, fmt_sec)
        for i, row in enumerate(totais_ano.itertuples(index=False, name=None), start=start + 2):
            ano, qtd, v, vipca = row
            ws.write_number(i, 0, int(ano))
            ws.write_number(i, 1, int(qtd))
            ws.write_number(i, 2, float(v), fmt_num)
            ws.write_number(i, 3, float(vipca), fmt_num)
        # total geral do período
        tg = start + 2 + len(totais_ano)
        ws.write(tg, 0, "TOTAL PERÍODO", fmt_sec)
        ws.write_number(tg, 1, int(totais_ano["Qtd Contratos"].sum()) if len(totais_ano) else 0)
        ws.write_number(
            tg, 2, float(totais_ano[COL_DESEMBOLSO].sum()) if len(totais_ano) else 0.0, fmt_num
        )
        ws.write_number(
            tg, 3, float(totais_ano[COL_IPCA].sum()) if len(totais_ano) else 0.0, fmt_num
        )

        resumo_periodos.append(
            {
                "aba": per["aba"],
                "titulo": per["titulo"],
                "contratos": len(parte),
                "desembolso": float(parte[COL_DESEMBOLSO].sum()),
                "ipca": float(parte[COL_IPCA].sum()),
            }
        )

    # Aba 6 — por cliente
    print("[ABA] 6_Por_Cliente_2002_hoje ...")
    sys.stdout.flush()
    por_cli = montar_aba_por_cliente(df)
    ws = wb.add_worksheet("6_Por_Cliente_2002_hoje"[:31])
    ws.write(0, 0, "Todos os contratos (01/01/2002 até hoje) por cliente", fmt_titulo)
    ws.write(
        1,
        0,
        "Clientes em ordem decrescente do total de empréstimos; "
        "sob cada cliente: contratos + TOTAL (corrente e IPCA)",
    )
    _escrever_df(ws, por_cli, fmt_data, fmt_num, start_row=3)

    # Resumo
    ws = wb.add_worksheet("Resumo")
    ws.write(0, 0, "Aba", fmt_sec)
    ws.write(0, 1, "Título", fmt_sec)
    ws.write(0, 2, "Qtd Contratos", fmt_sec)
    ws.write(0, 3, COL_DESEMBOLSO, fmt_sec)
    ws.write(0, 4, COL_IPCA, fmt_sec)
    for i, r in enumerate(resumo_periodos, start=1):
        ws.write(i, 0, r["aba"])
        ws.write(i, 1, r["titulo"])
        ws.write_number(i, 2, r["contratos"])
        ws.write_number(i, 3, r["desembolso"], fmt_num)
        ws.write_number(i, 4, r["ipca"], fmt_num)
    i = len(resumo_periodos) + 1
    ws.write(i, 0, "6_Por_Cliente_2002_hoje")
    ws.write(i, 1, "Todos por cliente")
    ws.write_number(i, 2, len(df))
    ws.write_number(i, 3, float(df[COL_DESEMBOLSO].sum()), fmt_num)
    ws.write_number(i, 4, float(df[COL_IPCA].sum()), fmt_num)

    wb.close()
    print(f"[OK] {saida} ({saida.stat().st_size / 1e6:.1f} MB)")
    return saida


def carregar_ipca_desde_2002(path: Path | None = None) -> pd.DataFrame:
    if path is not None and path.exists():
        ipca = carregar_ipca(path)
        if ipca["mes"].min() <= pd.Timestamp("2002-01-01"):
            return ipca
        print("[AVISO] IPCA local incompleto; baixando Bacen...")
    print("[INFO] Baixando IPCA Bacen SGS 433 desde 01/01/2002...")
    raw = _baixar_sgs(433, inicio="01/01/2002")
    raw = raw.sort_values("mes").drop_duplicates("mes").copy()
    raw["fator"] = (1.0 + raw["valor"] / 100.0).cumprod()
    return raw.reset_index(drop=True)


def processar(
    *,
    fonte: Path,
    saida: Path,
    ipca_path: Path | None = None,
    data_ref: datetime = DATA_REF_IPCA,
    hoje: datetime | None = None,
    baixar: bool = True,
) -> Path:
    print(f"[{MARKER}]")
    hoje = hoje or datetime.now()
    if baixar or not fonte.exists():
        baixar_fonte(fonte)
    df = carregar_contratos(fonte, hoje=hoje)
    print(f"[INFO] Contratos carregados: {len(df):,}")
    ipca = carregar_ipca_desde_2002(ipca_path)
    df = aplicar_ipca(df, ipca, data_ref=data_ref)
    mes_ref = df.attrs.get("mes_ipca_ref")
    print(f"[INFO] IPCA ref={pd.Timestamp(data_ref):%Y-%m-%d} efetivo={mes_ref:%Y-%m}")
    return escrever_planilha(
        df, saida, hoje=hoje, data_ref=data_ref, mes_ipca_ref=mes_ref
    )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--fonte",
        type=Path,
        default=ROOT / "data" / "bndes_naoautomaticas" / "naoautomaticas.xlsx",
    )
    p.add_argument(
        "--saida",
        type=Path,
        default=ROOT
        / "output"
        / "naoautomaticas_discriminativo"
        / "DISCRIMINATIVO_NAOAUTOMATICAS_IPCA_JUL2026.xlsx",
    )
    p.add_argument("--ipca", type=Path, default=None)
    p.add_argument("--data-ref", type=str, default="2026-07-31")
    p.add_argument("--hoje", type=str, default=None, help="YYYY-MM-DD (corte abas 5/6)")
    p.add_argument("--sem-baixar", action="store_true")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    data_ref = datetime.strptime(args.data_ref, "%Y-%m-%d")
    hoje = datetime.strptime(args.hoje, "%Y-%m-%d") if args.hoje else datetime.now()
    try:
        processar(
            fonte=args.fonte,
            saida=args.saida,
            ipca_path=args.ipca,
            data_ref=data_ref,
            hoje=hoje,
            baixar=not args.sem_baixar,
        )
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
