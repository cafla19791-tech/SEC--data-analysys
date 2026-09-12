#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Top 100 empresas — operações DIRETAS BNDES (2003–2018).

Colunas:
  - CNPJ
  - Nome da empresa
  - Soma dos empréstimos atualizados pelo IPCA até 31/07/2026

Fonte: naoautomaticas.xlsx (Forma de Apoio = DIRETA).
Valor do empréstimo por contrato: valor desembolsado; se zero, valor contratado.
Cada contrato é atualizado pelo IPCA (mês da contratação → 31/07/2026) e depois
somado por CNPJ.

Uso::

  python scripts/top100_diretas_2003_2018_ipca.py
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

MARKER = "top100-diretas-2003-2018-20260816a"
URL_BNDES = (
    "https://www.bndes.gov.br/arquivos/central-downloads/"
    "operacoes_financiamento/naoautomaticas/naoautomaticas.xlsx"
)
DATA_REF_IPCA = datetime(2026, 7, 31)
ANO_INI, ANO_FIM = 2003, 2018

COL_CNPJ = "CNPJ"
COL_NOME = "Nome da empresa"
COL_SOMA = "Soma dos empréstimos atualizados pelo IPCA até 31/07/2026"


def _norm(s: str) -> str:
    s = str(s).replace("\n", " ").strip().lower()
    return re.sub(r"\s+", " ", s)


def baixar_fonte(destino: Path, url: str = URL_BNDES) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists() and destino.stat().st_size > 100_000:
        print(f"[CACHE] {destino.name} ({destino.stat().st_size / 1e6:.1f} MB)")
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


def _col(df: pd.DataFrame, *candidatos: str) -> str:
    norms = {_norm(c): c for c in df.columns}
    for cand in candidatos:
        alvo = _norm(cand)
        if alvo in norms:
            return norms[alvo]
    for cand in candidatos:
        alvo = _norm(cand)
        hits = [(n, o) for n, o in norms.items() if alvo in n]
        if hits:
            hits.sort(key=lambda x: len(x[0]))
            return hits[0][1]
    raise ValueError(f"Coluna não encontrada ({candidatos}). Colunas: {list(df.columns)}")


def carregar_diretas_periodo(
    path: Path,
    *,
    ano_ini: int = ANO_INI,
    ano_fim: int = ANO_FIM,
) -> pd.DataFrame:
    hdr = detectar_header(path)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(path, sheet_name="SITE", header=hdr)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

    c_cli = _col(df, "Cliente")
    c_cnpj = _col(df, "CNPJ")
    c_data = _col(df, "Data da contratação", "Data da contratacao")
    c_des = _col(df, "Valor desembolsado R$", "Valor desembolsado")
    c_con = _col(df, "Valor contratado  R$", "Valor contratado")
    c_forma = _col(df, "Forma de apoio")

    out = pd.DataFrame(
        {
            "cliente": df[c_cli].astype(str).str.strip(),
            "cnpj": df[c_cnpj].astype(str).str.strip(),
            "data": pd.to_datetime(df[c_data], errors="coerce"),
            "desembolsado": pd.to_numeric(df[c_des], errors="coerce").fillna(0.0),
            "contratado": pd.to_numeric(df[c_con], errors="coerce").fillna(0.0),
            "forma": df[c_forma].astype(str).str.strip().str.upper(),
        }
    )
    out = out.dropna(subset=["data"])
    out = out[out["forma"].str.startswith("DIRETA")]
    out = out[
        (out["data"] >= pd.Timestamp(year=ano_ini, month=1, day=1))
        & (out["data"] <= pd.Timestamp(year=ano_fim, month=12, day=31))
    ].copy()

    # valor do empréstimo: desembolsado; se zero, contratado
    out["emprestimo"] = out["desembolsado"].where(out["desembolsado"] > 0, out["contratado"])
    # CNPJ só dígitos (14 posições quando possível)
    out["cnpj"] = out["cnpj"].str.replace(r"\D", "", regex=True)
    out["cnpj"] = out["cnpj"].map(
        lambda x: x.zfill(14) if x.isdigit() and 1 <= len(x) <= 14 else x
    )
    out = out[out["cnpj"].str.len() > 0]
    out = out[out["emprestimo"] > 0]
    return out.reset_index(drop=True)


def carregar_ipca_desde_2002(path: Path | None = None) -> pd.DataFrame:
    if path is not None and path.exists():
        ipca = carregar_ipca(path)
        if ipca["mes"].min() <= pd.Timestamp("2002-01-01"):
            return ipca
    print("[INFO] Baixando IPCA Bacen SGS 433 desde 01/01/2002...")
    raw = _baixar_sgs(433, inicio="01/01/2002")
    raw = raw.sort_values("mes").drop_duplicates("mes").copy()
    raw["fator"] = (1.0 + raw["valor"] / 100.0).cumprod()
    return raw.reset_index(drop=True)


def atualizar_ipca(
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
            "mes": out["data"].dt.to_period("M").dt.to_timestamp(),
        }
    ).sort_values("mes")
    merged = pd.merge_asof(
        tmp, ipca[["mes", "fator"]].sort_values("mes"), on="mes", direction="backward"
    ).sort_values("_i")
    f0 = merged["fator"].to_numpy(dtype=float)
    out["emprestimo_ipca"] = (out["emprestimo"].to_numpy(dtype=float) * (f_ref / f0)).round(2)
    out.attrs["mes_ipca_ref"] = mes_ref
    return out


def top100_empresas(df: pd.DataFrame, n: int = 100) -> pd.DataFrame:
    """Agrega por CNPJ; nome = mais frequente (desempate: maior soma)."""

    def _nome(s: pd.Series) -> str:
        vc = s.value_counts()
        return str(vc.index[0]) if len(vc) else ""

    g = (
        df.groupby("cnpj", sort=False)
        .agg(
            nome=("cliente", _nome),
            soma_ipca=("emprestimo_ipca", "sum"),
            qtd=("emprestimo_ipca", "size"),
        )
        .reset_index()
    )
    g = g.sort_values(["soma_ipca", "cnpj"], ascending=[False, True], kind="mergesort")
    g = g.head(n).reset_index(drop=True)
    g.insert(0, "Ranking", range(1, len(g) + 1))
    out = g.rename(
        columns={
            "cnpj": COL_CNPJ,
            "nome": COL_NOME,
            "soma_ipca": COL_SOMA,
        }
    )
    out[COL_SOMA] = out[COL_SOMA].round(2)
    return out[[COL_CNPJ, COL_NOME, COL_SOMA]].copy()


def formatar_cnpj(cnpj: str) -> str:
    d = re.sub(r"\D", "", str(cnpj))
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return d


def escrever_excel(df: pd.DataFrame, saida: Path, *, meta: dict) -> Path:
    import xlsxwriter

    saida.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(saida))
    fmt_num = wb.add_format({"num_format": "#,##0.00"})
    fmt_hdr = wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "white"})
    fmt_bold = wb.add_format({"bold": True})

    ws = wb.add_worksheet("Top100")
    for j, c in enumerate(df.columns):
        ws.write(0, j, c, fmt_hdr)
    for i, row in enumerate(df.itertuples(index=False, name=None), start=1):
        cnpj, nome, soma = row
        ws.write(i, 0, formatar_cnpj(cnpj))
        ws.write(i, 1, nome)
        ws.write_number(i, 2, float(soma), fmt_num)
    ws.set_column(0, 0, 22)
    ws.set_column(1, 1, 55)
    ws.set_column(2, 2, 55)

    ws = wb.add_worksheet("Capa")
    capa = [
        ("Título", "Top 100 empresas — empréstimos BNDES (operações DIRETAS)"),
        ("Período", f"{meta['ano_ini']}–{meta['ano_fim']}"),
        ("Fonte", URL_BNDES),
        ("Critério", "Soma dos empréstimos atualizados pelo IPCA até 31/07/2026"),
        ("Valor base", "Desembolso; se zero, valor contratado"),
        ("Referência IPCA", meta["data_ref"]),
        ("Mês IPCA efetivo", meta["mes_ipca"]),
        ("Contratos considerados", meta["contratos"]),
        ("Empresas no ranking", len(df)),
        ("Marker", MARKER),
        ("Gerado em", datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    ws.write(0, 0, "Campo", fmt_bold)
    ws.write(0, 1, "Valor", fmt_bold)
    for i, (k, v) in enumerate(capa, start=1):
        ws.write(i, 0, k)
        ws.write(i, 1, v)
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 80)
    wb.close()
    print(f"[OK] {saida} ({saida.stat().st_size / 1e3:.1f} KB)")
    return saida


def processar(
    *,
    fonte: Path,
    saida: Path,
    ipca_path: Path | None = None,
    n: int = 100,
    ano_ini: int = ANO_INI,
    ano_fim: int = ANO_FIM,
    data_ref: datetime = DATA_REF_IPCA,
    baixar: bool = True,
) -> pd.DataFrame:
    print(f"[{MARKER}]")
    if baixar or not fonte.exists():
        baixar_fonte(fonte)
    contratos = carregar_diretas_periodo(fonte, ano_ini=ano_ini, ano_fim=ano_fim)
    print(f"[INFO] Contratos DIRETA {ano_ini}–{ano_fim}: {len(contratos):,}")
    ipca = carregar_ipca_desde_2002(ipca_path)
    contratos = atualizar_ipca(contratos, ipca, data_ref=data_ref)
    mes_ref = contratos.attrs.get("mes_ipca_ref")
    print(f"[INFO] IPCA efetivo: {pd.Timestamp(mes_ref):%Y-%m}")
    ranking = top100_empresas(contratos, n=n)
    escrever_excel(
        ranking,
        saida,
        meta={
            "ano_ini": ano_ini,
            "ano_fim": ano_fim,
            "data_ref": pd.Timestamp(data_ref).strftime("%d/%m/%Y"),
            "mes_ipca": pd.Timestamp(mes_ref).strftime("%Y-%m"),
            "contratos": f"{len(contratos):,}",
        },
    )
    return ranking


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
        / "top100_diretas"
        / "TOP100_DIRETAS_2003_2018_IPCA_JUL2026.xlsx",
    )
    p.add_argument("--ipca", type=Path, default=None)
    p.add_argument("--n", type=int, default=100)
    p.add_argument("--ano-ini", type=int, default=ANO_INI)
    p.add_argument("--ano-fim", type=int, default=ANO_FIM)
    p.add_argument("--data-ref", type=str, default="2026-07-31")
    p.add_argument("--sem-baixar", action="store_true")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    data_ref = datetime.strptime(args.data_ref, "%Y-%m-%d")
    try:
        ranking = processar(
            fonte=args.fonte,
            saida=args.saida,
            ipca_path=args.ipca,
            n=args.n,
            ano_ini=args.ano_ini,
            ano_fim=args.ano_fim,
            data_ref=data_ref,
            baixar=not args.sem_baixar,
        )
        print(ranking.head(10).to_string(index=False))
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
