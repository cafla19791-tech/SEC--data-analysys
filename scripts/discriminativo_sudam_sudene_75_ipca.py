#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Discriminativo — benefício Sudam/Sudene de redução de 75% do IRPJ.

Fonte local (Receita Federal / consolidado do repositório):
  RENUNCIA FISCAL SUDAM-SUDENE (1).xlsx
  → data/sudam_sudene/renuncia_sudam_sudene.xlsx

Filtra: ``Sudam/Sudene - Redução 75% Projeto Setor Prioritário``.

Saídas:
  - Empresas beneficiadas (CNPJ, nome, totais corrente e IPCA)
  - Renúncias por ano (corrente e atualizado pelo IPCA até 31/07/2026)
  - Detalhe empresa × ano

Observação: o microdado por empresa disponível na fonte cobre anos-calendário
2015–2023. Anos 2003–2014 e 2024–2026 não constam com valor por CNPJ nesta base.

Uso::

  python scripts/discriminativo_sudam_sudene_75_ipca.py
"""

from __future__ import annotations

import argparse
import re
import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.calcular_diretas_ipca_selic import (  # noqa: E402
    _baixar_sgs,
    _idx_mes,
    carregar_ipca,
)

MARKER = "sudam-sudene-75-ipca-20260816a"
DATA_REF_IPCA = datetime(2026, 7, 31)
BENEFICIO_75 = "Sudam/Sudene - Redução 75% Projeto Setor Prioritário"

COL_ANO = "Ano-Calendário"
COL_CNPJ = "CNPJ"
COL_NOME = "Beneficiário"
COL_BENEF = "Benefício Fiscal"
COL_VALOR = "Valor Renunciado (R$)"
COL_IPCA = "Valor Renunciado atualizado pelo IPCA até 31/07/2026 (R$)"


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).replace("\n", " ").strip().lower())


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


def _resolver_fonte(explicit: Path | None) -> Path:
    candidatos = []
    if explicit is not None:
        candidatos.append(explicit)
    candidatos.extend(
        [
            ROOT / "data" / "sudam_sudene" / "renuncia_sudam_sudene.xlsx",
            ROOT / "RENUNCIA FISCAL SUDAM-SUDENE (1).xlsx",
            Path.cwd() / "RENUNCIA FISCAL SUDAM-SUDENE (1).xlsx",
        ]
    )
    for c in candidatos:
        if c is not None and c.exists():
            return c
    raise FileNotFoundError(
        "Arquivo de renúncia Sudam/Sudene não encontrado. "
        "Coloque RENUNCIA FISCAL SUDAM-SUDENE (1).xlsx na raiz ou em "
        "data/sudam_sudene/renuncia_sudam_sudene.xlsx"
    )


def carregar_renuncias_75(path: Path) -> pd.DataFrame:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        xl = pd.ExcelFile(path)
        # prefer sheet com nome conhecido
        sheet = None
        for s in xl.sheet_names:
            if "renúncia" in _norm(s) or "renuncia" in _norm(s):
                sheet = s
                break
        if sheet is None:
            sheet = xl.sheet_names[0]
        df = pd.read_excel(path, sheet_name=sheet)

    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    norms = {_norm(c): c for c in df.columns}

    def pick(*cands: str) -> str:
        for cand in cands:
            alvo = _norm(cand)
            if alvo in norms:
                return norms[alvo]
        for cand in cands:
            alvo = _norm(cand)
            for n, orig in norms.items():
                if alvo in n:
                    return orig
        raise ValueError(f"Coluna não encontrada ({cands}). Colunas: {list(df.columns)}")

    c_ano = pick("Ano-Calendário", "Ano Calendário", "Ano")
    c_cnpj = pick("CNPJ")
    c_nome = pick("Beneficiário", "Nome", "Empresa")
    c_benef = pick("Benefício Fiscal", "Beneficio Fiscal")
    c_valor = pick("Valor Renunciado(R$)", "Valor Renunciado (R$)", "Valor Renunciado")

    out = pd.DataFrame(
        {
            COL_ANO: pd.to_numeric(df[c_ano], errors="coerce"),
            COL_CNPJ: df[c_cnpj].astype(str).str.strip(),
            COL_NOME: df[c_nome].astype(str).str.strip(),
            COL_BENEF: df[c_benef].astype(str).str.strip(),
            COL_VALOR: pd.to_numeric(df[c_valor], errors="coerce").fillna(0.0),
        }
    )
    # extras opcionais
    for src, dest in [
        ("UF", "UF"),
        ("Município", "Município"),
        ("Nome Fantasia", "Nome Fantasia"),
    ]:
        if _norm(src) in norms:
            out[dest] = df[norms[_norm(src)]].astype(str).str.strip()

    out = out.dropna(subset=[COL_ANO]).copy()
    out[COL_ANO] = out[COL_ANO].astype(int)

    # filtro 75%
    m = out[COL_BENEF].str.contains("75%", na=False) | out[COL_BENEF].map(
        lambda x: _norm(x) == _norm(BENEFICIO_75)
    )
    out = out.loc[m].copy()
    out = out[out[COL_VALOR] != 0].copy()
    return out.reset_index(drop=True)


def aplicar_ipca(
    df: pd.DataFrame,
    ipca: pd.DataFrame,
    data_ref: datetime = DATA_REF_IPCA,
) -> pd.DataFrame:
    """Atualiza a renúncia do ano-calendário a partir de dez/ano até data_ref."""
    out = df.copy()
    ref = pd.Timestamp(data_ref).to_period("M").to_timestamp()
    i_ref = _idx_mes(ipca, ref)
    f_ref = float(ipca.loc[i_ref, "fator"])
    mes_ref = pd.Timestamp(ipca.loc[i_ref, "mes"])

    # data-base = dezembro do ano-calendário
    mes_base = pd.to_datetime(
        dict(year=out[COL_ANO], month=12, day=1)
    ).dt.to_period("M").dt.to_timestamp()

    tmp = pd.DataFrame({"_i": range(len(out)), "mes": mes_base}).sort_values("mes")
    merged = pd.merge_asof(
        tmp, ipca[["mes", "fator"]].sort_values("mes"), on="mes", direction="backward"
    ).sort_values("_i")
    f0 = merged["fator"].to_numpy(dtype=float)
    out[COL_IPCA] = (out[COL_VALOR].to_numpy(dtype=float) * (f_ref / f0)).round(2)
    out.attrs["mes_ipca_ref"] = mes_ref
    return out


def resumo_empresas(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby([COL_CNPJ, COL_NOME], sort=False)
        .agg(
            **{
                "Qtd anos": (COL_ANO, "nunique"),
                "Ano inicial": (COL_ANO, "min"),
                "Ano final": (COL_ANO, "max"),
                COL_VALOR: (COL_VALOR, "sum"),
                COL_IPCA: (COL_IPCA, "sum"),
            }
        )
        .reset_index()
    )
    g[COL_VALOR] = g[COL_VALOR].round(2)
    g[COL_IPCA] = g[COL_IPCA].round(2)
    g = g.sort_values([COL_IPCA, COL_NOME], ascending=[False, True], kind="mergesort")
    g.insert(0, "Ranking", range(1, len(g) + 1))
    return g.reset_index(drop=True)


def resumo_por_ano(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby(COL_ANO, sort=True)
        .agg(
            **{
                "Qtd empresas": (COL_CNPJ, "nunique"),
                "Qtd registros": (COL_CNPJ, "size"),
                COL_VALOR: (COL_VALOR, "sum"),
                COL_IPCA: (COL_IPCA, "sum"),
            }
        )
        .reset_index()
    )
    g[COL_VALOR] = g[COL_VALOR].round(2)
    g[COL_IPCA] = g[COL_IPCA].round(2)
    return g


def detalhe_empresa_ano(df: pd.DataFrame) -> pd.DataFrame:
    preferred = [
        COL_ANO,
        COL_CNPJ,
        COL_NOME,
        "Nome Fantasia",
        "UF",
        "Município",
        COL_BENEF,
        COL_VALOR,
        COL_IPCA,
    ]
    cols = [c for c in preferred if c in df.columns]
    return (
        df[cols]
        .sort_values([COL_ANO, COL_IPCA], ascending=[True, False], kind="mergesort")
        .reset_index(drop=True)
    )


def escrever_planilha(
    empresas: pd.DataFrame,
    por_ano: pd.DataFrame,
    detalhe: pd.DataFrame,
    saida: Path,
    *,
    meta: dict,
) -> Path:
    import xlsxwriter

    saida.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(saida))
    fmt_hdr = wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "white"})
    fmt_num = wb.add_format({"num_format": "#,##0.00"})
    fmt_bold = wb.add_format({"bold": True})

    def dump(name: str, df: pd.DataFrame, money_cols: set[str]):
        ws = wb.add_worksheet(name[:31])
        for j, c in enumerate(df.columns):
            ws.write(0, j, str(c), fmt_hdr)
        for i, row in enumerate(df.itertuples(index=False, name=None), start=1):
            for j, val in enumerate(row):
                col = df.columns[j]
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    continue
                if col in money_cols and isinstance(val, (int, float)):
                    ws.write_number(i, j, float(val), fmt_num)
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    ws.write_number(i, j, float(val))
                else:
                    ws.write(i, j, str(val))
        for j in range(len(df.columns)):
            ws.set_column(j, j, 18 if j else 14)
        if len(df.columns) > 2:
            ws.set_column(2, 2, 45)

    # Capa
    ws = wb.add_worksheet("Capa")
    capa = [
        ("Título", "Sudam/Sudene — Redução 75% do IRPJ (renúncia fiscal)"),
        ("Benefício", BENEFICIO_75),
        ("Pedido", "Empresas beneficiadas (2003–2026) e renúncias por ano (corrente e IPCA)"),
        ("Cobertura da fonte", meta["cobertura"]),
        ("Referência IPCA", meta["data_ref"]),
        ("Mês IPCA efetivo", meta["mes_ipca"]),
        ("Base da atualização", "Dezembro do ano-calendário → data de referência IPCA"),
        ("Empresas", meta["n_empresas"]),
        ("Registros", meta["n_registros"]),
        ("Total corrente (R$)", meta["total_corrente"]),
        ("Total IPCA 31/07/2026 (R$)", meta["total_ipca"]),
        ("Marker", MARKER),
        ("Gerado em", datetime.now().strftime("%Y-%m-%d %H:%M")),
        (
            "Nota",
            "Microdados por CNPJ na fonte disponível: 2015–2023. "
            "Anos 2003–2014 e 2024–2026 não possuem valor por empresa nesta base "
            "(transparência por beneficiário da RFB é recente).",
        ),
    ]
    ws.write(0, 0, "Campo", fmt_bold)
    ws.write(0, 1, "Valor", fmt_bold)
    for i, (k, v) in enumerate(capa, start=1):
        ws.write(i, 0, k)
        if isinstance(v, float):
            ws.write_number(i, 1, v, fmt_num)
        else:
            ws.write(i, 1, v)
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 90)

    dump("Empresas", empresas, {COL_VALOR, COL_IPCA})
    dump("Por_Ano", por_ano, {COL_VALOR, COL_IPCA})
    dump("Empresa_Ano", detalhe, {COL_VALOR, COL_IPCA})

    # total row on Por_Ano already as separate sheet ok
    wb.close()
    print(f"[OK] {saida} ({saida.stat().st_size / 1e6:.2f} MB)")
    return saida


def processar(
    *,
    fonte: Path | None,
    saida: Path,
    ipca_path: Path | None = None,
    data_ref: datetime = DATA_REF_IPCA,
    ano_ini: int = 2003,
    ano_fim: int = 2026,
) -> dict:
    print(f"[{MARKER}]")
    path = _resolver_fonte(fonte)
    print(f"[INFO] Fonte: {path}")
    df = carregar_renuncias_75(path)
    print(f"[INFO] Registros 75%: {len(df):,} | anos {df[COL_ANO].min()}–{df[COL_ANO].max()}")

    # janela pedida (intersecta com o disponível)
    df = df[(df[COL_ANO] >= ano_ini) & (df[COL_ANO] <= ano_fim)].copy()
    anos_disp = sorted(df[COL_ANO].unique())
    cobertura = (
        f"{anos_disp[0]}–{anos_disp[-1]}" if anos_disp else "sem dados na janela"
    )

    ipca = carregar_ipca_desde_2002(ipca_path)
    df = aplicar_ipca(df, ipca, data_ref=data_ref)
    mes_ref = df.attrs.get("mes_ipca_ref")

    empresas = resumo_empresas(df)
    por_ano = resumo_por_ano(df)
    # linha total
    if len(por_ano):
        tot = {
            COL_ANO: "TOTAL",
            "Qtd empresas": int(df[COL_CNPJ].nunique()),
            "Qtd registros": int(len(df)),
            COL_VALOR: round(float(df[COL_VALOR].sum()), 2),
            COL_IPCA: round(float(df[COL_IPCA].sum()), 2),
        }
        por_ano = pd.concat([por_ano, pd.DataFrame([tot])], ignore_index=True)

    detalhe = detalhe_empresa_ano(df)
    escrever_planilha(
        empresas,
        por_ano,
        detalhe,
        saida,
        meta={
            "cobertura": cobertura,
            "data_ref": pd.Timestamp(data_ref).strftime("%d/%m/%Y"),
            "mes_ipca": pd.Timestamp(mes_ref).strftime("%Y-%m"),
            "n_empresas": f"{df[COL_CNPJ].nunique():,}",
            "n_registros": f"{len(df):,}",
            "total_corrente": float(df[COL_VALOR].sum()),
            "total_ipca": float(df[COL_IPCA].sum()),
        },
    )
    return {
        "empresas": empresas,
        "por_ano": por_ano,
        "detalhe": detalhe,
        "cobertura": cobertura,
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--fonte", type=Path, default=None)
    p.add_argument(
        "--saida",
        type=Path,
        default=ROOT
        / "output"
        / "sudam_sudene"
        / "DISCRIMINATIVO_SUDAM_SUDENE_75_IPCA_JUL2026.xlsx",
    )
    p.add_argument("--ipca", type=Path, default=None)
    p.add_argument("--data-ref", type=str, default="2026-07-31")
    p.add_argument("--ano-ini", type=int, default=2003)
    p.add_argument("--ano-fim", type=int, default=2026)
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    data_ref = datetime.strptime(args.data_ref, "%Y-%m-%d")
    try:
        info = processar(
            fonte=args.fonte,
            saida=args.saida,
            ipca_path=args.ipca,
            data_ref=data_ref,
            ano_ini=args.ano_ini,
            ano_fim=args.ano_fim,
        )
        print(f"[INFO] Cobertura efetiva: {info['cobertura']}")
        print(info["por_ano"].to_string(index=False))
        print("\nTop 10 empresas (IPCA):")
        print(info["empresas"].head(10).to_string(index=False))
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
