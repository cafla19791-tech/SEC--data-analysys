#!/usr/bin/env python3
"""Atualiza a base BNDES de desembolsos 1995–2001 pelo IPCA e resume por ano.

Fonte oficial:
https://www.bndes.gov.br/wps/wcm/connect/site/d3f78a56-0e0a-4c54-a1fa-c8a06ebc23ec/BASE+DE+DADOS+DESEMBOLSO_1995+A+2001.xlsx

Cada linha (agregado mensal da base) é atualizada por:
  valor_atualizado = desembolso × fator_IPCA(ref) / fator_IPCA(mês do desembolso)

Saídas Excel:
  - Resumo_Anual: subtotais por ano (corrente e IPCA) + linha TOTAL 1995–2001
  - Total_Periodo: total do período
  - Por_Ano_Forma: opcional (DIRETA/INDIRETA)
  - Detalhe: linhas atualizadas (omitir com --sem-detalhe)
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.calcular_diretas_ipca_selic import (  # noqa: E402
    DATA_REF_DEFAULT,
    IPCA_COD,
    _baixar_sgs,
    carregar_ipca,
    fator_ipca_entre,
)

URL_BASE = (
    "https://www.bndes.gov.br/wps/wcm/connect/site/"
    "d3f78a56-0e0a-4c54-a1fa-c8a06ebc23ec/"
    "BASE+DE+DADOS+DESEMBOLSO_1995+A+2001.xlsx"
)
SHEET = "DESEMBOLSOS_BASE DE DADOS"
ANOS = list(range(1995, 2002))

MESES_PT = {
    "JANEIRO": 1,
    "FEVEREIRO": 2,
    "MARÇO": 3,
    "MARCO": 3,
    "ABRIL": 4,
    "MAIO": 5,
    "JUNHO": 6,
    "JULHO": 7,
    "AGOSTO": 8,
    "SETEMBRO": 9,
    "OUTUBRO": 10,
    "NOVEMBRO": 11,
    "DEZEMBRO": 12,
}


def baixar_base(destino: Path, url: str = URL_BASE) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists() and destino.stat().st_size > 1_000_000:
        print(f"[INFO] Base já em cache: {destino}")
        return destino
    print("[INFO] Baixando base BNDES 1995-2001...")
    r = requests.get(url, timeout=180)
    r.raise_for_status()
    destino.write_bytes(r.content)
    print(f"[OK] Salvo: {destino} ({destino.stat().st_size / 1e6:.1f} MB)")
    return destino


def _col_desembolso(columns: list) -> str:
    for c in columns:
        s = str(c).replace("\n", " ").strip().upper()
        if "DESEMBOLSO" in s and "R$" in s:
            return c
        if s in {"DESEMBOLSOS R$", "DESEMBOLSO R$", "VALOR"}:
            return c
    raise ValueError(f"Coluna de desembolso não encontrada. Colunas: {columns}")


def _mes_para_num(serie: pd.Series) -> pd.Series:
    """Converte MÊS (nome PT ou número) → 1..12."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce").fillna(1).astype(int).clip(1, 12)

    def one(v) -> int:
        if pd.isna(v):
            return 1
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            n = int(v)
            return n if 1 <= n <= 12 else 1
        s = str(v).strip().upper()
        if s.isdigit():
            n = int(s)
            return n if 1 <= n <= 12 else 1
        return MESES_PT.get(s, 1)

    return serie.map(one).astype(int)


def carregar_desembolsos(path: Path) -> pd.DataFrame:
    """Carrega a aba oficial (header na linha 2)."""
    df = pd.read_excel(path, sheet_name=SHEET, header=2)
    df.columns = [c if isinstance(c, str) else str(c) for c in df.columns]
    if "ANO" not in df.columns or "MÊS" not in df.columns:
        raise ValueError(f"Colunas ANO/MÊS ausentes. Disponíveis: {list(df.columns)}")

    col_val = _col_desembolso(list(df.columns))
    out = df.copy()
    out["ANO"] = pd.to_numeric(out["ANO"], errors="coerce")
    out = out[out["ANO"].isin(ANOS)].copy()
    out["ANO"] = out["ANO"].astype(int)
    out["mes_num"] = _mes_para_num(out["MÊS"])
    out["desembolso"] = pd.to_numeric(out[col_val], errors="coerce").fillna(0.0)
    out["data_desembolso"] = pd.to_datetime(
        dict(year=out["ANO"], month=out["mes_num"], day=1)
    )
    return out.reset_index(drop=True)


def carregar_ipca_desde_1995(path: Optional[Path] = None) -> pd.DataFrame:
    """IPCA com fator acumulado cobrindo 1995→data atual."""
    if path is not None and Path(path).exists():
        ipca = carregar_ipca(Path(path))
        if ipca["mes"].min() <= pd.Timestamp("1995-01-01"):
            return ipca
        print("[AVISO] IPCA local começa depois de 1995; completando via Bacen SGS 433...")

    print("[INFO] Baixando IPCA (Bacen SGS 433) desde 01/01/1995...")
    raw = _baixar_sgs(IPCA_COD, inicio="01/01/1995")
    df = raw.sort_values("mes").drop_duplicates("mes").copy()
    med = float(df["valor"].median())
    if med > 50:
        raise ValueError("IPCA não parece variação % a.m.")
    df["fator"] = (1.0 + df["valor"] / 100.0).cumprod()
    return df.reset_index(drop=True)


def atualizar_ipca(
    df: pd.DataFrame,
    ipca: pd.DataFrame,
    data_ref: datetime | pd.Timestamp | None = None,
) -> pd.DataFrame:
    """Atualiza cada linha pelo IPCA até data_ref (vetorizado via merge)."""
    ref = pd.Timestamp(data_ref or DATA_REF_DEFAULT)
    out = df.copy()

    ipca_idx = ipca.copy()
    ipca_idx["periodo"] = ipca_idx["mes"].dt.to_period("M")
    fator_por_mes = (
        ipca_idx.drop_duplicates("periodo")
        .set_index("periodo")["fator"]
        .sort_index()
        .astype(float)
    )

    per_ref = pd.Period(ref, freq="M")
    if per_ref not in fator_por_mes.index:
        ant = fator_por_mes.index[fator_por_mes.index <= per_ref]
        if len(ant) == 0:
            raise ValueError(f"IPCA sem observações até {ref.date()}")
        per_ref = ant.max()
    fator_ref = float(fator_por_mes.loc[per_ref])

    periodos = out["data_desembolso"].dt.to_period("M")
    # reindex com ffill para meses faltantes (usa último fator ≤ mês)
    full_idx = pd.period_range(fator_por_mes.index.min(), max(periodos.max(), per_ref), freq="M")
    fator_ffill = fator_por_mes.reindex(full_idx).ffill()
    fator_ini = periodos.map(fator_ffill)
    fator = fator_ref / fator_ini.replace(0, pd.NA)

    out["fator_IPCA"] = pd.to_numeric(fator, errors="coerce")
    out["valor_corrente"] = out["desembolso"]
    out["desembolso_ipca"] = out["desembolso"] * out["fator_IPCA"]
    out["valor_atualizado_IPCA"] = out["desembolso_ipca"]
    out["data_referencia_IPCA"] = str(ref.date())
    return out


def resumo_anual(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby("ANO", as_index=False)
        .agg(
            n_linhas=("valor_corrente", "size"),
            valor_corrente=("valor_corrente", "sum"),
            valor_atualizado_IPCA=("valor_atualizado_IPCA", "sum"),
        )
        .sort_values("ANO")
    )
    g["variacao_IPCA"] = g["valor_atualizado_IPCA"] - g["valor_corrente"]
    g["fator_medio_implicito"] = g["valor_atualizado_IPCA"] / g["valor_corrente"].replace(
        0, pd.NA
    )
    return g.reset_index(drop=True)


def resumo_por_forma(df: pd.DataFrame) -> pd.DataFrame:
    if "FORMA DE APOIO" not in df.columns:
        return pd.DataFrame()
    return (
        df.groupby(["ANO", "FORMA DE APOIO"], as_index=False)
        .agg(
            valor_corrente=("valor_corrente", "sum"),
            valor_atualizado_IPCA=("valor_atualizado_IPCA", "sum"),
        )
        .sort_values(["ANO", "FORMA DE APOIO"])
        .reset_index(drop=True)
    )


def total_periodo(resumo: pd.DataFrame, data_ref: pd.Timestamp) -> pd.DataFrame:
    corrente = float(resumo["valor_corrente"].sum())
    atualizado = float(resumo["valor_atualizado_IPCA"].sum())
    return pd.DataFrame(
        [
            {
                "periodo": "1995-2001",
                "n_linhas": int(resumo["n_linhas"].sum()),
                "valor_corrente": corrente,
                "valor_atualizado_IPCA": atualizado,
                "variacao_IPCA": atualizado - corrente,
                "fator_medio_implicito": (atualizado / corrente) if corrente else None,
                "data_referencia_IPCA": str(pd.Timestamp(data_ref).date()),
                "fonte_IPCA": "Bacen SGS 433 / IPCA_MENSAL.xlsx",
            }
        ]
    )


def _fmt_bi(x: float) -> str:
    """Formata em bilhões com vírgula decimal (ex.: R$ 119,97 bilhões)."""
    return f"R$ {x / 1e9:,.2f} bilhões".replace(",", "X").replace(".", ",").replace("X", ".")


def escrever_excel(
    detalhe: pd.DataFrame,
    resumo: pd.DataFrame,
    total: pd.DataFrame,
    por_forma: pd.DataFrame,
    saida: Path,
    incluir_detalhe: bool,
) -> None:
    saida.parent.mkdir(parents=True, exist_ok=True)
    tabela = resumo.copy()
    tabela.insert(0, "tipo", "Subtotal ano")
    linha_total = {
        "tipo": "TOTAL 1995-2001",
        "ANO": "1995-2001",
        "n_linhas": int(total.iloc[0]["n_linhas"]),
        "valor_corrente": float(total.iloc[0]["valor_corrente"]),
        "valor_atualizado_IPCA": float(total.iloc[0]["valor_atualizado_IPCA"]),
        "variacao_IPCA": float(total.iloc[0]["variacao_IPCA"]),
        "fator_medio_implicito": float(total.iloc[0]["fator_medio_implicito"] or 0),
    }
    tabela = pd.concat([tabela, pd.DataFrame([linha_total])], ignore_index=True)

    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        tabela.to_excel(writer, sheet_name="Resumo_Anual", index=False)
        total.to_excel(writer, sheet_name="Total_Periodo", index=False)
        if not por_forma.empty:
            por_forma.to_excel(writer, sheet_name="Por_Ano_Forma", index=False)
        if incluir_detalhe:
            cols = [
                c
                for c in [
                    "ANO",
                    "MÊS",
                    "mes_num",
                    "FORMA DE APOIO",
                    "PRODUTO",
                    "PORTE DE EMPRESA",
                    "REGIÃO",
                    "UF",
                    "MUNICÍPIO",
                    "SETOR CNAE",
                    "SETOR BNDES",
                    "valor_corrente",
                    "fator_IPCA",
                    "valor_atualizado_IPCA",
                    "desembolso_ipca",
                    "data_referencia_IPCA",
                ]
                if c in detalhe.columns
            ]
            detalhe[cols].to_excel(writer, sheet_name="Detalhe", index=False)


def escrever_markdown(resumo: pd.DataFrame, total: pd.DataFrame, md_path: Path) -> None:
    md_path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# Desembolsos BNDES 1995–2001 atualizados pelo IPCA",
        "",
        f"**Data de referência IPCA:** {total.iloc[0]['data_referencia_IPCA']}",
        "",
        "## Subtotais por ano",
        "",
        "| Ano | Linhas | Valor corrente | Valor atualizado IPCA | Variação |",
        "|-----|--------|----------------|----------------------|----------|",
    ]
    for _, r in resumo.iterrows():
        lines.append(
            f"| {int(r['ANO'])} | {int(r['n_linhas']):,} | {_fmt_bi(r['valor_corrente'])} | "
            f"{_fmt_bi(r['valor_atualizado_IPCA'])} | {_fmt_bi(r['variacao_IPCA'])} |"
        )
    t = total.iloc[0]
    lines += [
        "",
        "## Total do período 1995–2001",
        "",
        f"- **Valor corrente:** {_fmt_bi(float(t['valor_corrente']))}",
        f"- **Valor atualizado pelo IPCA:** {_fmt_bi(float(t['valor_atualizado_IPCA']))}",
        f"- **Variação (IPCA):** {_fmt_bi(float(t['variacao_IPCA']))}",
        f"- **Fator médio implícito:** {float(t['fator_medio_implicito']):.4f}",
        f"- **Linhas:** {int(t['n_linhas']):,}",
        "",
        "Fórmula por linha: `desembolso × fator_IPCA(ref) / fator_IPCA(mês)`.",
        "",
    ]
    md_path.write_text("\n".join(lines), encoding="utf-8")


def main(argv: Optional[list[str]] = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--excel",
        "--base",
        dest="excel",
        type=Path,
        default=ROOT / "data" / "bndes" / "BASE_DESEMBOLSO_1995_2001.xlsx",
        help="Arquivo da base BNDES",
    )
    p.add_argument("--ipca", type=Path, default=None, help="IPCA_MENSAL.xlsx (opcional)")
    p.add_argument(
        "--data-ref",
        default=str(pd.Timestamp(DATA_REF_DEFAULT).date()),
        help="Data de referência IPCA (YYYY-MM-DD)",
    )
    p.add_argument(
        "--saida",
        type=Path,
        default=ROOT / "output" / "desembolsos_1995_2001_ipca.xlsx",
    )
    p.add_argument(
        "--relatorio",
        type=Path,
        default=ROOT / "output" / "desembolsos_1995_2001_ipca.md",
    )
    p.add_argument("--baixar", action="store_true", help="Baixa a base se não existir")
    p.add_argument("--nao-baixar", action="store_true", help="Não baixa a base se faltar")
    p.add_argument(
        "--sem-detalhe",
        action="store_true",
        help="Não grava a aba Detalhe (arquivo bem menor)",
    )
    args = p.parse_args(argv)

    if not args.excel.exists():
        if args.nao_baixar:
            raise SystemExit(f"Base não encontrada: {args.excel}")
        if args.baixar or not args.nao_baixar:
            baixar_base(args.excel)
        else:
            raise SystemExit(f"Base não encontrada: {args.excel}")

    data_ref = pd.Timestamp(args.data_ref)
    print(f"[INFO] Carregando base: {args.excel}")
    df = carregar_desembolsos(args.excel)
    print(f"[INFO] Linhas 1995-2001: {len(df):,}")

    ipca = carregar_ipca_desde_1995(args.ipca)
    print(f"[INFO] IPCA: {ipca['mes'].min().date()} → {ipca['mes'].max().date()}")

    # sanity check com helper unitário
    _ = fator_ipca_entre(ipca, pd.Timestamp("1995-01-01"), data_ref)

    print("[INFO] Atualizando cada linha pelo IPCA...")
    detalhe = atualizar_ipca(df, ipca, data_ref)
    n_ok = int(detalhe["valor_atualizado_IPCA"].notna().sum())
    print(f"[INFO] Linhas atualizadas: {n_ok:,} / {len(detalhe):,}")

    resumo = resumo_anual(detalhe)
    total = total_periodo(resumo, data_ref)
    por_forma = resumo_por_forma(detalhe)

    escrever_excel(
        detalhe, resumo, total, por_forma, args.saida, incluir_detalhe=not args.sem_detalhe
    )
    escrever_markdown(resumo, total, args.relatorio)

    print()
    print("=== SUBTOTAIS POR ANO ===")
    print(resumo.to_string(index=False, float_format=lambda x: f"{x:,.2f}"))
    print()
    print("=== TOTAL 1995-2001 ===")
    print(f"  Valor corrente:           {_fmt_bi(float(total.iloc[0]['valor_corrente']))}")
    print(f"  Valor atualizado (IPCA):  {_fmt_bi(float(total.iloc[0]['valor_atualizado_IPCA']))}")
    print(f"  Variação IPCA:            {_fmt_bi(float(total.iloc[0]['variacao_IPCA']))}")
    print()
    print(f"[OK] Excel: {args.saida}")
    print(f"[OK] Relatório: {args.relatorio}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
