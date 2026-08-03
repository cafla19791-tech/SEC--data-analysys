#!/usr/bin/env python3
"""
Resumo de fluxos ContAgil — por contrato e por ano.

Lê o CSV/XLSX de parcelas gerado pelo ContAgil (ex.: ``fluxos_0.csv`` na pasta
``saida``) e grava:

  - resumo_contratos.xlsx  → subsidio, impacto, saldo final por contrato
  - resumo_por_ano.xlsx    → subsidio e impacto por (contrato, ano)

Compatível com colunas ContAgil/repo:
  impacto | impacto_fiscal
  saldo   | saldo_fiscal

Uso (WinPython ContAgil):
  python3 scripts/resumo_fluxos.py \\
      --fluxos "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida\\fluxos_0.csv"

Uso (repo / cloud):
  python3 scripts/resumo_fluxos.py --fluxos output/fluxos_amostra.xlsx
  python3 scripts/resumo_fluxos.py   # auto-detecta fluxos_0 / fluxos_completos_*
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from scripts.gerar_fluxos import CONTAGIL_PASTA_SAIDA, OUTPUT_DIR

# Caminho padrão do script colado no ContAgil (WinPython)
FLUXOS_0_DEFAULT = CONTAGIL_PASTA_SAIDA / "fluxos_0.csv"

CANDIDATOS_FLUXOS = (
    FLUXOS_0_DEFAULT,
    CONTAGIL_PASTA_SAIDA / "fluxos_0.xlsx",
    OUTPUT_DIR / "fluxos_completos_final.csv",
    OUTPUT_DIR / "fluxos_completos_final.xlsx",
    OUTPUT_DIR / "fluxos_completos_corrigido.csv",
    OUTPUT_DIR / "fluxos_completos_corrigido.xlsx",
    OUTPUT_DIR / "fluxos_amostra.csv",
    OUTPUT_DIR / "fluxos_amostra.xlsx",
    Path("/tmp/app-streamlit/output/fluxos_completos_corrigido.csv"),
)

COLUNAS_IMPACTO = ("impacto", "impacto_fiscal")
COLUNAS_SALDO = ("saldo", "saldo_fiscal")


def _primeira_coluna(columns: pd.Index, candidates: tuple[str, ...]) -> str | None:
    for name in candidates:
        if name in columns:
            return name
    return None


def resolver_fluxos(explicit: Path | None = None) -> Path:
    """Resolve o arquivo de parcelas (explícito ou candidatos ContAgil/output)."""
    if explicit is not None:
        return explicit
    for path in CANDIDATOS_FLUXOS:
        if path.exists() and path.is_file():
            return path
    # Qualquer fluxos_*.csv/.xlsx na pasta ContAgil saida ou output/
    for pasta in (CONTAGIL_PASTA_SAIDA, OUTPUT_DIR):
        if not pasta.exists():
            continue
        matches = sorted(pasta.glob("fluxos_*.csv")) + sorted(pasta.glob("fluxos_*.xlsx"))
        # Prefere fluxos_0.* e evita diários
        matches = [m for m in matches if "diario" not in m.stem.lower()]
        if matches:
            for m in matches:
                if m.stem == "fluxos_0" or m.name.startswith("fluxos_0."):
                    return m
            return matches[0]
    raise FileNotFoundError(
        "Nenhum arquivo de fluxos encontrado. Informe --fluxos ou gere com:\n"
        "  python3 scripts/contagil_fluxos.py --input data/sample_operacoes_com_agente.csv\n"
        "  python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --stem fluxos_amostra"
    )


def _eh_dataframe_parcelas(df: pd.DataFrame) -> bool:
    """True se o DataFrame parece detalhe de parcelas ContAgil."""
    cols = {str(c) for c in df.columns}
    tem_data = "data_fluxo" in cols
    tem_subsidio = "subsidio" in cols
    tem_contrato = "contrato" in cols
    tem_impacto = bool(cols & set(COLUNAS_IMPACTO))
    return tem_data and tem_subsidio and tem_contrato and tem_impacto


def carregar_fluxos(path: Path) -> pd.DataFrame:
    """Lê CSV ou Excel de parcelas ContAgil.

    Em Excel multi-aba, procura a primeira aba com colunas de parcelas
    (ignora abas Resumo / Por_Agente de workbooks agregados).
    """
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        xl = pd.ExcelFile(path)
        # Preferência: Sheet1 / Amostra_Parcelas / qualquer aba com parcelas
        preferidas = (
            "Sheet1",
            "Amostra_Parcelas",
            "Parcelas",
            "fluxos",
            "Fluxos",
        )
        ordem: list[str] = []
        for nome in preferidas:
            if nome in xl.sheet_names:
                ordem.append(nome)
        for nome in xl.sheet_names:
            if nome not in ordem:
                ordem.append(nome)
        for nome in ordem:
            df = pd.read_excel(xl, sheet_name=nome)
            if _eh_dataframe_parcelas(df):
                return df
        raise ValueError(
            f"Nenhuma aba de parcelas em {path.name}. "
            f"Abas: {xl.sheet_names}"
        )
    return pd.read_csv(path)


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Garante colunas canônicas: contrato, data_fluxo, subsidio, impacto, saldo."""
    work = df.copy()
    required = ("contrato", "data_fluxo", "subsidio")
    missing = [c for c in required if c not in work.columns]
    if missing:
        raise ValueError(
            f"Colunas ausentes: {missing}. Disponíveis: {list(work.columns)}"
        )

    impacto_col = _primeira_coluna(work.columns, COLUNAS_IMPACTO)
    if impacto_col is None:
        raise ValueError(
            "CSV precisa de 'impacto' ou 'impacto_fiscal'. "
            f"Disponíveis: {list(work.columns)}"
        )
    if impacto_col != "impacto":
        work["impacto"] = work[impacto_col]

    saldo_col = _primeira_coluna(work.columns, COLUNAS_SALDO)
    if saldo_col is None:
        work["saldo"] = pd.NA
    elif saldo_col != "saldo":
        work["saldo"] = work[saldo_col]

    work["data_fluxo"] = pd.to_datetime(work["data_fluxo"], errors="coerce")
    work["subsidio"] = pd.to_numeric(work["subsidio"], errors="coerce").fillna(0.0)
    work["impacto"] = pd.to_numeric(work["impacto"], errors="coerce").fillna(0.0)
    work["saldo"] = pd.to_numeric(work["saldo"], errors="coerce")
    before = len(work)
    work = work.dropna(subset=["data_fluxo"]).reset_index(drop=True)
    if len(work) < before:
        # Linhas sem data (abas/resumos misturados) são descartadas
        pass
    if work.empty:
        raise ValueError("Nenhuma parcela com data_fluxo válida após normalização.")
    return work


def resumo_por_contrato(df: pd.DataFrame) -> pd.DataFrame:
    """Total de subsídio/impacto e saldo final por contrato."""
    work = df.sort_values(["contrato", "data_fluxo"], kind="mergesort")
    resumo = (
        work.groupby("contrato", sort=True)
        .agg(
            subsidio=("subsidio", "sum"),
            impacto=("impacto", "sum"),
            saldo=("saldo", "last"),
            parcelas=("data_fluxo", "count"),
        )
        .round(2)
    )
    resumo = resumo.rename(
        columns={
            "subsidio": "Total Subsídio (R$)",
            "impacto": "Impacto Fiscal 2026 (R$)",
            "saldo": "Saldo Final (R$)",
            "parcelas": "Quantidade de Parcelas",
        }
    )
    return resumo


def resumo_por_ano(df: pd.DataFrame) -> pd.DataFrame:
    """Subsídio e impacto por contrato × ano de data_fluxo."""
    work = df.copy()
    work["ano"] = work["data_fluxo"].dt.year
    resumo = (
        work.groupby(["contrato", "ano"], sort=True)
        .agg(
            subsidio=("subsidio", "sum"),
            impacto=("impacto", "sum"),
            parcelas=("data_fluxo", "count"),
        )
        .round(2)
    )
    resumo = resumo.rename(
        columns={
            "subsidio": "Total Subsídio (R$)",
            "impacto": "Impacto Fiscal 2026 (R$)",
            "parcelas": "Quantidade de Parcelas",
        }
    )
    return resumo


def salvar_resumos(
    resumo_contrato: pd.DataFrame,
    resumo_ano: pd.DataFrame,
    output_dir: Path,
) -> tuple[Path, Path]:
    """Grava Excel (+ CSV espelho) na pasta de saída."""
    output_dir.mkdir(parents=True, exist_ok=True)
    path_contrato = output_dir / "resumo_contratos.xlsx"
    path_ano = output_dir / "resumo_por_ano.xlsx"

    resumo_contrato.to_excel(path_contrato)
    resumo_ano.to_excel(path_ano)
    resumo_contrato.to_csv(path_contrato.with_suffix(".csv"))
    resumo_ano.to_csv(path_ano.with_suffix(".csv"))
    return path_contrato, path_ano


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--fluxos",
        type=Path,
        default=None,
        help=(
            "CSV/XLSX de parcelas (default: ContAgil saida/fluxos_0.csv "
            "ou auto-detecta em output/)."
        ),
    )
    p.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Pasta de saída (default: mesma pasta do arquivo de fluxos).",
    )
    p.add_argument(
        "--contrato",
        type=int,
        default=None,
        help="Se informado, imprime o resumo por ano só desse contrato.",
    )
    p.add_argument(
        "--top",
        type=int,
        default=10,
        help="Linhas a imprimir no preview (default 10).",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        fluxos_path = resolver_fluxos(args.fluxos)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if not fluxos_path.exists():
        print(f"Arquivo não encontrado: {fluxos_path}", file=sys.stderr)
        return 1

    output_dir = args.output_dir if args.output_dir is not None else fluxos_path.parent

    print("Resumo de fluxos ContAgil — por contrato e por ano")
    print(f"Lendo: {fluxos_path}")
    try:
        df = normalizar_colunas(carregar_fluxos(fluxos_path))
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(f"Parcelas carregadas: {len(df):,}")

    resumo_contrato = resumo_por_contrato(df)
    resumo_ano = resumo_por_ano(df)

    print("\nResumo por Contrato:")
    print(resumo_contrato.head(args.top).to_string())

    contrato_preview = args.contrato
    if contrato_preview is None and len(resumo_ano.index):
        # Primeiro contrato do índice (compatível com print do script ContAgil)
        first = resumo_ano.index.get_level_values("contrato")[0]
        contrato_preview = first

    if contrato_preview is not None and contrato_preview in resumo_ano.index.get_level_values(
        "contrato"
    ):
        print(f"\nResumo por Ano (Contrato {contrato_preview}):")
        print(resumo_ano.loc[contrato_preview].head(args.top).to_string())
    else:
        print("\nResumo por Ano:")
        print(resumo_ano.head(args.top).to_string())

    path_contrato, path_ano = salvar_resumos(resumo_contrato, resumo_ano, output_dir)

    print(f"\n✅ Arquivos salvos em: {output_dir}")
    print(f"   {path_contrato.name}")
    print(f"   {path_ano.name}")
    print(f"   {path_contrato.with_suffix('.csv').name}")
    print(f"   {path_ano.with_suffix('.csv').name}")

    print("\n" + "=" * 60)
    print("TOTAIS")
    print("=" * 60)
    print(f"Contratos: {len(resumo_contrato):,}")
    print(
        f"Total Subsídio: R$ "
        f"{resumo_contrato['Total Subsídio (R$)'].sum():,.2f}"
    )
    print(
        f"Total Impacto Fiscal 2026: R$ "
        f"{resumo_contrato['Impacto Fiscal 2026 (R$)'].sum():,.2f}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
