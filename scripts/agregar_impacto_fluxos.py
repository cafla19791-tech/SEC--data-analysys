#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agrega impacto fiscal / subsídio a partir de fluxos_*.csv grandes (streaming).

Pensado para a massa ContAgil WinPython após ``contagil_fluxos.py``:
  ~70 milhões de parcelas em vários CSV — não cabe em memória.

Lê só as colunas necessárias em chunks, acumula por ano e por agente, e grava:

  - impacto_fiscal_por_ano.xlsx / .csv
  - resumo_por_agente.xlsx / .csv
  - resumo_impacto_bndes.xlsx  (abas: Impacto_Por_Ano, Por_Agente, Totais)

Modo padrão: ``coluna`` (usa ``impacto_fiscal`` já gravado na geração).

Uso (WinPython ContAgil)::

  python scripts\\agregar_impacto_fluxos.py --pasta \"%cd%\\saida\"

  # ou com caminho absoluto:
  python scripts\\agregar_impacto_fluxos.py ^
    --pasta \"C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida\"
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
import time
import types
from collections import defaultdict
from pathlib import Path

import pandas as pd

# WinPython ContAgil: carrega irmãos por caminho (não depende de pacote instalado).
_SCRIPTS_DIR = Path(__file__).resolve().parent
ROOT = _SCRIPTS_DIR.parent

# Marcador para o usuário/scripts de download validarem a versão
MARKER = "agregar-impacto-streaming-20260816d-ano-contrato"


def _load_sibling(mod_name: str):
    """Carrega ``scripts/<mod_name>.py`` via importlib (ContAgil/WinPython)."""
    full = f"scripts.{mod_name}"
    if full in sys.modules:
        return sys.modules[full]
    path = _SCRIPTS_DIR / f"{mod_name}.py"
    if not path.is_file():
        print(f"ERRO [{MARKER}]: falta o arquivo:")
        print(f"  {path}")
        print("No PowerShell:")
        b = (
            "https://raw.githubusercontent.com/cafla19791-tech/"
            "SEC--data-analysys/cursor/agregar-impacto-streaming-f342"
        )
        print(f'  $b="{b}"')
        print(
            f'  Invoke-WebRequest "$b/scripts/{mod_name}.py" '
            f"-OutFile scripts\\{mod_name}.py"
        )
        raise SystemExit(2)
    if "scripts" not in sys.modules:
        pkg = types.ModuleType("scripts")
        pkg.__path__ = [str(_SCRIPTS_DIR)]
        pkg.__package__ = "scripts"
        sys.modules["scripts"] = pkg
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    spec = importlib.util.spec_from_file_location(full, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Nao foi possivel carregar {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[full] = mod
    sys.modules[mod_name] = mod
    spec.loader.exec_module(mod)
    return mod


_gf = _load_sibling("gerar_fluxos")
_imp = _load_sibling("impacto_fiscal_por_ano")

CONTAGIL_PASTA_SAIDA = _gf.CONTAGIL_PASTA_SAIDA
OUTPUT_DIR = _gf.OUTPUT_DIR
TAXA_SELIC_ANUAL = _gf.TAXA_SELIC_ANUAL

DATA_REFERENCIA = _imp.DATA_REFERENCIA
MODOS = _imp.MODOS
_coluna_impacto = _imp._coluna_impacto
_impacto_composta = _imp._impacto_composta
_impacto_contagil = _imp._impacto_contagil
_impacto_recalcular = _imp._impacto_recalcular
calcular_meses_ate_2026 = _imp.calcular_meses_ate_2026
carregar_serie_selic = _imp.carregar_serie_selic

CHUNK_DEFAULT = 500_000
COLUNAS_UTEIS = (
    "data_fluxo",
    "data_contratacao",
    "ano_contrato",
    "ano_fluxo",
    "subsidio",
    "impacto_fiscal",
    "impacto",
    "Instituição Financeira",
    "agente",
    "mes",
    "contrato",
)
AGRUPAR_POR = ("fluxo", "contrato")
COL_ANO_FLUXO = "Ano"
COL_ANO_CONTRATO = "Ano do Contrato"


def listar_csvs_fluxos(pasta: Path) -> list[Path]:
    """Lista ``fluxos_*.csv`` na pasta (ignora diários)."""
    pasta = Path(pasta)
    if not pasta.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {pasta}")
    csvs = sorted(
        p
        for p in pasta.glob("fluxos_*.csv")
        if "diario" not in p.stem.lower()
    )
    if not csvs:
        raise FileNotFoundError(
            f"Nenhum fluxos_*.csv em {pasta}. "
            "Gere com scripts/contagil_fluxos.py primeiro."
        )
    return csvs


def _resolver_agente_col(columns: pd.Index) -> str | None:
    if "Instituição Financeira" in columns:
        return "Instituição Financeira"
    if "agente" in columns:
        return "agente"
    return None


def _impacto_chunk(
    chunk: pd.DataFrame,
    *,
    modo: str,
    taxa_selic_anual: float,
    selic_serie,
) -> pd.Series:
    """Calcula impacto individual do chunk conforme o modo."""
    if modo == "coluna":
        col = _coluna_impacto(chunk.columns)
        if col is None:
            raise ValueError(
                "Modo 'coluna' exige impacto_fiscal ou impacto no CSV."
            )
        return pd.to_numeric(chunk[col], errors="coerce").fillna(0.0)

    data = pd.to_datetime(chunk["data_fluxo"], errors="coerce")
    subsidio = pd.to_numeric(chunk["subsidio"], errors="coerce").fillna(0.0)

    if modo == "contagil":
        if selic_serie is None:
            raise ValueError("Modo contagil exige série SELIC.")
        return _impacto_contagil(subsidio, data, selic_serie)

    meses = data.apply(calcular_meses_ate_2026)
    if modo == "composta":
        return _impacto_composta(subsidio, meses, taxa_selic_anual)
    return _impacto_recalcular(subsidio, meses, taxa_selic_anual)


def _anos_agrupamento(chunk: pd.DataFrame, agrupar_por: str) -> pd.Series:
    """Extrai o ano de agrupamento (fluxo ou contrato) do chunk."""
    if agrupar_por == "contrato":
        if "ano_contrato" in chunk.columns:
            anos = pd.to_numeric(chunk["ano_contrato"], errors="coerce")
        elif "data_contratacao" in chunk.columns:
            anos = pd.to_datetime(chunk["data_contratacao"], errors="coerce").dt.year
        else:
            raise ValueError(
                "agrupar_por=contrato exige coluna ano_contrato ou data_contratacao no CSV."
            )
        return anos
    # fluxo (padrão)
    if "ano_fluxo" in chunk.columns:
        return pd.to_numeric(chunk["ano_fluxo"], errors="coerce")
    return pd.to_datetime(chunk["data_fluxo"], errors="coerce").dt.year


def agregar_streaming(
    arquivos: list[Path],
    *,
    modo: str = "coluna",
    chunksize: int = CHUNK_DEFAULT,
    taxa_selic_anual: float = TAXA_SELIC_ANUAL,
    selic_serie=None,
    agrupar_por: str = "fluxo",
) -> dict:
    """Agrega por ano e por agente sem carregar todos os CSV na memória.

    ``agrupar_por``:
      - ``fluxo`` (padrão): ano de ``data_fluxo`` / ``ano_fluxo``
      - ``contrato``: ano de ``ano_contrato`` / ``data_contratacao``

    Retorna dict com DataFrames ``por_ano``, ``por_agente``, ``totais`` e
    metadados (parcelas, arquivos, segundos).
    """
    if modo not in MODOS:
        raise ValueError(f"Modo desconhecido: {modo}")
    if agrupar_por not in AGRUPAR_POR:
        raise ValueError(f"agrupar_por deve ser um de {AGRUPAR_POR}")

    col_ano = COL_ANO_CONTRATO if agrupar_por == "contrato" else COL_ANO_FLUXO

    # ano -> [subsidio, impacto, qtd]
    acc_ano: dict[int, list[float]] = defaultdict(lambda: [0.0, 0.0, 0.0])
    # agente -> [subsidio, impacto, qtd_parcelas, set_contratos]
    acc_ag: dict[str, list] = defaultdict(lambda: [0.0, 0.0, 0.0, set()])

    total_parcelas = 0
    t0 = time.time()

    for idx_arq, path in enumerate(arquivos, start=1):
        print(f"[{idx_arq}/{len(arquivos)}] {path.name} ...")
        sys.stdout.flush()
        header = pd.read_csv(path, nrows=0)
        if "data_fluxo" not in header.columns or "subsidio" not in header.columns:
            raise ValueError(
                f"{path.name}: precisa de data_fluxo e subsidio. "
                f"Colunas: {list(header.columns)}"
            )
        if agrupar_por == "contrato" and (
            "ano_contrato" not in header.columns
            and "data_contratacao" not in header.columns
        ):
            raise ValueError(
                f"{path.name}: agrupar_por=contrato exige ano_contrato ou data_contratacao."
            )
        if modo == "coluna" and _coluna_impacto(header.columns) is None:
            raise ValueError(
                f"{path.name}: modo coluna exige impacto_fiscal/impacto."
            )

        usecols = [c for c in COLUNAS_UTEIS if c in header.columns]
        agente_col = _resolver_agente_col(header.columns)
        lidas_arq = 0

        for chunk in pd.read_csv(path, usecols=usecols, chunksize=chunksize):
            impacto = _impacto_chunk(
                chunk,
                modo=modo,
                taxa_selic_anual=taxa_selic_anual,
                selic_serie=selic_serie,
            )
            subsidio = pd.to_numeric(chunk["subsidio"], errors="coerce").fillna(0.0)
            anos = _anos_agrupamento(chunk, agrupar_por)

            # Por ano (groupby no chunk — rápido e leve)
            tmp = pd.DataFrame(
                {
                    "ano": anos,
                    "subsidio": subsidio.to_numpy(),
                    "impacto": impacto.to_numpy(),
                }
            ).dropna(subset=["ano"])
            tmp["ano"] = tmp["ano"].astype(int)
            g = tmp.groupby("ano", sort=False).agg(
                subsidio=("subsidio", "sum"),
                impacto=("impacto", "sum"),
                qtd=("impacto", "count"),
            )
            for ano, row in g.iterrows():
                a = acc_ano[int(ano)]
                a[0] += float(row["subsidio"])
                a[1] += float(row["impacto"])
                a[2] += float(row["qtd"])

            if agente_col is not None:
                agentes = (
                    chunk[agente_col]
                    .fillna("NÃO INFORMADO")
                    .astype(str)
                    .str.strip()
                    .replace("", "NÃO INFORMADO")
                )
                tmp_ag = pd.DataFrame(
                    {
                        "agente": agentes.to_numpy(),
                        "subsidio": subsidio.to_numpy(),
                        "impacto": impacto.to_numpy(),
                    }
                )
                g_ag = tmp_ag.groupby("agente", sort=False).agg(
                    subsidio=("subsidio", "sum"),
                    impacto=("impacto", "sum"),
                    qtd=("impacto", "count"),
                )
                for ag, row in g_ag.iterrows():
                    a = acc_ag[str(ag)]
                    a[0] += float(row["subsidio"])
                    a[1] += float(row["impacto"])
                    a[2] += float(row["qtd"])

                if "contrato" in chunk.columns:
                    for ag, grp in chunk.groupby(agentes, sort=False):
                        acc_ag[str(ag)][3].update(
                            grp["contrato"].dropna().unique().tolist()
                        )

            n = len(chunk)
            lidas_arq += n
            total_parcelas += n
            elapsed = max(time.time() - t0, 1e-6)
            rate = total_parcelas / elapsed
            print(
                f"  +{n:,} | arquivo={lidas_arq:,} | total={total_parcelas:,} "
                f"| {rate:,.0f} parc/s"
            )
            sys.stdout.flush()

        print(f"  OK {path.name}: {lidas_arq:,} parcelas")
        sys.stdout.flush()

    por_ano = (
        pd.DataFrame(
            [
                {
                    col_ano: ano,
                    "Soma Subsídio Nominal (R$)": round(v[0], 2),
                    "Impacto Fiscal 2026 (R$)": round(v[1], 2),
                    "Quantidade de Parcelas": int(v[2]),
                }
                for ano, v in sorted(acc_ano.items())
            ]
        )
        if acc_ano
        else pd.DataFrame(
            columns=[
                col_ano,
                "Soma Subsídio Nominal (R$)",
                "Impacto Fiscal 2026 (R$)",
                "Quantidade de Parcelas",
            ]
        )
    )

    por_agente_rows = []
    for ag, v in acc_ag.items():
        por_agente_rows.append(
            {
                "Instituição Financeira": ag,
                "Qtd Contratos": len(v[3]) if v[3] else None,
                "Qtd Parcelas": int(v[2]),
                "Total Subsídio (R$)": round(v[0], 2),
                "Impacto Fiscal 2026 (R$)": round(v[1], 2),
            }
        )
    por_agente = pd.DataFrame(por_agente_rows)
    if not por_agente.empty:
        por_agente = por_agente.sort_values(
            "Impacto Fiscal 2026 (R$)", ascending=False
        ).reset_index(drop=True)
        # Se não houve coluna contrato, remove Qtd Contratos vazia
        if por_agente["Qtd Contratos"].isna().all():
            por_agente = por_agente.drop(columns=["Qtd Contratos"])

    tot_sub = float(por_ano["Soma Subsídio Nominal (R$)"].sum()) if len(por_ano) else 0.0
    tot_imp = float(por_ano["Impacto Fiscal 2026 (R$)"].sum()) if len(por_ano) else 0.0
    totais = pd.DataFrame(
        [
            {"Métrica": "Total Subsídio Nominal (R$)", "Valor": round(tot_sub, 2)},
            {"Métrica": "Total Impacto Fiscal 2026 (R$)", "Valor": round(tot_imp, 2)},
            {"Métrica": "Total de Parcelas", "Valor": int(total_parcelas)},
            {"Métrica": "Arquivos CSV", "Valor": len(arquivos)},
            {
                "Métrica": "Referência de impacto",
                "Valor": DATA_REFERENCIA.strftime("%d/%m/%Y"),
            },
            {"Métrica": "Modo", "Valor": modo},
            {
                "Métrica": "Agrupamento por ano",
                "Valor": "contrato (data_contratacao)"
                if agrupar_por == "contrato"
                else "fluxo (data_fluxo)",
            },
        ]
    )

    return {
        "por_ano": por_ano,
        "por_agente": por_agente,
        "totais": totais,
        "parcelas": total_parcelas,
        "arquivos": [str(p) for p in arquivos],
        "segundos": round(time.time() - t0, 1),
        "modo": modo,
        "agrupar_por": agrupar_por,
        "col_ano": col_ano,
    }


def salvar_resultados(result: dict, pasta_saida: Path) -> dict[str, Path]:
    """Grava CSV/XLSX de impacto por ano, por agente e workbook combinado."""
    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)
    paths: dict[str, Path] = {}

    ano_xlsx = pasta_saida / "impacto_fiscal_por_ano.xlsx"
    ano_csv = pasta_saida / "impacto_fiscal_por_ano.csv"
    result["por_ano"].to_excel(ano_xlsx, index=False)
    result["por_ano"].to_csv(ano_csv, index=False)
    paths["impacto_ano_xlsx"] = ano_xlsx
    paths["impacto_ano_csv"] = ano_csv

    if not result["por_agente"].empty:
        ag_xlsx = pasta_saida / "resumo_por_agente.xlsx"
        ag_csv = pasta_saida / "resumo_por_agente.csv"
        result["por_agente"].to_excel(ag_xlsx, index=False)
        result["por_agente"].to_csv(ag_csv, index=False)
        paths["agente_xlsx"] = ag_xlsx
        paths["agente_csv"] = ag_csv

    wb = pasta_saida / "resumo_impacto_bndes.xlsx"
    with pd.ExcelWriter(wb, engine="openpyxl") as writer:
        sheet_ano = (
            "Impacto_Por_Ano_Contrato"
            if result.get("agrupar_por") == "contrato"
            else "Impacto_Por_Ano"
        )
        result["por_ano"].to_excel(writer, sheet_name=sheet_ano, index=False)
        if not result["por_agente"].empty:
            result["por_agente"].to_excel(writer, sheet_name="Por_Agente", index=False)
        result["totais"].to_excel(writer, sheet_name="Totais", index=False)
    paths["workbook"] = wb
    return paths


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "--pasta",
        type=Path,
        default=None,
        help=(
            "Pasta com fluxos_*.csv (default: ContAgil saida/ se existir, "
            "senão output/)."
        ),
    )
    p.add_argument(
        "--modo",
        choices=MODOS,
        default="coluna",
        help=(
            "coluna = impacto_fiscal já no CSV (padrão, recomendado após gerar); "
            "contagil / recalcular / composta = recalcula a partir do subsídio."
        ),
    )
    p.add_argument(
        "--arquivo-selic",
        type=Path,
        default=None,
        help="STP ContAgil / fatores (só modos contagil).",
    )
    p.add_argument(
        "--baixar-selic",
        action="store_true",
        help="Baixa SELIC Bacen se não houver STP (modo contagil).",
    )
    p.add_argument(
        "--taxa-selic",
        type=float,
        default=TAXA_SELIC_ANUAL,
        help=f"SELIC anual nos modos recalcular/composta (default {TAXA_SELIC_ANUAL}).",
    )
    p.add_argument(
        "--chunksize",
        type=int,
        default=CHUNK_DEFAULT,
        help=f"Linhas por chunk (default {CHUNK_DEFAULT:,}).",
    )
    p.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Pasta de saída dos resumos (default: --pasta).",
    )
    p.add_argument(
        "--agrupar-por",
        choices=AGRUPAR_POR,
        default="fluxo",
        help=(
            "fluxo = ano da parcela/data_fluxo (padrão); "
            "contrato = ano da data_contratacao / ano_contrato."
        ),
    )
    return p.parse_args(argv)


def _resolver_pasta(explicit: Path | None) -> Path:
    if explicit is not None:
        return Path(explicit)
    if CONTAGIL_PASTA_SAIDA.exists():
        return CONTAGIL_PASTA_SAIDA
    return OUTPUT_DIR


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print(f"[agregar_impacto_fluxos {MARKER}]")
    print("=" * 70)
    print("AGREGAÇÃO STREAMING — IMPACTO FISCAL BNDES")
    print(f"Referência: {DATA_REFERENCIA:%d/%m/%Y}")
    print("=" * 70)

    pasta = _resolver_pasta(args.pasta)
    try:
        arquivos = listar_csvs_fluxos(pasta)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(f"Pasta : {pasta}")
    print(f"CSVs  : {len(arquivos)}")
    for a in arquivos:
        try:
            mb = a.stat().st_size / (1024 * 1024)
            print(f"  - {a.name} ({mb:,.1f} MB)")
        except OSError:
            print(f"  - {a.name}")
    print(f"Modo  : {args.modo}")
    print(f"Ano   : {args.agrupar_por}")
    print()

    modo = args.modo
    selic_serie = None
    if modo == "contagil":
        try:
            selic_serie = carregar_serie_selic(args.arquivo_selic, args.baixar_selic)
        except FileNotFoundError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        if selic_serie is None:
            print(
                "Aviso: sem STP/Bacen; caindo para --modo coluna.",
                file=sys.stderr,
            )
            modo = "coluna"

    try:
        result = agregar_streaming(
            arquivos,
            modo=modo,
            chunksize=max(10_000, int(args.chunksize)),
            taxa_selic_anual=args.taxa_selic,
            selic_serie=selic_serie,
            agrupar_por=args.agrupar_por,
        )
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    out_dir = Path(args.output_dir) if args.output_dir is not None else pasta
    paths = salvar_resultados(result, out_dir)

    print()
    print("=" * 70)
    print("IMPACTO FISCAL POR ANO")
    print("=" * 70)
    print(result["por_ano"].to_string(index=False))
    print()
    print("=" * 70)
    print("TOTAIS")
    print("=" * 70)
    for _, row in result["totais"].iterrows():
        print(f"  {row['Métrica']}: {row['Valor']}")
    print(f"  Tempo: {result['segundos']} s")
    print()
    if not result["por_agente"].empty:
        print("Top 15 agentes (Impacto Fiscal 2026):")
        print(result["por_agente"].head(15).to_string(index=False))
        print()
    print("Arquivos gerados:")
    for key, path in paths.items():
        print(f"  → {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
