#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Impacto fiscal das OPERACOES DIRETAS (ContAgil).

Pipeline (mesma metodologia das indiretas — parcelas + impacto_fiscal):

  1. Localiza OPERACOES DIRETAS*.xlsx
  2. Gera fluxos (CSV streaming) em saida/fluxos_diretas/
  3. Agrega impacto por ANO DO CONTRATO / agente em saida/impacto_diretas/
  4. Monta APRESENTACAO_IMPACTO_BNDES_DIRETAS.xlsx (aba Por_Ano_Contrato)

Uso (ContAgil)::

  python sec_scripts\\impacto_operacoes_diretas.py
  python sec_scripts\\impacto_operacoes_diretas.py --excel \"OPERACOES DIRETAS.xlsx\"
  python sec_scripts\\impacto_operacoes_diretas.py --so-agregar
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
import types
from pathlib import Path

_SCRIPTS_DIR = Path(__file__).resolve().parent
ROOT = _SCRIPTS_DIR.parent

MARKER = "impacto-operacoes-diretas-20260816b"


def _load_sibling(mod_name: str):
    full = f"scripts.{mod_name}"
    if full in sys.modules:
        return sys.modules[full]
    path = _SCRIPTS_DIR / f"{mod_name}.py"
    if not path.is_file():
        print(f"ERRO [{MARKER}]: falta {path}")
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
_seg = _load_sibling("contagil_fluxos_seguro")
_ag = _load_sibling("agregar_impacto_fluxos")
_ap = _load_sibling("apresentacao_impacto_bndes")

CONTAGIL_WINPYTHON = _gf.CONTAGIL_WINPYTHON
CONTAGIL_PASTA_SAIDA = _gf.CONTAGIL_PASTA_SAIDA
load_from_excel = _gf.load_from_excel
gerar_e_gravar_fluxos = _gf.gerar_e_gravar_fluxos
_excel_tem_colunas_contratos = _gf._excel_tem_colunas_contratos

NOMES_EXPLICITOS = (
    "OPERACOES DIRETAS.xlsx",
    "OPERACOES DIRETAS - 2002 a 2018.xlsx",
    "OPERACOES DIRETAS - 2002 a 2018_calculado.xlsx",
    "OPERAÇÕES DIRETAS2002 A 302026.xlsx",
    "OPERACOES DIRETAS2002 A 302026.xlsx",
)


def resolver_excel_diretas(explicit: Path | None = None) -> Path:
    if explicit is not None:
        p = Path(explicit)
        if not p.is_file():
            raise FileNotFoundError(f"Excel nao encontrado: {p}")
        return p

    bases = [
        Path.cwd(),
        CONTAGIL_WINPYTHON,
        CONTAGIL_PASTA_SAIDA,
        Path.cwd() / "saida",
        Path.cwd() / "dados",
        CONTAGIL_WINPYTHON / "dados",
        ROOT / "saida",
        ROOT / "data",
    ]
    for base in bases:
        if not base.exists():
            continue
        for nome in NOMES_EXPLICITOS:
            cand = base / nome
            if cand.is_file():
                return cand
        for cand in sorted(base.glob("*DIRETA*.xlsx")):
            low = cand.name.lower()
            if "indireta" in low:
                continue
            if "discriminativo" in low:
                continue
            if low.endswith("_calculado.xlsx") and "direta" not in low:
                continue
            return cand
    raise FileNotFoundError(
        "OPERACOES DIRETAS.xlsx nao encontrado. "
        "Coloque o arquivo na pasta winpython ou passe --excel."
    )


def detectar_header(path: Path) -> int:
    import pandas as pd

    for h0 in (0, 5, 1, 2, 3, 4, 6):
        try:
            df = pd.read_excel(path, header=h0, nrows=5)
        except Exception:
            continue
        if _excel_tem_colunas_contratos(df):
            return h0
    return 0


def carregar_serie_fatores(path: Path | None = None):
    if path is not None:
        return _seg.carregar_fatores_mensais(Path(path))
    for c in (
        Path.cwd() / "fator_acumulado_SELIC_TJLP_TLP.xlsx",
        CONTAGIL_WINPYTHON / "fator_acumulado_SELIC_TJLP_TLP.xlsx",
        ROOT / "data" / "selic_taxas_contagil.xlsx",
    ):
        if c.exists():
            print(f"[INFO] Fatores: {c}")
            return _seg.carregar_fatores_mensais(c)
    print("[AVISO] Sem arquivo de fatores — usando SELIC 14,5% a.a. composta.")
    return 0.145


def gerar_fluxos_diretas(
    excel: Path,
    pasta_fluxos: Path,
    *,
    fatores,
    header: int | None = None,
) -> Path:
    h0 = detectar_header(excel) if header is None else int(header)
    print(f"[INFO] Excel: {excel}")
    print(f"[INFO] Header (0-based): {h0}")
    df = load_from_excel(excel, header=h0)
    pasta_fluxos.mkdir(parents=True, exist_ok=True)
    saida_xlsx = pasta_fluxos / "fluxos_OPERACOES_DIRETAS.xlsx"
    gerar_e_gravar_fluxos(
        df,
        fatores,
        saida_xlsx=saida_xlsx,
        lote=2_000,
    )
    csv_path = saida_xlsx.with_suffix(".csv")
    if not csv_path.is_file():
        raise FileNotFoundError(f"CSV de fluxos nao gerado: {csv_path}")
    print(f"[OK] Fluxos CSV: {csv_path}")
    return csv_path


def agregar_diretas(pasta_fluxos: Path, pasta_impacto: Path) -> dict:
    pasta_impacto.mkdir(parents=True, exist_ok=True)
    arquivos = _ag.listar_csvs_fluxos(pasta_fluxos)
    print(f"[INFO] Agregando {len(arquivos)} CSV(s) de {pasta_fluxos}")
    print("[INFO] Agrupamento por ANO DO CONTRATO (data_contratacao)")
    result = _ag.agregar_streaming(
        arquivos, modo="coluna", chunksize=500_000, agrupar_por="contrato"
    )
    paths = _ag.salvar_resultados(result, pasta_impacto)
    # workbook com nome DIRETAS
    wb_dir = pasta_impacto / "resumo_impacto_bndes_diretas.xlsx"
    if "workbook" in paths and paths["workbook"].exists():
        import shutil

        shutil.copy2(paths["workbook"], wb_dir)
        paths["workbook_diretas"] = wb_dir
    print(f"[OK] Impacto em: {pasta_impacto}")
    return {"result": result, "paths": paths}


def apresentar_diretas(pasta_impacto: Path, pasta_saida: Path) -> Path:
    out = pasta_saida / "APRESENTACAO_IMPACTO_BNDES_DIRETAS.xlsx"
    return _ap.construir_apresentacao(
        pasta_impacto,
        out,
        titulo="Impacto Fiscal — BNDES Diretas",
    )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--excel", type=Path, default=None, help="OPERACOES DIRETAS.xlsx")
    p.add_argument("--fatores", type=Path, default=None)
    p.add_argument("--header", type=int, default=None, help="Linha do cabecalho 0-based")
    p.add_argument(
        "--pasta-saida",
        type=Path,
        default=None,
        help="Pasta saida/ (default ContAgil ou ./saida)",
    )
    p.add_argument(
        "--so-agregar",
        action="store_true",
        help="Pula geracao de fluxos; so agrega CSV ja existente",
    )
    p.add_argument(
        "--so-apresentacao",
        action="store_true",
        help="Pula fluxos/agregar; so monta Excel de apresentacao",
    )
    return p.parse_args(argv)


def _resolver_pasta_saida(explicit: Path | None) -> Path:
    if explicit is not None:
        return Path(explicit)
    if CONTAGIL_PASTA_SAIDA.exists():
        return CONTAGIL_PASTA_SAIDA
    cand = Path.cwd() / "saida"
    cand.mkdir(parents=True, exist_ok=True)
    return cand


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print(f"[{MARKER}]")
    print("=" * 70)
    print("IMPACTO FISCAL — OPERACOES DIRETAS")
    print("=" * 70)

    pasta_saida = _resolver_pasta_saida(args.pasta_saida)
    pasta_fluxos = pasta_saida / "fluxos_diretas"
    pasta_impacto = pasta_saida / "impacto_diretas"

    try:
        if args.so_apresentacao:
            out = apresentar_diretas(pasta_impacto, pasta_saida)
            print(f"[OK] Apresentacao: {out}")
            return 0

        if not args.so_agregar:
            excel = resolver_excel_diretas(args.excel)
            fatores = carregar_serie_fatores(args.fatores)
            gerar_fluxos_diretas(
                excel, pasta_fluxos, fatores=fatores, header=args.header
            )

        if not pasta_fluxos.exists() or not any(pasta_fluxos.glob("fluxos_*.csv")):
            print(
                f"ERRO: nenhum fluxos_*.csv em {pasta_fluxos}. "
                "Rode sem --so-agregar primeiro.",
                file=sys.stderr,
            )
            return 1

        agregar_diretas(pasta_fluxos, pasta_impacto)
        out = apresentar_diretas(pasta_impacto, pasta_saida)
        print()
        print("[OK] Pipeline DIRETAS concluido")
        print(f"  Fluxos CSV : {pasta_fluxos}")
        print(f"  Impacto    : {pasta_impacto}")
        print(f"  Apresentacao: {out}")
        return 0
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
