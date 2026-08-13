#!/usr/bin/env python3
"""Gera fluxos a partir de BNDES_INDIRETAS_NUMERADOS.xlsx (uma aba por ano).

Para cada aba de contratos (2002, 2003, … com números ``N-AAAA``), gera as
parcelas correspondentes e grava:

  - CSV completo: ``saida/fluxos_por_ano_contrato/YYYY.csv``
  - Excel espelho (uma aba por ano, se couber em ~1M linhas):
    ``saida/FLUXOS_BNDES_INDIRETAS_POR_ANO_CONTRATO.xlsx``

Regra: todas as parcelas de um contrato ficam na aba/arquivo do **ano do
contrato** (ano da aba de origem). Impacto fiscal continua capitalizado na
``data_fluxo``.

Uso ContAgil::

  python scripts/fluxos_por_ano_contrato_numerados.py
  python scripts/fluxos_por_ano_contrato_numerados.py ^
    --numerados saida\\BNDES_INDIRETAS_NUMERADOS.xlsx ^
    --fatores fator_acumulado_SELIC_TJLP_TLP.xlsx
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

_SCRIPTS = Path(__file__).resolve().parent
ROOT = _SCRIPTS.parent
for _p in (str(ROOT), str(_SCRIPTS)):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _load_mod(name: str):
    import importlib.util

    path = _SCRIPTS / f"{name}.py"
    if path.exists():
        spec = importlib.util.spec_from_file_location(f"sec_{name}", path)
        if spec is not None and spec.loader is not None:
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    return __import__(f"scripts.{name}", fromlist=["*"])


_gf = _load_mod("gerar_fluxos")
try:
    _seg = _load_mod("contagil_fluxos_seguro")
except Exception:  # noqa: BLE001
    _seg = None

CONTAGIL_WINPYTHON = _gf.CONTAGIL_WINPYTHON
CONTAGIL_PASTA_SAIDA = _gf.CONTAGIL_PASTA_SAIDA
normalizar_colunas = _gf.normalizar_colunas
gerar_fluxos = _gf.gerar_fluxos
gerar_e_gravar_fluxos = _gf.gerar_e_gravar_fluxos

EXCEL_MAX = 1_000_000
_ANO_SHEET = re.compile(r"^(19|20)\d{2}$")
MARKER = "fluxos-por-ano-contrato-numerados-20260813"


def resolver_numerados(path: Optional[Path]) -> Path:
    candidatos = []
    if path is not None:
        candidatos.append(Path(path))
    candidatos.extend(
        [
            Path.cwd() / "saida" / "BNDES_INDIRETAS_NUMERADOS.xlsx",
            CONTAGIL_PASTA_SAIDA / "BNDES_INDIRETAS_NUMERADOS.xlsx",
            CONTAGIL_WINPYTHON / "saida" / "BNDES_INDIRETAS_NUMERADOS.xlsx",
            ROOT / "output" / "BNDES_INDIRETAS_NUMERADOS_DEMO.xlsx",
            ROOT / "saida" / "BNDES_INDIRETAS_NUMERADOS.xlsx",
        ]
    )
    for c in candidatos:
        if c.exists():
            return c
    raise FileNotFoundError(
        "BNDES_INDIRETAS_NUMERADOS.xlsx não encontrado. "
        "Rode numerar_contratos_indiretas.bat antes."
    )


def resolver_fatores(path: Optional[Path]):
    if _seg is None:
        return 0.145
    if path is not None and Path(path).exists():
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


def listar_abas_ano(path: Path) -> list[str]:
    xl = pd.ExcelFile(path)
    anos = [s for s in xl.sheet_names if _ANO_SHEET.match(str(s).strip())]
    if not anos:
        # fallback: qualquer aba cujo nome contenha ano
        anos = [s for s in xl.sheet_names if re.search(r"(19|20)\d{2}", str(s))]
    if not anos:
        raise ValueError(f"Nenhuma aba de ano em {path}: {xl.sheet_names}")
    return sorted(anos, key=lambda s: int(re.search(r"(19|20)\d{2}", str(s)).group(0)))


def _escrever_aba_excel(wb: Workbook, nome: str, df: pd.DataFrame, *, first: bool) -> None:
    if first:
        ws = wb.active
        ws.title = str(nome)[:31]
    else:
        ws = wb.create_sheet(str(nome)[:31])
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(color="FFFFFF", bold=True)
    for i, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
        ws.append(list(row))
        if i == 0:
            for col in range(1, len(row) + 1):
                cell = ws.cell(1, col)
                cell.fill = fill
                cell.font = font
                cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"


def processar(
    numerados: Path,
    pasta_saida: Path,
    *,
    fatores,
    ano_min: Optional[int] = None,
    ano_max: Optional[int] = None,
    lote: int = 2_000,
    excel_max: int = EXCEL_MAX,
) -> Path:
    pasta_saida = Path(pasta_saida)
    pasta_csv = pasta_saida / "fluxos_por_ano_contrato"
    pasta_csv.mkdir(parents=True, exist_ok=True)
    xlsx_out = pasta_saida / "FLUXOS_BNDES_INDIRETAS_POR_ANO_CONTRATO.xlsx"

    abas = listar_abas_ano(numerados)
    wb = Workbook()
    first_sheet = True
    resumo_rows = []

    print(f"[{MARKER}] numerados={numerados}")
    print(f"[{MARKER}] abas={abas}")

    for aba in abas:
        m = re.search(r"(19|20)\d{2}", str(aba))
        ano = int(m.group(0)) if m else None
        if ano is None:
            continue
        if ano_min is not None and ano < ano_min:
            continue
        if ano_max is not None and ano > ano_max:
            continue

        print(f"\n=== Ano {ano} (aba '{aba}') ===")
        bruto = pd.read_excel(numerados, sheet_name=aba)
        if bruto.empty:
            print("  [AVISO] aba vazia — pulando")
            continue

        contratos = normalizar_colunas(bruto)
        if contratos.empty:
            print("  [AVISO] nenhum contrato válido — pulando")
            continue

        csv_ano = pasta_csv / f"{ano}.csv"
        xlsx_ano = pasta_csv / f"{ano}.xlsx"
        t0 = time.time()
        stats = gerar_e_gravar_fluxos(
            contratos,
            fatores,
            saida_xlsx=xlsx_ano,
            lote=lote,
            excel_max_linhas=excel_max,
        )
        # Padroniza nome CSV ano
        csv_gerado = Path(stats["csv"])
        if csv_gerado.exists() and csv_gerado.resolve() != csv_ano.resolve():
            if csv_ano.exists():
                csv_ano.unlink()
            csv_gerado.replace(csv_ano)
            stats["csv"] = str(csv_ano)

        n_parc = int(stats["parcelas"])
        print(
            f"  OK ano {ano}: {stats['contratos']:,} contratos | "
            f"{n_parc:,} parcelas | {time.time() - t0:.1f}s"
        )
        resumo_rows.append(
            {
                "ano_contrato": ano,
                "qtd_contratos": stats["contratos"],
                "qtd_parcelas": n_parc,
                "csv": str(csv_ano),
                "xlsx_ano": str(xlsx_ano),
            }
        )

        # Aba no workbook consolidado (amostra/completo se couber)
        if csv_ano.exists():
            if n_parc <= excel_max:
                df_aba = pd.read_csv(csv_ano)
            else:
                df_aba = pd.read_csv(csv_ano, nrows=excel_max)
                print(
                    f"  [AVISO] Ano {ano} tem {n_parc:,} parcelas > {excel_max:,}; "
                    f"aba Excel com amostra; CSV completo em {csv_ano.name}"
                )
            _escrever_aba_excel(wb, str(ano), df_aba, first=first_sheet)
            first_sheet = False

    if first_sheet:
        wb.active.title = "vazio"
    else:
        # aba resumo
        ws = wb.create_sheet("RESUMO", 0)
        fill = PatternFill("solid", fgColor="1F4E79")
        font = Font(color="FFFFFF", bold=True)
        headers = list(resumo_rows[0].keys()) if resumo_rows else []
        ws.append(headers)
        for col in range(1, len(headers) + 1):
            cell = ws.cell(1, col)
            cell.fill = fill
            cell.font = font
        for row in resumo_rows:
            ws.append([row[h] for h in headers])
        ws2 = wb.create_sheet("NOTA")
        ws2["A1"] = (
            "Cada aba YYYY contém as parcelas dos contratos da aba YYYY de "
            "BNDES_INDIRETAS_NUMERADOS.xlsx (numeração N-AAAA). "
            "Todas as parcelas de um contrato ficam no ano do contrato; "
            "o impacto fiscal de cada parcela continua capitalizado na data_fluxo. "
            f"CSVs completos em: {pasta_csv}"
        )
        ws2.column_dimensions["A"].width = 110

    pasta_saida.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_out)
    print(f"\n[OK] Excel consolidado: {xlsx_out}")
    print(f"[OK] CSVs por ano: {pasta_csv}")
    return xlsx_out


def main(argv: Optional[list[str]] = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--numerados",
        type=Path,
        default=None,
        help="BNDES_INDIRETAS_NUMERADOS.xlsx (default: saida/ ContAgil)",
    )
    p.add_argument(
        "--saida",
        type=Path,
        default=None,
        help="Pasta de saída (default: saida ContAgil ou ./saida)",
    )
    p.add_argument("--fatores", type=Path, default=None)
    p.add_argument("--ano-min", type=int, default=None)
    p.add_argument("--ano-max", type=int, default=None)
    p.add_argument("--lote", type=int, default=2_000)
    args = p.parse_args(argv)

    numerados = resolver_numerados(args.numerados)
    if args.saida is not None:
        pasta_saida = Path(args.saida)
    elif CONTAGIL_PASTA_SAIDA.exists():
        pasta_saida = CONTAGIL_PASTA_SAIDA
    else:
        pasta_saida = Path.cwd() / "saida"
        pasta_saida.mkdir(exist_ok=True)

    fatores = resolver_fatores(args.fatores)
    processar(
        numerados,
        pasta_saida,
        fatores=fatores,
        ano_min=args.ano_min,
        ano_max=args.ano_max,
        lote=args.lote,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
