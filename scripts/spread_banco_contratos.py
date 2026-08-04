#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Calcula o valor (R$) do spread do banco em cada contrato BNDES indireto.

No BNDES indireto, a coluna **Juros** (% a.a.) é a remuneração do agente
financeiro credenciado (spread do banco). Em cada mês do cronograma SAC::

    spread_banco_mes = saldo_fiscal × ((1 + juros/100)^(1/12) − 1)

Este script lê a massa em ``dados\\*.xlsx`` (não precisa reler os CSV de
parcelas) e grava:

  - spread_banco_por_contrato.csv   (um linha por contrato — completo)
  - spread_banco_por_agente.xlsx    (agregado por instituição)
  - resumo_spread_banco.xlsx        (Por_Agente + Totais + amostra)

Uso (WinPython ContAgil)::

  python scripts\\spread_banco_contratos.py --massa-dados dados --pasta-saida saida ^
      --arquivo-fatores fator_acumulado_SELIC_TJLP_TLP.xlsx
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
import time
import types
from pathlib import Path

import pandas as pd

_SCRIPTS_DIR = Path(__file__).resolve().parent
ROOT = _SCRIPTS_DIR.parent
MARKER = "spread-banco-contratos-20260727a"


def _load_sibling(mod_name: str):
    full = f"scripts.{mod_name}"
    if full in sys.modules:
        return sys.modules[full]
    path = _SCRIPTS_DIR / f"{mod_name}.py"
    if not path.is_file():
        print(f"ERRO [{MARKER}]: falta {path}")
        b = (
            "https://raw.githubusercontent.com/cafla19791-tech/"
            "SEC--data-analysys/cursor/spread-banco-contratos-f342"
        )
        print(f'  Invoke-WebRequest "{b}/scripts/{mod_name}.py" -OutFile scripts\\{mod_name}.py')
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
_seguro = _load_sibling("contagil_fluxos_seguro")

CONTAGIL_PASTA_DADOS = _gf.CONTAGIL_PASTA_DADOS
CONTAGIL_PASTA_SAIDA = _gf.CONTAGIL_PASTA_SAIDA
DATA_IMPACTO = _gf.DATA_IMPACTO
OUTPUT_DIR = _gf.OUTPUT_DIR
calcular_spread_banco_contrato = _gf.calcular_spread_banco_contrato
normalizar_colunas = _gf.normalizar_colunas
carregar_fatores_mensais = _seguro.carregar_fatores_mensais
listar_contratos = _seguro.listar_contratos
_ler_bruto_com_header = _seguro._ler_bruto_com_header

EXCEL_MAX_AMOSTRAS = 500_000


def processar_arquivo(
    arquivo: Path,
    *,
    selic_serie,
    csv_path: Path,
    header_escrito: bool,
    max_contratos: int | None = None,
) -> tuple[pd.DataFrame, bool, dict]:
    """Processa um Excel de contratos; anexa linhas ao CSV e retorna amostra + stats."""
    print(f"\n>>> {arquivo.name} ...")
    sys.stdout.flush()
    try:
        bruto = _ler_bruto_com_header(arquivo)
        df = normalizar_colunas(bruto)
    except Exception as exc:  # noqa: BLE001
        print(f"    ERRO ao ler/normalizar: {exc}")
        return pd.DataFrame(), header_escrito, {"contratos": 0, "ok": 0}

    if max_contratos is not None:
        df = df.head(int(max_contratos))

    n = len(df)
    print(f"    Contratos válidos: {n:,}")
    sys.stdout.flush()

    rows: list[dict] = []
    t0 = time.time()
    for i, row in enumerate(df.itertuples(index=False), start=1):
        try:
            res = calcular_spread_banco_contrato(
                pd.Timestamp(row.data_contratacao),
                float(row.valor_desembolsado),
                int(row.prazo_carencia),
                int(row.prazo_amortizacao),
                float(row.juros),
                selic_serie=selic_serie,
            )
        except Exception:  # noqa: BLE001
            continue

        rows.append(
            {
                "arquivo_origem": arquivo.name,
                "contrato": getattr(row, "contrato", i - 1),
                "Instituição Financeira": getattr(row, "agente", ""),
                "data_contratacao": pd.Timestamp(row.data_contratacao).date(),
                "valor_desembolsado": round(float(row.valor_desembolsado), 2),
                "custo_financeiro": str(getattr(row, "custo_financeiro", "") or ""),
                "juros_aa_pct": res["juros_aa_pct"],
                "taxa_spread_banco_mensal": res["taxa_spread_banco_mensal"],
                "prazo_carencia": int(row.prazo_carencia),
                "prazo_amortizacao": int(row.prazo_amortizacao),
                "parcelas": res["parcelas"],
                "spread_banco_nominal": res["spread_banco_nominal"],
                "spread_banco_2026": res["spread_banco_2026"],
            }
        )
        if i % 5_000 == 0 or i == n:
            elapsed = max(time.time() - t0, 1e-6)
            rate = i / elapsed
            eta = (n - i) / rate if rate > 0 else 0.0
            print(
                f"    {i:,}/{n:,} ({100.0 * i / n:.1f}%) | "
                f"{rate:,.0f} contr/s | ETA ~{eta / 60:.1f} min"
            )
            sys.stdout.flush()

    if not rows:
        print("    Nenhum contrato processado.")
        return pd.DataFrame(), header_escrito, {"contratos": n, "ok": 0}

    out = pd.DataFrame(rows)
    out.to_csv(csv_path, mode="a", index=False, header=not header_escrito)
    header_escrito = True

    stats = {
        "contratos": n,
        "ok": len(out),
        "spread_nominal": float(out["spread_banco_nominal"].sum()),
        "spread_2026": float(out["spread_banco_2026"].sum()),
    }
    print(
        f"    → +{len(out):,} contratos | "
        f"spread nominal R$ {stats['spread_nominal']:,.2f} | "
        f"spread 2026 R$ {stats['spread_2026']:,.2f}"
    )
    # Amostra para Excel (prioriza maiores spreads)
    amostra = out.nlargest(min(len(out), 50_000), "spread_banco_2026")
    return amostra, header_escrito, stats


def agregar_por_agente(csv_path: Path, chunksize: int = 200_000) -> pd.DataFrame:
    """Agrega spread por agente a partir do CSV completo (streaming)."""
    acc: dict[str, list] = {}
    for chunk in pd.read_csv(csv_path, chunksize=chunksize):
        g = chunk.groupby("Instituição Financeira", dropna=False).agg(
            qtd_contratos=("contrato", "count"),
            valor_desembolsado=("valor_desembolsado", "sum"),
            spread_nominal=("spread_banco_nominal", "sum"),
            spread_2026=("spread_banco_2026", "sum"),
            juros_medio=("juros_aa_pct", "mean"),
        )
        for ag, row in g.iterrows():
            key = str(ag) if pd.notna(ag) and str(ag).strip() else "Não informado"
            if key not in acc:
                acc[key] = [0, 0.0, 0.0, 0.0, 0.0, 0]  # qtd, valor, nom, 2026, juros*qtd, qtd
            a = acc[key]
            q = int(row["qtd_contratos"])
            a[0] += q
            a[1] += float(row["valor_desembolsado"])
            a[2] += float(row["spread_nominal"])
            a[3] += float(row["spread_2026"])
            a[4] += float(row["juros_medio"]) * q
            a[5] += q

    rows = []
    for ag, a in acc.items():
        rows.append(
            {
                "Instituição Financeira": ag,
                "Qtd Contratos": a[0],
                "Valor Desembolsado (R$)": round(a[1], 2),
                "Spread Banco Nominal (R$)": round(a[2], 2),
                "Spread Banco 2026 (R$)": round(a[3], 2),
                "Juros Médio (% a.a.)": round(a[4] / a[5], 4) if a[5] else 0.0,
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out.sort_values("Spread Banco 2026 (R$)", ascending=False).reset_index(drop=True)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "--massa-dados",
        type=Path,
        default=None,
        help="Pasta com Excel de contratos (default: ContAgil dados/).",
    )
    p.add_argument(
        "--pasta-saida",
        type=Path,
        default=None,
        help="Pasta de saída (default: ContAgil saida/).",
    )
    p.add_argument(
        "--arquivo-fatores",
        type=Path,
        default=None,
        help="Excel fator_acumulado_SELIC_TJLP_TLP.xlsx (capitalização 2026).",
    )
    p.add_argument(
        "--max-contratos",
        type=int,
        default=None,
        help="Limite por arquivo (teste).",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print(f"[spread_banco_contratos {MARKER}]")
    print("=" * 70)
    print("SPREAD DO BANCO POR CONTRATO — BNDES INDIRETOS")
    print(f"Referência de capitalização: {DATA_IMPACTO:%d/%m/%Y}")
    print("=" * 70)

    massa = Path(args.massa_dados) if args.massa_dados else CONTAGIL_PASTA_DADOS
    if not massa.exists() and (ROOT / "dados").exists():
        massa = ROOT / "dados"
    saida = Path(args.pasta_saida) if args.pasta_saida else CONTAGIL_PASTA_SAIDA
    if not saida.exists() and args.pasta_saida is None:
        saida = OUTPUT_DIR
    saida.mkdir(parents=True, exist_ok=True)

    if not massa.exists():
        print(f"Pasta de dados não encontrada: {massa}", file=sys.stderr)
        return 1

    print(f"Massa de dados : {massa}")
    print(f"Pasta de saída : {saida}")

    selic_serie = None
    if args.arquivo_fatores is not None:
        print(f"Carregando fatores: {args.arquivo_fatores}")
        selic_serie = carregar_fatores_mensais(Path(args.arquivo_fatores))
    else:
        # Tenta fator na pasta winpython / massa
        for cand in (
            massa.parent / "fator_acumulado_SELIC_TJLP_TLP.xlsx",
            ROOT / "fator_acumulado_SELIC_TJLP_TLP.xlsx",
            massa / "fator_acumulado_SELIC_TJLP_TLP.xlsx",
        ):
            if cand.is_file():
                print(f"Carregando fatores: {cand}")
                selic_serie = carregar_fatores_mensais(cand)
                break
        if selic_serie is None:
            print(
                "Aviso: sem arquivo de fatores — capitalização com SELIC 14,5%/12.",
                file=sys.stderr,
            )

    arquivos = listar_contratos(massa)
    if not arquivos:
        print(f"Nenhum Excel de contratos em {massa}", file=sys.stderr)
        return 1
    print(f"Arquivos de contratos ({len(arquivos)}):")
    for a in arquivos:
        print(f"  - {a.name}")

    csv_path = saida / "spread_banco_por_contrato.csv"
    if csv_path.exists():
        csv_path.unlink()

    header = False
    amostras: list[pd.DataFrame] = []
    tot_ok = 0
    tot_nom = 0.0
    tot_2026 = 0.0
    t0 = time.time()

    for arq in arquivos:
        amostra, header, stats = processar_arquivo(
            arq,
            selic_serie=selic_serie,
            csv_path=csv_path,
            header_escrito=header,
            max_contratos=args.max_contratos,
        )
        if not amostra.empty:
            amostras.append(amostra)
        tot_ok += int(stats.get("ok", 0))
        tot_nom += float(stats.get("spread_nominal", 0.0))
        tot_2026 += float(stats.get("spread_2026", 0.0))

    if tot_ok == 0:
        print("Nenhum contrato processado.", file=sys.stderr)
        return 1

    print("\nAgregando por agente (streaming do CSV)...")
    por_agente = agregar_por_agente(csv_path)
    ag_xlsx = saida / "spread_banco_por_agente.xlsx"
    ag_csv = saida / "spread_banco_por_agente.csv"
    por_agente.to_excel(ag_xlsx, index=False)
    por_agente.to_csv(ag_csv, index=False)

    totais = pd.DataFrame(
        [
            {"Métrica": "Contratos", "Valor": tot_ok},
            {"Métrica": "Spread Banco Nominal (R$)", "Valor": round(tot_nom, 2)},
            {"Métrica": "Spread Banco 2026 (R$)", "Valor": round(tot_2026, 2)},
            {
                "Métrica": "Referência",
                "Valor": DATA_IMPACTO.strftime("%d/%m/%Y"),
            },
            {"Métrica": "Arquivos", "Valor": len(arquivos)},
            {"Métrica": "Tempo (s)", "Valor": round(time.time() - t0, 1)},
        ]
    )

    amostra_all = (
        pd.concat(amostras, ignore_index=True)
        .nlargest(EXCEL_MAX_AMOSTRAS, "spread_banco_2026")
        if amostras
        else pd.DataFrame()
    )

    wb = saida / "resumo_spread_banco.xlsx"
    with pd.ExcelWriter(wb, engine="openpyxl") as writer:
        por_agente.to_excel(writer, sheet_name="Por_Agente", index=False)
        totais.to_excel(writer, sheet_name="Totais", index=False)
        if not amostra_all.empty:
            amostra_all.to_excel(writer, sheet_name="Amostra_Contratos", index=False)

    print()
    print("=" * 70)
    print("TOTAIS — SPREAD DO BANCO")
    print("=" * 70)
    print(f"  Contratos              : {tot_ok:,}")
    print(f"  Spread nominal         : R$ {tot_nom:,.2f}")
    print(f"  Spread capitalizado 2026: R$ {tot_2026:,.2f}")
    print(f"  Tempo                  : {time.time() - t0:,.1f} s")
    print()
    print("Top 15 agentes (Spread Banco 2026):")
    print(por_agente.head(15).to_string(index=False))
    print()
    print("Arquivos gerados:")
    print(f"  → {csv_path}  (completo, {tot_ok:,} contratos)")
    print(f"  → {ag_xlsx}")
    print(f"  → {ag_csv}")
    print(f"  → {wb}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
