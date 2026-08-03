#!/usr/bin/env python3
"""
ContAgil — geração de fluxos SAC com carência (script estilo WinPython).

Equivalente ao rascunho ContAgil de lotes/Excel, com a lógica já corrigida:
  - cronograma carência + n (dia 15)
  - taxa_contrato_efetiva (TAXA FIXA / TJLP/TLP)
  - dual balance (saldo_fiscal / saldo_contrato)
  - subsídio + impacto ContAgil (fatores SELIC col E, +1 dia → 30/06/2026)

Uso (ContAgil / WinPython):
  python3 scripts/gerar_fluxo_sac.py \\
    --excel "C:\\caminho\\para\\operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx" \\
    --pasta-saida "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
    --arquivo-selic "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\STP-20260716182715078 (1).xlsx"

Sem WinPython local: usa data/ + output/ e CSV BNDES 2009–2010 (ou --download).
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from scripts.gerar_fluxos import (
    CONTAGIL_PASTA_SAIDA,
    CONTAGIL_SELIC_DEFAULT,
    DATA_DIR,
    FILTERED_CSV,
    OUTPUT_DIR,
    carregar_selic_serie,
    download_and_filter_csv,
    load_from_csv,
    load_from_excel,
    processar_em_lotes,
    resolver_arquivo_selic,
    resumo_from_agent_agg,
)

CONTAGIL_SAIDA_DEFAULT = Path(
    r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida"
)
LOCAL_SAIDA = DATA_DIR / "contagil_winpython" / "saida"
VALOR_CONTRATO0_PADRAO = 485_000.0


def _parece_caminho_contagil(path: Path | None) -> bool:
    if path is None:
        return False
    texto = str(path).replace("/", "\\").upper()
    return "CONTAGIL" in texto or "WINPYTHON" in texto or texto.startswith("C:\\ARQUIVOS")


def _resolver_pasta_saida(pasta: Path | None) -> Path:
    if pasta is None:
        if CONTAGIL_PASTA_SAIDA.exists():
            return CONTAGIL_PASTA_SAIDA
        if CONTAGIL_SAIDA_DEFAULT.exists():
            return CONTAGIL_SAIDA_DEFAULT
        return OUTPUT_DIR
    if not pasta.exists() and _parece_caminho_contagil(pasta):
        print(
            f"⚠️ Pasta ContAgil de saída ausente: {pasta}\n"
            f"   Usando espelho local: {LOCAL_SAIDA}"
        )
        return LOCAL_SAIDA
    return pasta


def _aplicar_correcao_contrato0(df: pd.DataFrame, valor: float | None) -> pd.DataFrame:
    """Correção ContAgil: valor desembolsado do contrato 0."""
    if valor is None or df.empty or "contrato" not in df.columns:
        return df
    out = df.copy()
    mask = out["contrato"] == 0
    if mask.any():
        antigo = float(out.loc[mask, "valor_desembolsado"].iloc[0])
        out.loc[mask, "valor_desembolsado"] = float(valor)
        print(
            f"Correção contrato 0: valor_desembolsado {antigo:,.2f} → {float(valor):,.2f}"
        )
    return out


def _carregar_contratos(args: argparse.Namespace) -> pd.DataFrame:
    if args.excel is not None:
        path = Path(args.excel)
        if not path.exists():
            print(f"⚠️ Excel não encontrado: {path}")
            print("   Tentando CSV filtrado / download BNDES 2009–2010...")
        else:
            print(f"Lendo Excel: {path}")
            return load_from_excel(path, header=args.excel_header)

    if args.input is not None:
        print(f"Lendo CSV: {args.input}")
        return load_from_csv(Path(args.input))

    if FILTERED_CSV.exists() and not args.download:
        print(f"Lendo cache: {FILTERED_CSV}")
        return load_from_csv(FILTERED_CSV)

    if args.download or args.excel is not None:
        path = download_and_filter_csv(start=args.start, end=args.end)
        return load_from_csv(path)

    sample = DATA_DIR / "sample_operacoes_com_agente.csv"
    if sample.exists():
        print(f"Lendo amostra: {sample}")
        return load_from_csv(sample)

    raise FileNotFoundError(
        "Nada para processar. Informe --excel, --input ou --download."
    )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--excel",
        type=Path,
        default=None,
        help=(
            "Excel portal/ContAgil "
            "(ex.: operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx)."
        ),
    )
    p.add_argument("--input", type=Path, help="CSV filtrado (sep=;).")
    p.add_argument(
        "--download",
        action="store_true",
        help="Baixa/filtra CSV aberto BNDES 2009–2010.",
    )
    p.add_argument("--start", default="2009-01-01")
    p.add_argument("--end", default="2010-12-31")
    p.add_argument(
        "--pasta-saida",
        type=Path,
        default=CONTAGIL_SAIDA_DEFAULT,
        help="Pasta ContAgil de saída (default: WinPython/saida).",
    )
    p.add_argument(
        "--arquivo-selic",
        type=Path,
        default=None,
        help=f"STP ContAgil (default auto: {CONTAGIL_SELIC_DEFAULT.name} / Bacen).",
    )
    p.add_argument("--stem", default="fluxos_completos_corrigido")
    p.add_argument("--lote", type=int, default=2000)
    p.add_argument("--max-contratos", type=int, default=None)
    p.add_argument(
        "--excel-header",
        type=int,
        default=None,
        help="Linha do header (None=auto: 0 ContAgil ou 5 portal).",
    )
    p.add_argument(
        "--corrigir-contrato0",
        type=float,
        nargs="?",
        const=VALOR_CONTRATO0_PADRAO,
        default=None,
        help=(
            f"Corrige valor desembolsado do contrato 0 "
            f"(default se flag sem valor: {VALOR_CONTRATO0_PADRAO:,.0f})."
        ),
    )
    p.add_argument(
        "--sem-selic-fatores",
        action="store_true",
        help="Usa SELIC 14,5%% composta constante.",
    )
    p.add_argument(
        "--baixar-selic",
        action="store_true",
        help="Força Bacen SGS 11 se não houver STP.",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    pasta_saida = _resolver_pasta_saida(args.pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Diretório de saída configurado: {pasta_saida}")

    # SELIC ContAgil: não falhar se o path Windows do STP não existir
    selic_arg = args.arquivo_selic
    if selic_arg is not None and not Path(selic_arg).exists():
        print(f"⚠️ Arquivo SELIC não encontrado: {selic_arg}")
        print("   Tentando auto-descoberta ContAgil/data/Bacen...")
        selic_arg = resolver_arquivo_selic(None)

    ns = argparse.Namespace(
        arquivo_selic=selic_arg,
        baixar_selic=args.baixar_selic or not args.sem_selic_fatores,
        sem_selic_fatores=args.sem_selic_fatores,
    )
    if args.sem_selic_fatores:
        ns.baixar_selic = False
    selic_serie = carregar_selic_serie(ns)

    df = _carregar_contratos(args)
    print(f"Total de contratos carregados: {len(df):,}")

    df = _aplicar_correcao_contrato0(df, args.corrigir_contrato0)

    if args.max_contratos is not None:
        df = df.head(args.max_contratos).copy()
        df["contrato"] = df.index
        print(f"Limitado a {len(df):,} contratos (--max-contratos)")

    print("Gerando fluxos completos (SAC + carência corrigida)...")
    # Grava direto na pasta ContAgil (evita sobrescrever output/resumo do run completo)
    csv_path = pasta_saida / f"{args.stem}.csv"
    stats = processar_em_lotes(
        df,
        csv_path,
        lote=args.lote,
        selic_serie=selic_serie,
    )
    resumo = resumo_from_agent_agg(stats.get("por_agente", {}))
    dest_xlsx_resumo = pasta_saida / "resumo_por_agente.xlsx"
    dest_csv_resumo = pasta_saida / "resumo_por_agente.csv"
    resumo.to_csv(dest_csv_resumo, index=False)
    resumo.to_excel(dest_xlsx_resumo, index=False, sheet_name="Por_Agente")

    # Espelho opcional em output/ (mesmo stem; não toca resumo_por_agente global)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(csv_path, OUTPUT_DIR / f"{args.stem}.csv")

    print(f"✅ Fluxos salvos em: {csv_path}")
    print(f"Tamanho: {stats['n_parcelas']:,} linhas")
    print(f"✅ Resumo por agente salvo em: {dest_xlsx_resumo}")
    print(
        f"   Contratos={stats['n_contratos_ok']:,}  "
        f"Subsídio={stats['total_subsidio']:,.2f}  "
        f"Impacto={stats['total_impacto_fiscal_2026']:,.2f}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
