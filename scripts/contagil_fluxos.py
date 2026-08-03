#!/usr/bin/env python3
"""
Entrypoint no estilo do script ContAgil/RFB (WinPython):

  massa_dados / pasta_dados = .../python_jep/winpython/dados
  pasta_saida               = .../python_jep/winpython/saida
  arquivo_selic             = .../STP-20260716182715078 (1).xlsx

Processa todos os .xlsx da massa de dados, gera fluxos (SAC + carência corrigida)
e impacto fiscal ContAgil (col E, capitalização a partir do dia seguinte → 30/06/2026).

Correções vs script ContAgil original (colado/corrompido):
  - sintaxe Python válida (True/values/method='nearest'/etc.)
  - gerar_fluxos(df, selic) — não gerar_fluxos(df, df)
  - carência: cronograma cobre (carência + n) meses
  - TJLP/TLP: taxa_aa = 6% + juros do contrato
  - fator SELIC na coluna E; idx_inicio = nearest(data_parcela + 1 dia)

Uso (ContAgil / WinPython):
  python3 scripts/contagil_fluxos.py \\
    --massa-dados "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\dados" \\
    --pasta-saida "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
    --arquivo-selic "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\STP-20260716182715078 (1).xlsx"

  # Sem args: usa defaults WinPython se existirem
  python3 scripts/contagil_fluxos.py

  # Repo local: pasta data/ → output/
  python3 scripts/contagil_fluxos.py --pasta-dados data --pasta-saida output \\
      --input data/sample_operacoes_com_agente.csv

  # Um arquivo / download BNDES
  python3 scripts/contagil_fluxos.py --input data/sample_operacoes_com_agente.csv
  python3 scripts/contagil_fluxos.py --download
  python3 scripts/contagil_fluxos.py --teste-contrato0
"""

from __future__ import annotations

import argparse
import glob
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from scripts.gerar_fluxos import (
    CONTAGIL_PASTA_DADOS,
    CONTAGIL_PASTA_SAIDA,
    CONTAGIL_SELIC_DEFAULT,
    DATA_DIR,
    DATA_IMPACTO,
    OUTPUT_DIR,
    SelicSerie,
    calcular_impacto_fiscal_real,
    carregar_selic_serie,
    gerar_fluxos,
    load_from_csv,
    load_from_excel,
    main as gerar_fluxos_main,
    resolver_arquivo_selic,
    resolver_excel_operacoes,
)


def listar_excels(pasta: Path) -> list[Path]:
    """Lista *.xlsx em pasta (equivale a glob ContAgil)."""
    return sorted(Path(p) for p in glob.glob(os.path.join(str(pasta), "*.xlsx")))


def teste_contrato0(serie: SelicSerie) -> float:
    """Validação ContAgil: subsidio=1886.11 em 15/02/2009."""
    subsidio = 1886.11
    data_parcela = datetime(2009, 2, 15)
    impacto = calcular_impacto_fiscal_real(subsidio, data_parcela, serie)
    print(f"Contrato 0 — subsidio={subsidio} data={data_parcela.date()}")
    print(f"Impacto Fiscal (fatores ContAgil): R$ {impacto:,.2f}")
    data_proxima = data_parcela + timedelta(days=1)
    idx_inicio = serie.idx_proximo(data_proxima)
    idx_fim = serie.idx_proximo(DATA_IMPACTO)
    if idx_fim > idx_inicio:
        fator = serie.fatores[idx_fim] / serie.fatores[idx_inicio]
        print(f"  fator = {fator:.6f}  (idx {idx_inicio} → {idx_fim})")
    return impacto


def carregar_selic(arquivo_selic: Path | None, baixar: bool) -> SelicSerie | None:
    """Carrega STP ContAgil (col E) ou Bacen; espelha o script RFB."""
    # Preferência: caminho explícito → auto ContAgil/data → Bacen se baixar
    if arquivo_selic is not None and Path(arquivo_selic).exists():
        serie = SelicSerie.from_excel(Path(arquivo_selic))
        print(f"SELIC ContAgil: {arquivo_selic} ({len(serie.datas):,} pontos)")
        return serie

    resolvido = resolver_arquivo_selic(None)
    if resolvido is not None:
        serie = SelicSerie.from_excel(resolvido)
        print(f"SELIC ContAgil: {resolvido} ({len(serie.datas):,} pontos)")
        return serie

    if arquivo_selic is not None:
        raise FileNotFoundError(f"Arquivo SELIC não encontrado: {arquivo_selic}")

    return carregar_selic_serie(
        argparse.Namespace(
            arquivo_selic=None,
            baixar_selic=baixar,
            sem_selic_fatores=not baixar,
        )
    )


def processar_arquivo(
    arquivo: Path,
    pasta_saida: Path,
    selic_serie: SelicSerie | None,
    header: int | None = None,
) -> Path:
    """Gera fluxos de um Excel e grava fluxos_<basename>.xlsx (script ContAgil)."""
    print(f"Processando: {arquivo}")
    df = load_from_excel(arquivo, header=header)
    df_fluxos = gerar_fluxos(df, selic_serie if selic_serie is not None else 0.145)
    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_saida = pasta_saida / f"fluxos_{arquivo.name}"
    if nome_saida.suffix.lower() != ".xlsx":
        nome_saida = nome_saida.with_suffix(".xlsx")
    df_fluxos.to_excel(nome_saida, index=False)
    print(f"  → Salvo: {nome_saida} ({len(df_fluxos):,} parcelas)")
    return nome_saida


def processar_pasta_dados(
    pasta_dados: Path,
    pasta_saida: Path,
    selic_serie: SelicSerie | None,
    header: int | None = None,
) -> list[Path]:
    """Processa todos os *.xlsx do diretório de dados (loop ContAgil)."""
    arquivos = listar_excels(pasta_dados)
    if not arquivos:
        raise FileNotFoundError(f"Nenhum .xlsx em {pasta_dados}")

    saidas: list[Path] = []
    for arquivo in arquivos:
        # Ignora STP / SELIC se estiverem na mesma pasta
        nome = arquivo.name.upper()
        if nome.startswith("STP") or "SELIC" in nome:
            print(f"Ignorando série SELIC: {arquivo.name}")
            continue
        try:
            saidas.append(processar_arquivo(arquivo, pasta_saida, selic_serie, header))
        except Exception as exc:  # noqa: BLE001 — continua lote ContAgil
            print(f"  Ignorado ({exc})")
    return saidas


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--pasta-dados",
        "--massa-dados",
        dest="pasta_dados",
        type=Path,
        default=None,
        metavar="DIR",
        help=(
            "Massa de dados ContAgil: pasta com vários .xlsx "
            "(alias: --massa-dados). Default: WinPython/dados se existir, "
            "senão processa --input/--excel."
        ),
    )
    p.add_argument(
        "--pasta-saida",
        type=Path,
        default=None,
        help="Pasta de saída (default: WinPython/saida ou output/).",
    )
    p.add_argument("--excel", type=Path, help="Excel único de operações.")
    p.add_argument("--input", type=Path, help="CSV de operações (sep=';).")
    p.add_argument(
        "--download",
        action="store_true",
        help="Baixa/filtra CSV aberto BNDES 2009–2010 e processa em lotes.",
    )
    p.add_argument(
        "--arquivo-selic",
        type=Path,
        default=None,
        help=f"STP ContAgil (default auto: {CONTAGIL_SELIC_DEFAULT.name} / Bacen).",
    )
    p.add_argument("--stem", default="fluxos_completos_final")
    p.add_argument("--max-contratos", type=int, default=None)
    p.add_argument(
        "--excel-header",
        type=int,
        default=None,
        help="Linha do header no Excel (None=auto: 0 ContAgil ou 5 portal).",
    )
    p.add_argument(
        "--teste-contrato0",
        action="store_true",
        help="Só valida impacto ContAgil do contrato 0 (1886.11 @ 15/02/2009).",
    )
    p.add_argument(
        "--sem-selic-fatores",
        action="store_true",
        help="Usa SELIC 14,5%% composta constante (não ContAgil).",
    )
    p.add_argument(
        "--baixar-selic",
        action="store_true",
        help="Força download Bacen SGS 11 se não houver STP local.",
    )
    return p.parse_args(argv)


def _resolver_pastas(args: argparse.Namespace) -> tuple[Path | None, Path]:
    """Resolve pasta_dados / pasta_saida com defaults ContAgil → repo."""
    pasta_saida = args.pasta_saida
    if pasta_saida is None:
        if CONTAGIL_PASTA_SAIDA.exists():
            pasta_saida = CONTAGIL_PASTA_SAIDA
        else:
            pasta_saida = OUTPUT_DIR

    pasta_dados = args.pasta_dados
    if pasta_dados is None and args.excel is None and args.input is None:
        if CONTAGIL_PASTA_DADOS.exists():
            pasta_dados = CONTAGIL_PASTA_DADOS
        elif (DATA_DIR / "dados").exists():
            pasta_dados = DATA_DIR / "dados"
    return pasta_dados, pasta_saida


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print("🚀 Processando arquivos ContAgil (fluxos + impacto fiscal)...")

    baixar_selic = args.baixar_selic or not args.sem_selic_fatores
    if args.sem_selic_fatores:
        baixar_selic = False

    serie = None
    if not args.sem_selic_fatores:
        serie = carregar_selic(args.arquivo_selic, baixar=baixar_selic)

    if args.teste_contrato0:
        if serie is None:
            raise RuntimeError("--teste-contrato0 exige fatores SELIC.")
        teste_contrato0(serie)
        print("✅ Concluído!")
        return 0

    # Pipeline completo (lotes + resumo) via gerar_fluxos.main
    if args.download:
        cli = ["--stem", args.stem, "--download"]
        if args.arquivo_selic is not None:
            cli += ["--arquivo-selic", str(args.arquivo_selic)]
        if args.sem_selic_fatores:
            cli.append("--sem-selic-fatores")
        elif args.baixar_selic:
            cli.append("--baixar-selic")
        if args.max_contratos is not None:
            cli += ["--max-contratos", str(args.max_contratos)]
        return gerar_fluxos_main(cli)

    pasta_dados, pasta_saida = _resolver_pastas(args)

    # Modo ContAgil: todos os arquivos do diretório de dados
    if pasta_dados is not None and args.excel is None and args.input is None:
        saidas = processar_pasta_dados(
            pasta_dados, pasta_saida, serie, header=args.excel_header
        )
        if not saidas:
            print("Nenhum arquivo processado.", file=sys.stderr)
            return 1
        print(f"✅ Processamento de todos os arquivos concluído! ({len(saidas)} saídas)")
        return 0

    if args.excel:
        print(f"Lendo Excel: {args.excel}")
        df = load_from_excel(args.excel, header=args.excel_header)
        out = pasta_saida / f"fluxos_{args.excel.name}"
    elif args.input:
        print(f"Lendo CSV: {args.input}")
        df = load_from_csv(args.input)
        out = pasta_saida / f"{args.stem}.xlsx"
    else:
        excel = resolver_excel_operacoes()
        if excel is None:
            # Fallback: amostra do repo
            sample = DATA_DIR / "sample_operacoes_com_agente.csv"
            if sample.exists():
                print(f"Lendo amostra: {sample}")
                df = load_from_csv(sample)
                out = pasta_saida / f"{args.stem}.xlsx"
            else:
                raise FileNotFoundError(
                    "Nada para processar. Use --massa-dados/--pasta-dados, "
                    "--excel, --input ou --download."
                )
        else:
            print(f"Lendo Excel: {excel}")
            df = load_from_excel(excel, header=args.excel_header)
            out = pasta_saida / f"fluxos_{excel.name}"

    if args.max_contratos is not None:
        df = df.head(args.max_contratos).copy()
        df["contrato"] = df.index

    # Equivalente ContAgil corrigido: df_fluxos = gerar_fluxos(df, selic)
    if serie is not None:
        df_fluxos = gerar_fluxos(df, serie)
    else:
        df_fluxos = gerar_fluxos(df)

    pasta_saida.mkdir(parents=True, exist_ok=True)
    if out.suffix.lower() != ".xlsx":
        out = out.with_suffix(".xlsx")
    df_fluxos.to_excel(out, index=False)
    print(f"✅ Concluído! → {out} ({len(df_fluxos):,} parcelas)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
