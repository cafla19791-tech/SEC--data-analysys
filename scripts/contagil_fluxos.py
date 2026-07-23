#!/usr/bin/env python3
"""
Entrypoint no estilo do script ContAgil/RFB (WinPython):

  massa_dados / pasta_dados = .../python_jep/winpython/dados
  pasta_saida               = .../python_jep/winpython/saida
  arquivo_selic / --fatores = STP-*.xlsx  OU  fator_acumulado_SELIC_TJLP_TLP.xlsx

Processa todos os .xlsx da massa de dados, gera fluxos (SAC + carência corrigida)
e impacto fiscal ContAgil (fatores mensais ou col D STP).

Correções vs script ContAgil original (colado/corrompido):
  - sintaxe Python válida (True/values/method='nearest'/etc.)
  - gerar_fluxos(df, df) = df_original (instituição); gerar_fluxos(df, selic) = fatores
  - carência: cronograma cobre (carência + n) meses
  - taxa_contrato_efetiva: TAXA FIXA / TJLP/TLP com composição mensal
  - dual balance: saldo_fiscal (principal) + saldo_contrato (com juros)
  - fator SELIC mensal (--fatores) ou col D STP; idx = nearest(data_parcela)

Uso (ContAgil / WinPython) — capitalização mensal:
  python3 scripts/contagil_fluxos.py \\
    --massa-dados "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\dados" \\
    --pasta-saida "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
    --fatores "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\fator_acumulado_SELIC_TJLP_TLP.xlsx"

  # STP diário (legado ContAgil col D):
  python3 scripts/contagil_fluxos.py \\
    --massa-dados "...\\dados" --pasta-saida "...\\saida" \\
    --arquivo-selic "...\\STP-20260716182715078 (1).xlsx"

  # Sem args: usa defaults WinPython se existirem
  python3 scripts/contagil_fluxos.py

  # Mesmo comando ContAgil sem WinPython local (Linux/cloud):
  # cai para data/contagil_winpython/{dados,saida} + SELIC Bacen

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
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from scripts.gerar_fluxos import (
    CONTAGIL_PASTA_DADOS,
    CONTAGIL_PASTA_SAIDA,
    CONTAGIL_SELIC_DEFAULT,
    CONTAGIL_WINPYTHON,
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
    from scripts.gerar_fluxos import FATOR_30_06_2026

    subsidio = 1886.11
    data_parcela = datetime(2009, 2, 15)
    impacto = calcular_impacto_fiscal_real(subsidio, data_parcela, serie)
    print(f"Contrato 0 — subsidio={subsidio} data={data_parcela.date()}")
    print(f"Impacto Fiscal (fatores ContAgil col D): R$ {impacto:,.2f}")
    idx = serie.idx_proximo(data_parcela)
    fator_parcela = float(serie.fatores[idx])
    fator_fim = (
        float(serie.fator_referencia)
        if serie.fator_referencia is not None
        else float(serie.fatores[serie.idx_proximo(DATA_IMPACTO)])
    )
    if fator_parcela > 0:
        fator = fator_fim / fator_parcela
        print(
            f"  fator = {fator:.6f}  (parcela={fator_parcela:.5f} → "
            f"ref={fator_fim:.5f}; FATOR_30_06_2026={FATOR_30_06_2026})"
        )
    return impacto


def _parece_fatores_mensais(path: Path) -> bool:
    """True se o Excel é fator_acumulado_SELIC_TJLP_TLP (capitalização mensal)."""
    nome = path.name.lower()
    if "fator_acumulado" in nome or "selic_tjlp" in nome or "selic_mensal" in nome:
        return True
    try:
        cols = {str(c).strip().lower() for c in pd.read_excel(path, nrows=0).columns}
    except Exception:  # noqa: BLE001
        return False
    return "fator_acumulado" in cols and (
        any("taxa" in c for c in cols) or "data" in cols
    )


def carregar_selic(arquivo_selic: Path | None, baixar: bool) -> SelicSerie | None:
    """Carrega fatores mensais (--fatores), STP ContAgil (col D) ou Bacen."""
    from scripts.contagil_fluxos_seguro import carregar_fatores_mensais

    # Preferência: caminho explícito → auto ContAgil/data → Bacen se baixar
    if arquivo_selic is not None:
        caminho = Path(arquivo_selic)
        if caminho.exists():
            if _parece_fatores_mensais(caminho):
                return carregar_fatores_mensais(caminho)
            serie = SelicSerie.from_excel(caminho)
            print(f"SELIC ContAgil (col D): {caminho} ({len(serie.datas):,} pontos)")
            return serie
        # Caminho ContAgil explícito ausente: tenta auto-descoberta antes de falhar
        print(f"⚠️ Arquivo de fatores/SELIC não encontrado: {caminho}")
        print("   Tentando auto-descoberta ContAgil/data/Bacen...")

    # Default ContAgil: fator mensal na pasta winpython, senão STP
    for cand in (
        Path.cwd() / "fator_acumulado_SELIC_TJLP_TLP.xlsx",
        CONTAGIL_WINPYTHON / "fator_acumulado_SELIC_TJLP_TLP.xlsx",
        DATA_DIR / "fator_acumulado_SELIC_TJLP_TLP.xlsx",
        DATA_DIR / "selic_mensal.xlsx",
    ):
        if cand.exists() and _parece_fatores_mensais(cand):
            return carregar_fatores_mensais(cand)

    resolvido = resolver_arquivo_selic(None)
    if resolvido is not None:
        if _parece_fatores_mensais(resolvido):
            return carregar_fatores_mensais(resolvido)
        serie = SelicSerie.from_excel(resolvido)
        print(f"SELIC ContAgil (col D): {resolvido} ({len(serie.datas):,} pontos)")
        return serie

    if arquivo_selic is not None and not baixar:
        raise FileNotFoundError(
            f"Arquivo de fatores/SELIC não encontrado: {arquivo_selic}\n"
            "Informe --fatores (fator_acumulado_SELIC_TJLP_TLP.xlsx) "
            "ou --arquivo-selic (STP col D), ou use --baixar-selic."
        )

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
    fluxo_diario: bool = False,
) -> Path:
    """Gera fluxos de um Excel e grava fluxos_<basename>.xlsx (script ContAgil)."""
    print(f"Processando: {arquivo}")
    df = load_from_excel(arquivo, header=header)
    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_saida = pasta_saida / f"fluxos_{arquivo.name}"
    if nome_saida.suffix.lower() != ".xlsx":
        nome_saida = nome_saida.with_suffix(".xlsx")
    saida_diario = (
        pasta_saida / f"fluxos_diarios_{arquivo.stem}.xlsx" if fluxo_diario else None
    )
    selic_arg = selic_serie if selic_serie is not None else 0.145
    df_fluxos = gerar_fluxos(
        df,
        selic_arg,
        fluxo_diario=fluxo_diario,
        saida_diario=saida_diario,
    )
    df_fluxos.to_excel(nome_saida, index=False)
    print(f"  → Salvo: {nome_saida} ({len(df_fluxos):,} parcelas)")
    if saida_diario is not None and Path(saida_diario).exists():
        print(f"  → Diário: {saida_diario}")
    return nome_saida


def _parece_caminho_contagil(path: Path | None) -> bool:
    """True se o caminho parece o WinPython ContAgil/RFB (Windows)."""
    if path is None:
        return False
    texto = str(path).replace("/", "\\").upper()
    return "CONTAGIL" in texto or "WINPYTHON" in texto or texto.startswith("C:\\ARQUIVOS")


def preparar_massa_local_fallback() -> Path:
    """
    Monta massa local em data/contagil_winpython/dados a partir da amostra do repo.

    Usado quando --massa-dados aponta para o WinPython ContAgil e a pasta
    não existe neste ambiente (Linux/cloud sem instalação RFB).
    """
    pasta = DATA_DIR / "contagil_winpython" / "dados"
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / "sample_operacoes_com_agente.xlsx"
    sample_csv = DATA_DIR / "sample_operacoes_com_agente.csv"
    if not destino.exists():
        if not sample_csv.exists():
            raise FileNotFoundError(
                "Massa ContAgil ausente e amostra local não encontrada "
                f"({sample_csv}). Envie os .xlsx para --massa-dados ou use --download."
            )
        df = load_from_csv(sample_csv)
        # Layout ContAgil (header PT na 1ª linha) para load_from_excel
        pd.DataFrame(
            {
                "Data da contratação": df["data_contratacao"].dt.strftime("%d/%m/%Y"),
                "Valor Desembolsado R$ (*)": df["valor_desembolsado"].astype(float),
                "Juros": df["juros"].astype(float),
                "Prazo - Carência (meses)": df["prazo_carencia"].astype(int),
                "Prazo - Amortização (meses)": df["prazo_amortizacao"].astype(int),
                "Instituição Financeira Credenciada": df["agente"],
                "Custo financeiro": df["custo_financeiro"],
            }
        ).to_excel(destino, index=False)
        print(f"Massa local gerada a partir da amostra: {destino}")
    return pasta


def processar_pasta_dados(
    pasta_dados: Path,
    pasta_saida: Path,
    selic_serie: SelicSerie | None,
    header: int | None = None,
    fluxo_diario: bool = False,
) -> list[Path]:
    """Processa todos os *.xlsx do diretório de dados (loop ContAgil)."""
    pasta_dados = Path(pasta_dados)
    if not pasta_dados.exists():
        raise FileNotFoundError(
            f"Massa de dados não encontrada: {pasta_dados}\n"
            "Use --massa-dados apontando para a pasta WinPython/dados do ContAgil."
        )
    if not pasta_dados.is_dir():
        raise NotADirectoryError(f"--massa-dados deve ser uma pasta: {pasta_dados}")

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
            saidas.append(
                processar_arquivo(
                    arquivo,
                    pasta_saida,
                    selic_serie,
                    header=header,
                    fluxo_diario=fluxo_diario,
                )
            )
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
        "--fatores",
        dest="arquivo_selic",
        type=Path,
        default=None,
        metavar="FILE",
        help=(
            "Fatores ContAgil: fator_acumulado_SELIC_TJLP_TLP.xlsx (mensal) "
            f"ou STP (col D). Alias: --fatores. Default auto: "
            f"{CONTAGIL_SELIC_DEFAULT.name} / Bacen."
        ),
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
    p.add_argument(
        "--fluxo-diario",
        action="store_true",
        help="Gera tabela detalhada dia a dia (fluxos_diarios_detalhados.xlsx).",
    )
    return p.parse_args(argv)


def _resolver_pastas(args: argparse.Namespace) -> tuple[Path | None, Path]:
    """Resolve pasta_dados / pasta_saida com defaults ContAgil → repo.

    Se --massa-dados/--pasta-saida apontarem para o WinPython ContAgil e as
    pastas não existirem (ambiente Linux/cloud), usa espelho local:
      data/contagil_winpython/dados  → amostra do repo
      data/contagil_winpython/saida  → saída dos fluxos
    """
    pasta_saida = args.pasta_saida
    if pasta_saida is None:
        if CONTAGIL_PASTA_SAIDA.exists():
            pasta_saida = CONTAGIL_PASTA_SAIDA
        else:
            pasta_saida = OUTPUT_DIR
    elif not pasta_saida.exists() and _parece_caminho_contagil(pasta_saida):
        pasta_saida = DATA_DIR / "contagil_winpython" / "saida"
        print(
            "⚠️ Pasta ContAgil de saída ausente neste ambiente.\n"
            f"   Usando espelho local: {pasta_saida}"
        )

    pasta_dados = args.pasta_dados
    if pasta_dados is None and args.excel is None and args.input is None:
        if CONTAGIL_PASTA_DADOS.exists():
            pasta_dados = CONTAGIL_PASTA_DADOS
        elif (DATA_DIR / "dados").exists():
            pasta_dados = DATA_DIR / "dados"
    elif (
        pasta_dados is not None
        and not pasta_dados.exists()
        and _parece_caminho_contagil(pasta_dados)
        and args.excel is None
        and args.input is None
    ):
        print(
            f"⚠️ Massa ContAgil não encontrada: {pasta_dados}\n"
            "   Ambiente sem WinPython RFB — gerando massa local da amostra do repo."
        )
        pasta_dados = preparar_massa_local_fallback()
        if args.pasta_saida is not None and not _parece_caminho_contagil(
            Path(args.pasta_saida)
        ):
            # Usuário pediu saida ContAgil ausente já redirecionada acima;
            # se pediu outra pasta existente, mantém.
            pass
        elif not pasta_saida.exists() or _parece_caminho_contagil(pasta_saida):
            pasta_saida = DATA_DIR / "contagil_winpython" / "saida"
            print(f"   Pasta de saída local: {pasta_saida}")

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

    pasta_dados, pasta_saida = _resolver_pastas(args)

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
        if args.fluxo_diario:
            cli.append("--fluxo-diario")
            cli += [
                "--saida-diario",
                str(pasta_saida / "fluxos_diarios_detalhados.xlsx"),
            ]
        return gerar_fluxos_main(cli)

    # Modo ContAgil: todos os arquivos do diretório de dados
    if pasta_dados is not None and args.excel is None and args.input is None:
        print(f"Massa de dados: {pasta_dados}")
        print(f"Pasta de saída: {pasta_saida}")
        saidas = processar_pasta_dados(
            pasta_dados,
            pasta_saida,
            serie,
            header=args.excel_header,
            fluxo_diario=args.fluxo_diario,
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
    pasta_saida.mkdir(parents=True, exist_ok=True)
    saida_diario = pasta_saida / "fluxos_diarios_detalhados.xlsx"
    if serie is not None:
        df_fluxos = gerar_fluxos(
            df,
            serie,
            fluxo_diario=args.fluxo_diario,
            saida_diario=saida_diario if args.fluxo_diario else None,
        )
    else:
        df_fluxos = gerar_fluxos(
            df,
            fluxo_diario=args.fluxo_diario,
            saida_diario=saida_diario if args.fluxo_diario else None,
        )

    if out.suffix.lower() != ".xlsx":
        out = out.with_suffix(".xlsx")
    df_fluxos.to_excel(out, index=False)
    print(f"✅ Concluído! → {out} ({len(df_fluxos):,} parcelas)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
