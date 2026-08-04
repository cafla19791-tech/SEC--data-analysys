#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Entrypoint ContAgil/RFB (WinPython) - calculo de fluxos BNDES indiretos.

Este arquivo e PYTHON. Nao cole aqui o conteudo de contagil_fluxos_bndes.bat
(linhas REM / @echo off). O .bat fica na raiz do projeto e so chama este script.

Uso (uma linha, sem ^):
  python scripts/contagil_fluxos.py --massa-dados dados --pasta-saida saida --arquivo-fatores fator_acumulado_SELIC_TJLP_TLP.xlsx
"""

from __future__ import annotations

import argparse
import glob
import importlib.util
import os
import sys
import types
from datetime import datetime
from pathlib import Path

# WinPython: carrega irmãos por caminho de arquivo (nao depende de pacote scripts).
_SCRIPTS_DIR = Path(__file__).resolve().parent
ROOT = _SCRIPTS_DIR.parent
_CONTAGIL_BUILD = "importlib-20260725c-progresso-lotes"


def _load_sibling(mod_name: str):
    """Carrega ``scripts/<mod_name>.py`` via importlib (ContAgil/WinPython)."""
    full = f"scripts.{mod_name}"
    if full in sys.modules:
        return sys.modules[full]
    path = _SCRIPTS_DIR / f"{mod_name}.py"
    if not path.is_file():
        print(f"ERRO [{_CONTAGIL_BUILD}]: falta o arquivo:")
        print(f"  {path}")
        print("No PowerShell (so isto):")
        b = (
            "https://raw.githubusercontent.com/cafla19791-tech/"
            "SEC--data-analysys/cursor/normalizar-colunas-6f97"
        )
        print(f'  $b="{b}"')
        print(f'  Invoke-WebRequest "$b/scripts/{mod_name}.py" -OutFile scripts\\{mod_name}.py')
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

import pandas as pd

CONTAGIL_PASTA_DADOS = _gf.CONTAGIL_PASTA_DADOS
CONTAGIL_PASTA_SAIDA = _gf.CONTAGIL_PASTA_SAIDA
CONTAGIL_SELIC_DEFAULT = _gf.CONTAGIL_SELIC_DEFAULT
CONTAGIL_WINPYTHON = _gf.CONTAGIL_WINPYTHON
DATA_DIR = _gf.DATA_DIR
DATA_IMPACTO = _gf.DATA_IMPACTO
OUTPUT_DIR = _gf.OUTPUT_DIR
SelicSerie = _gf.SelicSerie
calcular_impacto_fiscal_real = _gf.calcular_impacto_fiscal_real
carregar_selic_serie = _gf.carregar_selic_serie
gerar_fluxos = _gf.gerar_fluxos
load_from_csv = _gf.load_from_csv
load_from_excel = _gf.load_from_excel
gerar_fluxos_main = _gf.main
normalizar_colunas = _gf.normalizar_colunas
resolver_arquivo_selic = _gf.resolver_arquivo_selic
resolver_excel_operacoes = _gf.resolver_excel_operacoes

# Reexporta para scripts ContAgil que fazem ``from scripts.contagil_fluxos import normalizar_colunas``
__all__ = ["main", "normalizar_colunas", "parse_args", "carregar_selic"]


def listar_excels(pasta: Path) -> list[Path]:
    """Lista *.xlsx em pasta (equivale a glob ContAgil)."""
    return sorted(Path(p) for p in glob.glob(os.path.join(str(pasta), "*.xlsx")))


def teste_contrato0(serie: SelicSerie) -> float:
    """Validação ContAgil: subsidio=1886.11 em 15/02/2009."""
    FATOR_30_06_2026 = _gf.FATOR_30_06_2026

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
    carregar_fatores_mensais = _load_sibling("contagil_fluxos_seguro").carregar_fatores_mensais

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
    # Arquivos grandes: grava CSV em lotes (evita "travar" sem log / OOM).
    if len(df) >= 5_000 and not fluxo_diario:
        gerar_e_gravar = _gf.gerar_e_gravar_fluxos
        gerar_e_gravar(df, selic_arg, saida_xlsx=nome_saida, lote=2_000)
        return nome_saida

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
        "--arquivo-fatores",
        dest="arquivo_selic",
        type=Path,
        default=None,
        metavar="FILE",
        help=(
            "Fatores ContAgil: fator_acumulado_SELIC_TJLP_TLP.xlsx (mensal) "
            f"ou STP (col D). Aliases: --fatores, --arquivo-fatores. Default auto: "
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
    # Identifica build no log (confirma que o .py atualizado foi baixado)
    print(f"[contagil_fluxos {_CONTAGIL_BUILD}]")

    # ContAgil WinPython: --massa-dados + --arquivo-fatores/--fatores mensal
    # → pipeline seguro (define normalizar_colunas + banner BNDES INDIRETOS)
    mensal = (
        args.arquivo_selic is not None
        and _parece_fatores_mensais(Path(args.arquivo_selic))
        and not args.teste_contrato0
        and not args.download
        and not args.fluxo_diario
        and args.excel is None
    )
    if mensal and (args.pasta_dados is not None or args.input is not None):
        seguro_main = _load_sibling("contagil_fluxos_seguro").main

        cli: list[str] = ["--arquivo-fatores", str(args.arquivo_selic)]
        if args.pasta_dados is not None:
            cli += ["--massa-dados", str(args.pasta_dados)]
        if args.pasta_saida is not None:
            cli += ["--pasta-saida", str(args.pasta_saida)]
        if args.max_contratos is not None:
            cli += ["--max-contratos", str(args.max_contratos)]
        if args.input is not None:
            cli += ["--input", str(args.input)]
        return seguro_main(cli)

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
        # Fatores mensais já carregados (auto-descoberta): pipeline seguro
        if serie is not None and "mensal:" in (serie.origem or ""):
            seguro_main = _load_sibling("contagil_fluxos_seguro").main

            cli = [
                "--massa-dados",
                str(pasta_dados),
                "--pasta-saida",
                str(pasta_saida),
            ]
            if args.arquivo_selic is not None:
                cli += ["--arquivo-fatores", str(args.arquivo_selic)]
            if args.max_contratos is not None:
                cli += ["--max-contratos", str(args.max_contratos)]
            return seguro_main(cli)

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
