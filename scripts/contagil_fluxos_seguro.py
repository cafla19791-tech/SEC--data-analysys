#!/usr/bin/env python3
"""
Cálculo de fluxos e impactos - BNDES Indiretos
Capitalização Mensal | Versão Segura

Uso (ContAgil / WinPython):
  python contagil_fluxos_seguro.py

  python contagil_fluxos_seguro.py \\
      --massa-dados "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\dados" \\
      --pasta-saida "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
      --fatores "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\fator_acumulado_SELIC_TJLP_TLP.xlsx"

Lê todos os .xlsx da massa de dados (ex.: BNDES INDIRETAS 2002.xlsx),
mapeia colunas de forma tolerante (acentos, aliases, header em linhas 0-8)
e gera fluxos com impacto capitalizado pela série mensal de fatores.
"""

from __future__ import annotations

import argparse
import glob
import os
import sys
from datetime import datetime
from pathlib import Path

_SCRIPTS = Path(__file__).resolve().parent
ROOT = _SCRIPTS.parent
for _p in (str(ROOT), str(_SCRIPTS)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import numpy as np
import pandas as pd


def _load_gerar_fluxos():
    """Importa gerar_fluxos do diretório irmão (sec_scripts ContAgil ou scripts/).

    Prefere ``scripts.gerar_fluxos`` no repo (uma só identidade de SelicSerie).
    No ContAgil WinPython, ``scripts`` colide com site-packages - aí carrega o
    arquivo irmão via importlib.
    """
    import importlib.util

    try:
        mod = __import__("scripts.gerar_fluxos", fromlist=["*"])
        # Garante que não pegamos outro pacote 'scripts' sem gerar_fluxos útil
        if hasattr(mod, "gerar_fluxos") and hasattr(mod, "SelicSerie"):
            return mod
    except ModuleNotFoundError:
        pass

    path = _SCRIPTS / "gerar_fluxos.py"
    if path.exists():
        spec = importlib.util.spec_from_file_location("sec_gerar_fluxos_seguro", path)
        if spec is not None and spec.loader is not None:
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise ModuleNotFoundError(
        f"Não achou {_SCRIPTS / 'gerar_fluxos.py'} nem scripts.gerar_fluxos"
    )


_flux = _load_gerar_fluxos()
CONTAGIL_PASTA_DADOS = _flux.CONTAGIL_PASTA_DADOS
CONTAGIL_PASTA_SAIDA = _flux.CONTAGIL_PASTA_SAIDA
CONTAGIL_WINPYTHON = _flux.CONTAGIL_WINPYTHON
DATA_DIR = _flux.DATA_DIR
DATA_IMPACTO = _flux.DATA_IMPACTO
OUTPUT_DIR = _flux.OUTPUT_DIR
SelicSerie = _flux.SelicSerie
_excel_tem_colunas_contratos = _flux._excel_tem_colunas_contratos
_mapear_colunas_contratos = _flux._mapear_colunas_contratos
gerar_fluxos = _flux.gerar_fluxos
load_from_excel = _flux.load_from_excel

FATORES_DEFAULT_NOME = "fator_acumulado_SELIC_TJLP_TLP.xlsx"
DATA_REF_DEFAULT = datetime(2026, 6, 1)


def _banner() -> None:
    print("=" * 70)
    print(" CÁLCULO DE FLUXOS E IMPACTOS - BNDES INDIRETOS")
    print(" Capitalização Mensal | Versão Segura")
    print("=" * 70)


def _parece_caminho_contagil(path: Path | None) -> bool:
    if path is None:
        return False
    texto = str(path).replace("/", "\\").upper()
    return "CONTAGIL" in texto or "WINPYTHON" in texto or texto.startswith("C:\\ARQUIVOS")


def resolver_pasta_dados(arg: Path | None) -> Path:
    if arg is not None:
        if arg.exists():
            return arg
        if _parece_caminho_contagil(arg):
            local = DATA_DIR / "contagil_winpython" / "dados"
            local.mkdir(parents=True, exist_ok=True)
            print(f"[AVISO] Massa ContAgil ausente: {arg}")
            print(f"   Usando espelho local: {local}")
            return local
        return arg
    if CONTAGIL_PASTA_DADOS.exists():
        return CONTAGIL_PASTA_DADOS
    # Script na pasta winpython: ./dados
    cwd_dados = Path.cwd() / "dados"
    if cwd_dados.exists():
        return cwd_dados
    local = DATA_DIR / "contagil_winpython" / "dados"
    local.mkdir(parents=True, exist_ok=True)
    return local


def resolver_pasta_saida(arg: Path | None) -> Path:
    if arg is not None:
        if arg.exists() or not _parece_caminho_contagil(arg):
            arg.mkdir(parents=True, exist_ok=True)
            return arg
        local = DATA_DIR / "contagil_winpython" / "saida"
        local.mkdir(parents=True, exist_ok=True)
        print(f"[AVISO] Saída ContAgil ausente: {arg}")
        print(f"   Usando espelho local: {local}")
        return local
    if CONTAGIL_PASTA_SAIDA.exists():
        return CONTAGIL_PASTA_SAIDA
    cwd_saida = Path.cwd() / "saida"
    if cwd_saida.exists() or (Path.cwd() / "dados").exists():
        cwd_saida.mkdir(parents=True, exist_ok=True)
        return cwd_saida
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUT_DIR


def resolver_fatores(arg: Path | None) -> Path:
    candidatos: list[Path] = []
    if arg is not None:
        candidatos.append(arg)
    candidatos.extend(
        [
            Path.cwd() / FATORES_DEFAULT_NOME,
            CONTAGIL_WINPYTHON / FATORES_DEFAULT_NOME,
            DATA_DIR / FATORES_DEFAULT_NOME,
            DATA_DIR / "selic_mensal.xlsx",
            DATA_DIR / "selic_taxas_contagil.xlsx",
        ]
    )
    for cand in candidatos:
        if cand is not None and cand.exists():
            return cand
    raise FileNotFoundError(
        f"Arquivo de fatores não encontrado. Procurado: {FATORES_DEFAULT_NOME}\n"
        "Coloque fator_acumulado_SELIC_TJLP_TLP.xlsx na pasta winpython "
        "ou informe --fatores."
    )


def carregar_fatores_mensais(path: Path) -> SelicSerie:
    """Carrega fator_acumulado_SELIC_TJLP_TLP.xlsx (Data + Fator_Acumulado)."""
    raw = pd.read_excel(path)
    print(f"Carregando fatores combinados: {path}")
    print(f"Colunas encontradas: {list(raw.columns)}")

    cols_norm = {str(c).strip().lower(): c for c in raw.columns}
    data_col = None
    for cand in ("data", "date", "mes", "mês"):
        if cand in cols_norm:
            data_col = cols_norm[cand]
            break
    if data_col is None:
        data_col = raw.columns[0]

    fator_col = None
    for cand in ("fator_acumulado", "fator", "fator acumulado"):
        if cand in cols_norm:
            fator_col = cols_norm[cand]
            break
    if fator_col is None:
        # ContAgil: Taxa_Mensal_% + Fator_Acumulado - pega última numérica
        for c in raw.columns:
            if "fator" in str(c).lower():
                fator_col = c
                break
    if fator_col is None:
        fator_col = raw.columns[-1]

    datas = pd.to_datetime(raw[data_col], dayfirst=True, errors="coerce")
    fatores = pd.to_numeric(raw[fator_col], errors="coerce")
    mask = datas.notna() & fatores.notna() & (fatores > 0)
    datas_arr = datas[mask].to_numpy(dtype="datetime64[ns]")
    fatores_arr = fatores[mask].to_numpy(dtype=float)

    # Referência: 01/06/2026 (ou último fator ≤ data de impacto)
    data_ref = np.datetime64(DATA_REF_DEFAULT)
    ate_ref = fatores_arr[datas_arr <= data_ref]
    if len(ate_ref):
        fator_ref = float(ate_ref[-1])
        data_ref_usada = pd.Timestamp(
            datas_arr[datas_arr <= data_ref][-1]
        ).strftime("%Y-%m-%d")
    else:
        fator_ref = float(fatores_arr[-1])
        data_ref_usada = pd.Timestamp(datas_arr[-1]).strftime("%Y-%m-%d")

    print(f"Referência de atualização: {data_ref_usada} | Fator SELIC = {fator_ref:.8f}")
    print(f"SELIC: {len(fatores_arr)} meses (arquivo de fatores)")

    return SelicSerie(
        datas_arr,
        fatores_arr,
        origem=f"mensal:{path.name}",
        fator_referencia=fator_ref,
    )


def listar_contratos(pasta: Path) -> list[Path]:
    arquivos = sorted(Path(p) for p in glob.glob(os.path.join(str(pasta), "*.xlsx")))
    saida: list[Path] = []
    for arq in arquivos:
        nome = arq.name.upper()
        if nome.startswith("STP") or "SELIC" in nome or "FATOR" in nome or "TJLP" in nome:
            continue
        if nome.startswith("~$"):
            continue
        if "NUMERADOS" in nome or "DISCRIMINATIV" in nome:
            continue
        if "DIRETA" in nome and "INDIRET" not in nome:
            continue
        saida.append(arq)
    return saida


def diagnosticar_colunas(path: Path) -> None:
    """Imprime colunas brutas e mapeamento - ajuda quando o arquivo é pulado."""
    for h in (0, 5, 1, 2, 3, 4):
        try:
            df = pd.read_excel(path, sheet_name=0, header=h, nrows=5)
        except Exception as exc:  # noqa: BLE001
            print(f"    header={h}: erro ao ler ({exc})")
            continue
        mapped, rename = _mapear_colunas_contratos(df)
        ok = _excel_tem_colunas_contratos(df)
        print(f"    header={h}: colunas={list(df.columns)[:12]}")
        print(f"             mapeadas={rename or '{}'} | ok={ok}")
        if ok:
            missing = [
                c
                for c in (
                    "data_contratacao",
                    "valor_desembolsado",
                    "juros",
                    "prazo_carencia",
                    "prazo_amortizacao",
                )
                if c not in mapped.columns
            ]
            if missing:
                print(f"             ainda faltam: {missing}")
            break


def _contar_linhas_excel(path: Path) -> int | None:
    """Conta linhas de dados sem carregar a planilha inteira em memória."""
    try:
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        # max_row inclui header; reporta linhas brutas como o log ContAgil
        n = ws.max_row or 0
        wb.close()
        return int(n)
    except Exception:  # noqa: BLE001
        return None


def processar_arquivo(
    arquivo: Path,
    pasta_saida: Path,
    serie: SelicSerie,
    max_contratos: int | None = None,
) -> Path | None:
    print(f"\n>>> {arquivo.name} ...")
    n_linhas = _contar_linhas_excel(arquivo)
    if n_linhas is not None:
        print(f"    Linhas: {n_linhas:,}")

    try:
        df = load_from_excel(arquivo)
    except ValueError as exc:
        print("    [AVISO] Colunas obrigatórias não encontradas. Pulando.")
        print(f"    Detalhe: {exc}")
        diagnosticar_colunas(arquivo)
        return None
    except Exception as exc:  # noqa: BLE001
        print(f"    [AVISO] Falha ao carregar contratos ({exc}). Pulando.")
        diagnosticar_colunas(arquivo)
        return None

    if n_linhas is None:
        print(f"    Linhas: {len(df):,} (após limpeza)")

    if max_contratos is not None:
        df = df.head(int(max_contratos)).copy()

    if df.empty:
        print("    [AVISO] Nenhum contrato válido após limpeza. Pulando.")
        return None

    # Garante número único N-AAAA (já aplicado em load_from_excel / preparar)
    if "numero_contrato" in df.columns and len(df):
        print(
            f"    Contratos: {df['numero_contrato'].iloc[0]} ... {df['numero_contrato'].iloc[-1]}"
        )

    pasta_saida.mkdir(parents=True, exist_ok=True)
    saida = pasta_saida / f"fluxos_{arquivo.stem}.xlsx"
    fluxos = gerar_fluxos(df, serie)
    # Excel ~1M linhas: se passar, grava CSV completo + amostra xlsx
    if len(fluxos) > 1_000_000:
        csv_path = saida.with_suffix(".csv")
        fluxos.to_csv(csv_path, index=False)
        fluxos.head(1_000_000).to_excel(saida, index=False)
        print(f"    -> CSV completo: {csv_path} ({len(fluxos):,} parcelas)")
        print(f"    -> Excel (amostra 1M): {saida}")
    else:
        fluxos.to_excel(saida, index=False)
        print(f"    -> Salvo: {saida} ({len(fluxos):,} parcelas)")
    return saida


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--massa-dados",
        "--pasta-dados",
        dest="massa_dados",
        type=Path,
        default=None,
        help="Pasta com BNDES INDIRETAS *.xlsx (default: winpython/dados).",
    )
    p.add_argument(
        "--pasta-saida",
        type=Path,
        default=None,
        help="Pasta de saída (default: winpython/saida).",
    )
    p.add_argument(
        "--fatores",
        type=Path,
        default=None,
        help=f"Excel de fatores mensais (default: {FATORES_DEFAULT_NOME}).",
    )
    p.add_argument(
        "--max-contratos",
        type=int,
        default=None,
        help="Limita contratos por arquivo (teste rápido).",
    )
    p.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Processa um único Excel/CSV em vez da massa.",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    _banner()

    pasta_dados = resolver_pasta_dados(args.massa_dados)
    pasta_saida = resolver_pasta_saida(args.pasta_saida)
    print(f"Massa de dados : {pasta_dados}")
    print(f"Pasta de saída : {pasta_saida}")
    print()

    fatores_path = resolver_fatores(args.fatores)
    serie = carregar_fatores_mensais(fatores_path)
    print()

    if args.input is not None:
        arquivos = [Path(args.input)]
    else:
        if not pasta_dados.exists():
            print(f"[ERRO] Massa de dados não encontrada: {pasta_dados}")
            return 1
        arquivos = listar_contratos(pasta_dados)

    print(f"Arquivos de contratos ({len(arquivos)}):")
    for a in arquivos:
        print(f"  - {a.name}")
    print()

    if not arquivos:
        print("Nenhum arquivo de contrato encontrado.")
        return 1

    ok: list[Path] = []
    for arq in arquivos:
        out = processar_arquivo(
            arq, pasta_saida, serie, max_contratos=args.max_contratos
        )
        if out is not None:
            ok.append(out)

    print()
    if not ok:
        print("Nenhum contrato válido processado.")
        return 2

    print(f"Concluído: {len(ok)} arquivo(s) processado(s).")
    print(f"Referência de impacto: {DATA_IMPACTO.date()} (capitalização mensal).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
