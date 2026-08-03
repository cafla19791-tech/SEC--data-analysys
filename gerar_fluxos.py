#!/usr/bin/env python3
"""ContAgil Fluxos com SELIC — lógica SAC + Saldo Fiscal + fluxo diário."""

from __future__ import annotations

import argparse
import glob
import os
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

print("🚀 Gerando fluxos com fator Selic correto (Coluna D)...")

DATA_CORTE = datetime(2026, 6, 30)
FLUXOS_DIARIOS_NOME = "fluxos_diarios_detalhados.xlsx"
TJLP_TLP_BASE = 0.06
TAXA_SELIC_ANUAL = 0.145
# Fator ContAgil de referência em 30/06/2026 (coluna D do STP)
FATOR_30_06_2026 = 82.84819


def taxa_contrato_efetiva(custo_financeiro: str | None, juros_pct: float) -> float:
    """Taxa mensal efetiva do contrato (lógica corrigida ContAgil).

    - TAXA FIXA / demais: ``(1 + juros)^(1/12) − 1``
    - TJLP / TLP: ``(1 + 0,06)^(1/12) × (1 + juros)^(1/12) − 1``
    """
    try:
        juros = float(juros_pct) / 100.0
    except (TypeError, ValueError):
        juros = 0.0

    custo = str(custo_financeiro or "").upper()
    if "TJLP" in custo:
        return (1.0 + TJLP_TLP_BASE) ** (1.0 / 12.0) * (1.0 + juros) ** (1.0 / 12.0) - 1.0
    if "TLP" in custo:
        return (1.0 + TJLP_TLP_BASE) ** (1.0 / 12.0) * (1.0 + juros) ** (1.0 / 12.0) - 1.0
    return (1.0 + juros) ** (1.0 / 12.0) - 1.0


def _normalize_col(name: str) -> str:
    """Normaliza nome de coluna ContAgil: minúsculas, sem acentos/símbolos."""
    text = str(name).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    # ContAgil: "Valor Desembolsado R$ (*)", "Prazo - Carência (meses)"
    cleaned: list[str] = []
    for ch in text:
        if ch.isalnum():
            cleaned.append(ch)
        else:
            cleaned.append("_")
    text = "".join(cleaned)
    while "__" in text:
        text = text.replace("__", "_")
    return text.strip("_")


def load_selic(selic_file: str | None = None) -> pd.DataFrame:
    """Load SELIC data, prefer local file, fallback placeholder.

    ContAgil corrigido: fatores na coluna D (índice 3).
    Preferência: fator_acumulado / fator nomeado → col D → última coluna.
    """
    if selic_file and os.path.exists(selic_file):
        selic = pd.read_excel(selic_file)
        datas_selic = pd.to_datetime(selic.iloc[:, 0], dayfirst=True)
        cols_lower = {str(c).strip().lower(): c for c in selic.columns}
        fator_ref = None
        if "fator_acumulado" in cols_lower:
            fatores_selic = selic[cols_lower["fator_acumulado"]].values
        elif "fator" in cols_lower:
            fatores_selic = selic[cols_lower["fator"]].values
        elif selic.shape[1] >= 4:
            fatores_selic = selic.iloc[:, 3].values  # Coluna D
            fator_ref = FATOR_30_06_2026
        else:
            fatores_selic = selic.iloc[:, -1].values
        out = pd.DataFrame({"data": datas_selic, "fator": fatores_selic})
        if fator_ref is not None:
            out.attrs["fator_referencia"] = fator_ref
        return out

    print("⚠️ No SELIC file, using placeholder. In production fetch SGS 11")
    dates = pd.date_range(start="2009-01-01", end="2026-06-30", freq="D")
    df = pd.DataFrame({"data": dates})
    # Fator acumulado dummy (diário) — sem fator de referência ContAgil
    df["fator"] = (1 + 0.0001) ** (df.index + 1)
    return df


def get_selic_daily_factor(date, selic_df: pd.DataFrame) -> float:
    """Fator acumulado na data mais próxima (ContAgil nearest)."""
    if selic_df is None or len(selic_df) == 0:
        return 1.0
    idx = selic_df["data"].sub(pd.to_datetime(date)).abs().idxmin()
    return float(selic_df["fator"].iloc[idx])


def calcular_impacto_fiscal_real(subsidio, data_parcela, selic_df: pd.DataFrame) -> float:
    """Capitaliza o subsídio com SELIC até 30/06/2026 (fator coluna D).

    ContAgil:
      idx = nearest(data_parcela)
      impacto = subsidio * FATOR_30_06_2026 / fator_parcela
    """
    if subsidio <= 0:
        return 0.0
    idx = selic_df["data"].sub(pd.to_datetime(data_parcela)).abs().idxmin()
    fator_parcela = float(selic_df["fator"].iloc[idx])
    if fator_parcela <= 0:
        return 0.0
    fator_ref = selic_df.attrs.get("fator_referencia")
    if fator_ref is not None:
        fator_fim = float(fator_ref)
    else:
        idx_fim = selic_df["data"].sub(pd.to_datetime(DATA_CORTE)).abs().idxmin()
        fator_fim = float(selic_df["fator"].iloc[idx_fim])
    if fator_fim <= 0:
        return round(float(subsidio), 2)
    return round(float(subsidio) * (fator_fim / fator_parcela), 2)


def _col_or_default(df: pd.DataFrame, col_name: str | None, default=0):
    """Retorna a série da coluna ou um default escalar quando a coluna não existe."""
    if col_name is None or col_name not in df.columns:
        return default
    return df[col_name]


def _to_float(value) -> float:
    """Converte BR (1.234,56) / US (1234.56) / numérico."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return float("nan")
    if isinstance(value, (int, float, np.floating, np.integer)):
        return float(value)
    text = str(value).replace("R$", "").strip()
    if not text or text.lower() in {"nan", "none"}:
        return float("nan")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return float("nan")


def _taxa_diaria_contrato(taxa_mensal: float) -> float:
    """Equivalente diário composto da taxa mensal do contrato."""
    return (1.0 + float(taxa_mensal)) ** (1.0 / 30.0) - 1.0


def _selic_diaria_exata(dia, selic_df: pd.DataFrame) -> float:
    """Taxa SELIC do dia via fatores ContAgil: fator(d+1)/fator(d) − 1."""
    f0 = get_selic_daily_factor(dia, selic_df)
    f1 = get_selic_daily_factor(pd.Timestamp(dia) + timedelta(days=1), selic_df)
    if f0 > 0:
        return (f1 / f0) - 1.0
    return 0.0


def gerar_fluxos(
    df_ops: pd.DataFrame,
    selic_df: pd.DataFrame,
    fluxo_diario: bool = False,
    max_contratos: int | None = None,
    saida_diario: str | Path | None = None,
) -> pd.DataFrame:
    """Lógica SAC completa com SELIC diário exato.

    Quando fluxo_diario=True, anexa linhas dia a dia e grava
    fluxos_diarios_detalhados.xlsx (ou saida_diario).
    """
    df_fluxos = df_ops.copy()

    # Normaliza colunas
    df_fluxos.columns = [_normalize_col(col) for col in df_fluxos.columns]

    # Mapeamento de colunas
    col_data = next(
        (
            c
            for c in [
                "data_da_contratacao",
                "data_contratacao",
                "data_parcela",
                "data",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )
    col_valor = next(
        (
            c
            for c in [
                "valor_da_operacao_em_reais",
                "valor_desembolsado_reais",
                "valor_desembolsado_r",
                "valor_principal",
                "valor",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )
    if col_valor is None:
        # ContAgil: valor_desembolsado_r_* após normalizar "R$ (*)"
        for c in df_fluxos.columns:
            if "valor_desembolsado" in c or c.startswith("valor_da_operacao"):
                col_valor = c
                break
    col_taxa = next(
        (
            c
            for c in [
                "juros",
                "taxa_juros",
                "taxa_contrato",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )
    col_carencia = next(
        (
            c
            for c in [
                "prazo_carencia_meses",
                "prazo_carencia",
                "carencia_meses",
                "carencia",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )
    col_amort = next(
        (
            c
            for c in [
                "prazo_amortizacao_meses",
                "prazo_amortizacao",
                "amortizacao_meses",
                "prazo",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )
    col_agente = next(
        (
            c
            for c in [
                "instituicao_financeira_credenciada",
                "agente",
                "instituicao_financeira",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )
    col_custo = next(
        (
            c
            for c in [
                "custo_financeiro",
                "custo_financeiro_da_operacao",
                "custo",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )

    if not (col_data and col_valor):
        print("⚠️ Colunas essenciais não encontradas.")
        return df_fluxos

    if max_contratos is not None:
        df_fluxos = df_fluxos.head(int(max_contratos)).copy()

    datas = df_fluxos[col_data]
    if pd.api.types.is_datetime64_any_dtype(datas):
        df_fluxos[col_data] = pd.to_datetime(datas, errors="coerce")
    else:
        s = datas.astype(str).str.strip()
        iso = s.str.match(r"^\d{4}-\d{2}-\d{2}", na=False).fillna(False)
        out = pd.Series(pd.NaT, index=df_fluxos.index, dtype="datetime64[ns]")
        if iso.any():
            out.loc[iso] = pd.to_datetime(s[iso], errors="coerce")
        if (~iso).any():
            out.loc[~iso] = pd.to_datetime(s[~iso], errors="coerce", dayfirst=True)
        df_fluxos[col_data] = out
    df_fluxos[col_valor] = df_fluxos[col_valor].map(_to_float)
    taxa_src = _col_or_default(df_fluxos, col_taxa, 0)
    if not isinstance(taxa_src, (int, float)):
        taxa_src = taxa_src.map(_to_float)
    juros_pct = pd.to_numeric(taxa_src, errors="coerce").fillna(0.0)
    if col_custo is not None:
        custos = df_fluxos[col_custo].astype(str)
        df_fluxos["taxa_mensal_contrato"] = [
            taxa_contrato_efetiva(c, j) for c, j in zip(custos, juros_pct)
        ]
    else:
        df_fluxos["taxa_mensal_contrato"] = [
            taxa_contrato_efetiva(None, j) for j in juros_pct
        ]
    carencia_src = _col_or_default(df_fluxos, col_carencia, 0)
    if not isinstance(carencia_src, (int, float)):
        carencia_src = carencia_src.map(_to_float)
    df_fluxos["meses_carencia"] = (
        pd.to_numeric(carencia_src, errors="coerce").fillna(0).astype(int)
    )
    amort_src = _col_or_default(df_fluxos, col_amort, 0)
    if not isinstance(amort_src, (int, float)):
        amort_src = amort_src.map(_to_float)
    df_fluxos["meses_amort"] = (
        pd.to_numeric(amort_src, errors="coerce").fillna(0).astype(int)
    )

    fluxos_diarios: list[dict] = []

    def processar_contrato(row: pd.Series) -> pd.Series:
        vazio = {
            "amortizacao_mensal": 0.0,
            "saldo_contrato_final": 0.0,
            "saldo_fiscal_final": 0.0,
            "subsidio_acumulado": 0.0,
            "impacto_fiscal_real": 0.0,
        }

        valor = float(row[col_valor]) if pd.notna(row[col_valor]) else 0.0
        if valor <= 0 or pd.isna(row[col_data]):
            return pd.Series(vazio)

        data_inicio = pd.Timestamp(row[col_data])
        amort_mensal = valor / row["meses_amort"] if row["meses_amort"] > 0 else 0.0
        taxa_contrato_m = float(row["taxa_mensal_contrato"])
        taxa_contrato_d = _taxa_diaria_contrato(taxa_contrato_m)

        saldo_fiscal = valor
        saldo_contrato = valor
        subsidio_acumulado = 0.0
        data_atual = data_inicio

        contrato_id = int(row.name) if row.name is not None else 0
        agente = "Não informado"
        if col_agente is not None and pd.notna(row.get(col_agente)):
            agente = str(row[col_agente])

        total_meses = int(row["meses_carencia"] + row["meses_amort"])
        for mes in range(1, total_meses + 1):
            if data_atual > DATA_CORTE:
                break

            em_carencia = mes <= int(row["meses_carencia"])
            amort = 0.0 if em_carencia else amort_mensal

            # SELIC mensal via fatores ContAgil (col D); fallback 14,5% a.a. composta
            fator_inicio = get_selic_daily_factor(data_atual, selic_df)
            data_fim_mes = data_atual + relativedelta(months=1)
            fator_fim = get_selic_daily_factor(data_fim_mes, selic_df)
            if fator_inicio > 0:
                selic_mensal_aprox = (fator_fim / fator_inicio) - 1.0
            else:
                selic_mensal_aprox = (1.0 + TAXA_SELIC_ANUAL) ** (1.0 / 12.0) - 1.0

            # Subsídio sobre saldo fiscal ANTES da amortização (lógica corrigida)
            subsidio_mes = saldo_fiscal * (selic_mensal_aprox - taxa_contrato_m)
            subsidio_acumulado += subsidio_mes

            if fluxo_diario:
                dia = pd.Timestamp(data_atual)
                fim_periodo = min(
                    pd.Timestamp(data_fim_mes) - timedelta(days=1),
                    pd.Timestamp(DATA_CORTE),
                )
                while dia <= fim_periodo:
                    selic_d = _selic_diaria_exata(dia, selic_df)
                    subsidio_d = saldo_fiscal * (selic_d - taxa_contrato_d)
                    impacto_d = calcular_impacto_fiscal_real(
                        subsidio_d, dia.to_pydatetime(), selic_df
                    )
                    fluxos_diarios.append(
                        {
                            "contrato": contrato_id,
                            "Instituição Financeira": agente,
                            "mes": mes,
                            "data_fluxo": dia.date(),
                            "saldo_fiscal": round(saldo_fiscal, 2),
                            "saldo_contrato": round(saldo_contrato, 2),
                            "saldo": round(saldo_fiscal, 2),
                            "amortizacao": round(
                                amort if dia == pd.Timestamp(data_atual) else 0.0, 2
                            ),
                            "taxa_selic_diaria": round(selic_d, 10),
                            "taxa_contrato_diaria": round(taxa_contrato_d, 10),
                            "selic_mensal_periodo": round(selic_mensal_aprox, 8),
                            "taxa_contrato_mensal": (
                                round(taxa_contrato_m, 8) if mes == 1 else None
                            ),
                            "subsidio": round(subsidio_d, 4),
                            "impacto_fiscal": impacto_d,
                            "em_carencia": em_carencia,
                            "dia_parcela": dia == pd.Timestamp(data_atual),
                        }
                    )
                    dia += timedelta(days=1)

            # Atualização dos saldos (dual balance)
            if not em_carencia:
                saldo_fiscal -= amort
                saldo_contrato = (saldo_contrato - amort) * (1.0 + taxa_contrato_m)
            else:
                saldo_contrato = saldo_contrato * (1.0 + taxa_contrato_m)

            if saldo_fiscal <= 1e-9:
                break

            data_atual += relativedelta(months=1)

        impacto_capitalizado = calcular_impacto_fiscal_real(
            subsidio_acumulado, data_inicio.to_pydatetime(), selic_df
        )

        return pd.Series(
            {
                "amortizacao_mensal": round(amort_mensal, 2),
                "saldo_contrato_final": round(saldo_contrato, 2),
                "saldo_fiscal_final": round(max(saldo_fiscal, 0), 2),
                "subsidio_acumulado": round(subsidio_acumulado, 2),
                "impacto_fiscal_real": round(impacto_capitalizado, 2),
            }
        )

    extra = df_fluxos.apply(processar_contrato, axis=1)
    df_fluxos = pd.concat([df_fluxos, extra], axis=1)

    if fluxo_diario:
        out = Path(saida_diario) if saida_diario is not None else Path(FLUXOS_DIARIOS_NOME)
        out.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(fluxos_diarios).to_excel(out, index=False)
        print(f"✅ Fluxos diários: {out} ({len(fluxos_diarios):,} linhas)")

    print(f"✅ Processados {len(df_fluxos)} contratos com lógica SAC + SELIC diário.")
    return df_fluxos


def _listar_entradas(caminho: str | Path) -> list[Path]:
    """Aceita arquivo .xlsx/.csv, pasta com vários .xlsx, ou glob."""
    p = Path(caminho)
    if p.is_file():
        return [p]
    if p.is_dir():
        xlsx = sorted(Path(x) for x in glob.glob(os.path.join(str(p), "*.xlsx")))
        csvs = sorted(Path(x) for x in glob.glob(os.path.join(str(p), "*.csv")))
        return xlsx or csvs
    return sorted(Path(x) for x in glob.glob(str(caminho)))


def _ler_operacoes(path: Path) -> pd.DataFrame:
    """Lê CSV (sep=;) ou Excel ContAgil (sheet/header=5) / padrão."""
    if path.suffix.lower() == ".csv":
        try:
            return pd.read_csv(path, sep=";", dtype=str)
        except Exception:
            return pd.read_csv(path, dtype=str)

    try:
        return pd.read_excel(
            path,
            sheet_name="operacoes_indiretas_automaticas",
            header=5,
        )
    except Exception:
        try:
            return pd.read_excel(path, header=5)
        except Exception:
            return pd.read_excel(path)


CONTAGIL_WINPYTHON = Path(
    r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
)
CONTAGIL_PASTA_DADOS = CONTAGIL_WINPYTHON / "dados"
CONTAGIL_PASTA_SAIDA = CONTAGIL_WINPYTHON / "saida"
CONTAGIL_SELIC_DEFAULT = CONTAGIL_WINPYTHON / "STP-20260716182715078 (1).xlsx"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="ContAgil Fluxos with SELIC")
    parser.add_argument(
        "--excel",
        type=str,
        default=None,
        help="Arquivo .xlsx, pasta com .xlsx, ou glob",
    )
    parser.add_argument(
        "--massa-dados",
        "--pasta-dados",
        dest="massa_dados",
        type=str,
        default=None,
        help=(
            "Massa ContAgil (pasta com vários .xlsx). "
            "Equivalente a --excel apontando para a pasta."
        ),
    )
    parser.add_argument(
        "--arquivo-selic",
        type=str,
        default=None,
        help="SELIC STP (col A=data, col D=fator)",
    )
    parser.add_argument(
        "--fluxo-diario",
        action="store_true",
        help="Gera tabela detalhada dia a dia (fluxos_diarios_detalhados.xlsx)",
    )
    parser.add_argument(
        "--pasta-saida",
        "--output-dir",
        dest="output_dir",
        type=str,
        default=None,
        help="Pasta de saída ContAgil (alias: --output-dir).",
    )
    parser.add_argument("--max-contratos", type=int, default=None)
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="Pula arquivos cujo Excel de saída já existe",
    )
    parser.add_argument("--prefix", type=str, default="")
    return parser.parse_args(argv)


def _resolver_entrada_e_saida(args: argparse.Namespace) -> tuple[str, str]:
    """Resolve massa ContAgil / excel e pasta de saída com defaults WinPython."""
    entrada = args.excel or args.massa_dados
    if entrada is None and CONTAGIL_PASTA_DADOS.exists():
        entrada = str(CONTAGIL_PASTA_DADOS)
    if entrada is None:
        raise SystemExit(
            "Informe --massa-dados/--excel (pasta ou arquivo ContAgil)."
        )

    saida = args.output_dir
    if saida is None:
        if CONTAGIL_PASTA_SAIDA.exists():
            saida = str(CONTAGIL_PASTA_SAIDA)
        else:
            saida = "saida"
    return entrada, saida


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    entrada, pasta_saida = _resolver_entrada_e_saida(args)

    arquivo_selic = args.arquivo_selic
    if arquivo_selic is None and CONTAGIL_SELIC_DEFAULT.exists():
        arquivo_selic = str(CONTAGIL_SELIC_DEFAULT)

    selic_df = load_selic(arquivo_selic)
    os.makedirs(pasta_saida, exist_ok=True)
    print(f"Massa/entrada: {entrada}")
    print(f"Pasta de saída: {pasta_saida}")

    arquivos = _listar_entradas(entrada)
    if not arquivos:
        print(f"⚠️ Nenhum .xlsx/.csv em: {entrada}")
        return 1

    saidas: list[Path] = []
    for entrada in arquivos:
        nome = entrada.name.upper()
        if nome.startswith("STP") or "SELIC" in nome:
            print(f"Ignorando série SELIC: {entrada.name}")
            continue

        out_name = f"{args.prefix}fluxos_{entrada.stem}.xlsx"
        out_path = Path(pasta_saida) / out_name
        if args.skip_existing and out_path.exists():
            print(f"⏭️  Já existe, pulando: {out_path}")
            continue

        print(f"Processando: {entrada}")
        try:
            df = _ler_operacoes(entrada)
            print(f"  Carreguei {len(df):,} operações")
        except Exception as exc:  # noqa: BLE001
            print(f"  Erro carregando arquivo: {exc}")
            continue

        diario_path = Path(pasta_saida) / (
            f"{args.prefix}{FLUXOS_DIARIOS_NOME}"
            if len(arquivos) == 1
            else f"{args.prefix}fluxos_diarios_{entrada.stem}.xlsx"
        )

        df_fluxos = gerar_fluxos(
            df,
            selic_df,
            fluxo_diario=args.fluxo_diario,
            max_contratos=args.max_contratos,
            saida_diario=diario_path if args.fluxo_diario else None,
        )
        df_fluxos.to_excel(out_path, index=False)
        print(f"  → Salvo: {out_path}")
        saidas.append(out_path)

    if not saidas:
        print("Nenhum arquivo processado.")
        return 1

    print(f"✅ Concluído! {len(saidas)} arquivo(s) em {pasta_saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
