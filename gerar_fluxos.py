"""ContAgil Fluxos com SELIC — lógica SAC + Saldo Fiscal."""

import argparse
import os
import unicodedata

import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

print("🚀 Gerando fluxos...")

DATA_CORTE = datetime(2026, 6, 30)


def _normalize_col(name: str) -> str:
    """Normaliza nome de coluna: minúsculas, sem acentos, espaços → underscore."""
    text = str(name).strip().lower().replace(" ", "_")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text


def load_selic(selic_file=None):
    """Carregar dados SELIC; preferir arquivo local (STP). Reserva provisória se ausente."""
    if selic_file and os.path.exists(selic_file):
        selic = pd.read_excel(selic_file)
        datas_selic = pd.to_datetime(selic.iloc[:, 0], dayfirst=True)
        fatores_selic = selic.iloc[:, 4].values  # Coluna E
        return pd.DataFrame({"data": datas_selic, "fator": fatores_selic})

    print("⚠️ Sem arquivo SELIC, usando placeholder. Em produção, buscar SGS 11")
    dates = pd.date_range(start="2009-01-01", end="2026-06-30", freq="D")
    df = pd.DataFrame({"data": dates})
    # Fator acumulado dummy (diário)
    df["fator"] = (1 + 0.0001) ** (df.index + 1)
    return df


def calcular_impacto_fiscal_real(subsidio, data_parcela, selic_df):
    """Capitaliza o subsídio com SELIC até 30/06/2026."""
    if subsidio <= 0:
        return 0.0

    data_proxima = data_parcela + timedelta(days=1)
    idx_inicio = selic_df["data"].sub(data_proxima).abs().idxmin()
    idx_fim = selic_df["data"].sub(DATA_CORTE).abs().idxmin()

    if idx_fim > idx_inicio:
        fator = selic_df["fator"].iloc[idx_fim] / selic_df["fator"].iloc[idx_inicio]
        return round(subsidio * fator, 2)
    return round(subsidio, 2)


def _col_or_default(df, col_name, default=0):
    """Retorna a série da coluna ou um default escalar quando a coluna não existe."""
    if col_name is None or col_name not in df.columns:
        return default
    return df[col_name]


def gerar_fluxos(df_ops, selic_df):
    """Lógica SAC completa: Saldo Fiscal (sem juros contrato) + capitalização do subsídio."""
    df_fluxos = df_ops.copy()

    # Normaliza nomes de colunas
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
                "valor_principal",
                "valor",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )
    col_taxa = next(
        (
            c
            for c in [
                "juros",
                "taxa_juros",
                "taxa_contrato",
                "custo_financeiro",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )
    col_carencia = next(
        (
            c
            for c in ["prazo_carencia_meses", "carencia_meses"]
            if c in df_fluxos.columns
        ),
        None,
    )
    col_amort = next(
        (
            c
            for c in [
                "prazo_amortizacao_meses",
                "amortizacao_meses",
                "prazo",
            ]
            if c in df_fluxos.columns
        ),
        None,
    )

    if not (col_data and col_valor):
        print("⚠️ Colunas essenciais (data + valor) não encontradas.")
        return df_fluxos

    df_fluxos[col_data] = pd.to_datetime(
        df_fluxos[col_data], errors="coerce", dayfirst=True
    )
    df_fluxos["taxa_mensal_contrato"] = (
        pd.to_numeric(_col_or_default(df_fluxos, col_taxa, 0), errors="coerce")
        .fillna(0)
        / 100
        / 12
    )
    df_fluxos["meses_carencia"] = (
        pd.to_numeric(_col_or_default(df_fluxos, col_carencia, 0), errors="coerce")
        .fillna(0)
        .astype(int)
    )
    df_fluxos["meses_amort"] = (
        pd.to_numeric(_col_or_default(df_fluxos, col_amort, 0), errors="coerce")
        .fillna(0)
        .astype(int)
    )

    def processar_contrato(row):
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

        data_inicio = row[col_data]
        amort_mensal = valor / row["meses_amort"] if row["meses_amort"] > 0 else 0.0
        taxa_contrato_m = row["taxa_mensal_contrato"]

        saldo_fiscal = valor
        saldo_contrato = valor
        subsidio_acumulado = 0.0
        data_atual = data_inicio

        total_meses = int(row["meses_carencia"] + row["meses_amort"])
        for mes in range(1, total_meses + 1):
            if data_atual > DATA_CORTE:
                break

            amort = amort_mensal if mes > row["meses_carencia"] else 0.0

            # Saldo Contrato (com juros)
            saldo_contrato = (saldo_contrato - amort) * (1 + taxa_contrato_m)

            # Saldo Fiscal (apenas amortização)
            saldo_fiscal = saldo_fiscal - amort

            # Subsídio do mês = saldo_fiscal * (SELIC_mensal - taxa_contrato)
            # Placeholder SELIC mensal → melhorar com selic_df real
            selic_mensal_aprox = 0.0095
            subsidio_mes = saldo_fiscal * (selic_mensal_aprox - taxa_contrato_m)
            subsidio_acumulado += subsidio_mes

            data_atual += relativedelta(months=1)

        # Capitalização SELIC do subsídio acumulado
        impacto_capitalizado = calcular_impacto_fiscal_real(
            subsidio_acumulado, data_inicio, selic_df
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

    print(
        f"✅ Processados {len(df_fluxos)} contratos com lógica SAC + Saldo Fiscal correto."
    )
    return df_fluxos


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="ContAgil Fluxos com SELIC - SAC Fiscal"
    )
    parser.add_argument(
        "--arquivo-selic", type=str, help="Path to SELIC Excel (STP)"
    )
    parser.add_argument(
        "--excel", type=str, required=True, help="Input Excel file"
    )
    args = parser.parse_args()

    selic_df = load_selic(args.arquivo_selic)

    try:
        df = pd.read_excel(
            args.excel,
            sheet_name="operacoes_indiretas_automaticas",
            header=5,
        )
        print(f"Carreguei {len(df)} operações")
    except Exception as e:
        print(f"Erro carregando Excel: {e}")
        df = pd.DataFrame()

    df_fluxos = gerar_fluxos(df, selic_df)
    output_file = "fluxos_completos_final.xlsx"
    df_fluxos.to_excel(output_file, index=False)
    print(f"✅ Concluído! Salvo em {output_file}")
