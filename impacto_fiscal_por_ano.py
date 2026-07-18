import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Configurações
TAXA_SELIC_ANUAL = 0.145
DATA_REFERENCIA = datetime(2026, 6, 30)

print("Carregando arquivo de fluxos...")

# Carregar o arquivo CSV gerado
df = pd.read_csv("output/fluxos_completos_corrigido.csv")

print(f"Total de parcelas carregadas: {len(df):,}")

# Converter data_fluxo para datetime
df["data_fluxo"] = pd.to_datetime(df["data_fluxo"])

# Extrair o ano do pagamento
df["ano_pagamento"] = df["data_fluxo"].dt.year


# Calcular meses até 30/06/2026 para cada parcela
def calcular_meses_ate_2026(data):
    delta = relativedelta(DATA_REFERENCIA, data)
    return delta.years * 12 + delta.months


df["meses_ate_2026"] = df["data_fluxo"].apply(calcular_meses_ate_2026)

# Calcular impacto individual de cada parcela
df["impacto_individual"] = df["subsidio"] * (1 + TAXA_SELIC_ANUAL / 12) ** df["meses_ate_2026"]

# Agrupar por ano de pagamento
resumo = (
    df.groupby("ano_pagamento")
    .agg({"subsidio": "sum", "impacto_individual": "sum", "mes": "count"})
    .reset_index()
)

resumo.columns = [
    "Ano",
    "Soma Subsídio Nominal (R$)",
    "Impacto Fiscal 2026 (R$)",
    "Quantidade de Parcelas",
]

# Formatar valores
resumo["Soma Subsídio Nominal (R$)"] = resumo["Soma Subsídio Nominal (R$)"].round(2)
resumo["Impacto Fiscal 2026 (R$)"] = resumo["Impacto Fiscal 2026 (R$)"].round(2)

print("\n" + "=" * 80)
print("IMPACTO FISCAL POR ANO DE PAGAMENTO")
print("=" * 80)
print(resumo.to_string(index=False))

# Salvar resultado
resumo.to_excel("impacto_fiscal_por_ano.xlsx", index=False)
print("\nArquivo salvo: impacto_fiscal_por_ano.xlsx")

# Totais
print("\n" + "=" * 80)
print("TOTAIS GERAIS")
print("=" * 80)
print(f"Total Subsídio Nominal: R$ {resumo['Soma Subsídio Nominal (R$)'].sum():,.2f}")
print(f"Total Impacto Fiscal 2026: R$ {resumo['Impacto Fiscal 2026 (R$)'].sum():,.2f}")
print(f"Total de Parcelas: {resumo['Quantidade de Parcelas'].sum():,}")
