Gere o CSV detalhado e o resumo por agente com:

```bash
# amostra (com fatores SELIC Bacen)
python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --baixar-selic --stem fluxos_amostra

# período completo com fatores ContAgil/Bacen
python3 scripts/gerar_fluxos.py --download --baixar-selic
```

Arquivos principais:

- `fluxos_completos_final.xlsx` — abas Resumo, **Por_Agente**, Impacto_Mensal, Amostra_Parcelas (run completo 2009–2010)
- `resumo_por_agente.csv` / `.xlsx` — ranking por Instituição Financeira Credenciada
- `impacto_fiscal_por_ano.xlsx` / `.csv` — impacto capitalizado até 30/06/2026 por ano de pagamento
- `resumo_fluxos_avancado.xlsx` — workbook ContAgil (Contratos, Por_Ano, Por_Agente, Impacto_Por_Ano, Totais)
- `resumo_fluxos_polars_final.xlsx` — Polars FINAL (SELIC/TJLP mensais + Totais_Gerais)
- `RELATORIO_EXECUTIVO.md` / `grafico_interativo.html` / `grafico_top_subsidio.png`
- `amostra_fluxos_detalhados.xlsx` — primeiras parcelas com colunas ContAgil
- `fluxos_amostra.xlsx` — amostra rápida (20 contratos de exemplo)

```bash
# ContAgil (fatores Bacen se não houver STP local)
python3 scripts/impacto_fiscal_por_ano.py --baixar-selic --fluxos output/fluxos_amostra.csv

# STP ContAgil da RFB (coluna D)
python3 scripts/impacto_fiscal_por_ano.py \
  --arquivo-selic "C:/Arquivos de Programas RFB/ContAgilAppBeta64/python_jep/winpython/STP-20260716182715078 (1).xlsx" \
  --fluxos output/fluxos_completos_final.csv
```

Colunas detalhadas: Instituição Financeira, taxa_selic_mensal, taxa_contrato_mensal,
spread, subsidio, impacto_fiscal, em_carencia.

Run completo (fatores SELIC Bacen SGS 11 → ContAgil): 348.864 contratos ·
22.151.051 parcelas · subsídio R$ 23,45 bi · impacto fiscal 2026 R$ 90,17 bi ·
72 agentes.

Interface web: `streamlit run app.py`

Montantes anuais da modalidade indireta (2002–2026), a partir dos microdados
abertos do BNDES + série oficial de aprovações:

```bash
python3 scripts/montantes_contratos_indiretas.py
```

- `montantes_contratos_indiretas_2002_2026.csv` / `.xlsx` / `.md`

Dívida bruta da Petrobras nos Forms 20-F (2002–2025), com página e link do EDGAR:

```bash
python3 scripts/petrobras_divida_bruta_20f.py
```

- `petrobras_divida_bruta_20f_2002_2025.csv` / `.xlsx` / `.md` / `.png`

Juros pagos pela Petrobras (caixa) nos Forms 20-F 2002–2025 e no 6-K do 1S2026:

```bash
python3 scripts/petrobras_juros_pagos_20f.py
```

- `petrobras_juros_pagos_20f_2002_2026.csv` / `.xlsx` / `.md` / `.png`
