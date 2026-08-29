# SEC--data-analysys

Gera **fluxos financeiros detalhados** (carência + amortização SAC) e impacto fiscal
a valor de 30/06/2026 das operações indiretas automáticas do BNDES (2009–2010),
com colunas ContAgil (instituição, taxas compostas, spread) e resumo por agente.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
# Amostra rápida (com agentes) — fatores ContAgil via Bacen se não houver STP
python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --stem fluxos_amostra

# Tabela detalhada dia a dia (além das parcelas mensais)
python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv \
  --fluxo-diario --sem-selic-fatores --max-contratos 5

# Script ContAgil (raiz): SAC + SELIC diário exato (fatores col D) + --fluxo-diario
python3 gerar_fluxos.py --excel data/sample_operacoes_com_agente.csv \
  --fluxo-diario --output-dir output --max-contratos 5
# Mesmo CLI ContAgil WinPython (--massa-dados / --pasta-saida):
python3 gerar_fluxos.py \
  --massa-dados "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\dados" \
  --pasta-saida "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida" \
  --arquivo-selic "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\STP-20260716182715078 (1).xlsx"

# Entrypoint ContAgil (mesmo fluxo do script RFB: col D + FATOR_30_06_2026)
python3 scripts/contagil_fluxos.py --input data/sample_operacoes_com_agente.csv
python3 scripts/contagil_fluxos.py --teste-contrato0

# ContAgil WinPython: massa de dados → pasta_saida/fluxos_*.xlsx
python3 scripts/contagil_fluxos.py --massa-dados "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\dados" \
  --pasta-saida "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida" \
  --arquivo-selic "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\STP-20260716182715078 (1).xlsx"
# (--massa-dados e --pasta-dados são equivalentes; sem args: usa esses caminhos se existirem)
# Com tabela dia a dia por arquivo da massa:
python3 scripts/contagil_fluxos.py --massa-dados "...\dados" --pasta-saida "...\saida" \
  --arquivo-selic "...\STP-....xlsx" --fluxo-diario

# Baixa contratos 2009–2010 (CSV aberto BNDES) e gera fluxos detalhados
python3 scripts/gerar_fluxos.py --download

# Excel local do portal / attachments (header=5)
python3 scripts/gerar_fluxos.py --excel caminho/operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx

# Com fatores SELIC ContAgil (STP-*.xlsx exportado da RFB/ContAgil)
python3 scripts/gerar_fluxos.py --download --arquivo-selic "caminho/STP-20260716182715078 (1).xlsx"

# Força download Bacen (já é o padrão quando não há STP)
python3 scripts/gerar_fluxos.py --download --baixar-selic

# Só o ranking (CLI)
python3 scripts/resumo_por_agente.py --from-output

# Impacto fiscal por ano de pagamento (ContAgil: col D, FATOR_30_06_2026)
python3 scripts/impacto_fiscal_por_ano.py --baixar-selic --fluxos output/fluxos_amostra.csv
# Com STP ContAgil local (Windows RFB):
python3 scripts/impacto_fiscal_por_ano.py \
  --arquivo-selic "C:/Arquivos de Programas RFB/ContAgilAppBeta64/python_jep/winpython/STP-20260716182715078 (1).xlsx" \
  --fluxos output/fluxos_completos_final.csv
# Com impacto ContAgil já gravado no CSV:
python3 scripts/impacto_fiscal_por_ano.py --modo coluna --fluxos output/fluxos_completos_final.csv

# Massa BNDES (~70M parcelas): agregação streaming (não carrega tudo na RAM)
# Após contagil_fluxos.py gerar saida/fluxos_*.csv:
python3 scripts/agregar_impacto_fluxos.py \
  --pasta "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida" \
  --modo coluna
# Ou no WinPython: powershell -File .\agregar_impacto_saida.ps1
# Saídas: resumo_impacto_bndes.xlsx, impacto_fiscal_por_ano.xlsx, resumo_por_agente.xlsx

# Discriminativo INDIRETAS por ANO DO CONTRATO (não do fluxo):
# Ex.: contrato 12/12/2022 com 180 parcelas → as 180 linhas na pasta/aba 2022.
# Impacto fiscal de cada parcela continua capitalizado na data_fluxo.
python3 scripts/discriminativos_indiretas_ano_contrato.py --pasta saida
# ContAgil: discriminativos_indiretas_ano_contrato.bat
# Saída: saida/discriminativos_ano_contrato/fluxos_ano_contrato_YYYY.csv

# Discriminativo de taxas básicas de juros reais por país (BIS)
# Fonte: https://data.bis.org/bulkdownload  (WS_CBPOL + WS_LONG_CPI)
# Uma aba por país: Mês/ano, nominal, IPC, inflação no mês, real no mês;
# após dezembro: real acumulada no ano (Fisher composto).
python3 scripts/discriminativo_juros_reais_paises.py
python3 scripts/discriminativo_juros_reais_paises.py --ano-inicio 2000
# ContAgil: discriminativo_juros_reais_paises.bat
# Saída: output/discriminativo_juros_reais_paises.xlsx

# Resumo por contrato + por ano (script ContAgil / WinPython saida/fluxos_0.csv)
python3 scripts/resumo_fluxos.py \
  --fluxos "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida\fluxos_0.csv"
# Repo local (auto-detecta output/fluxos_*):
python3 scripts/resumo_fluxos.py --fluxos output/fluxos_amostra.xlsx
python3 scripts/resumo_fluxos.py --contrato 0

# Resumo avançado ContAgil (pasta saida + Excel original + SELIC col D)
python scripts/resumo_fluxos_avancado.py \
  --pasta "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida" \
  --original "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx" \
  --selic "STP-20260716182715078.xlsx"
# Repo / cloud (amostra + Bacen se não houver STP):
python3 scripts/resumo_fluxos_avancado.py \
  --pasta output \
  --original data/sample_operacoes_com_agente.csv \
  --baixar-selic

# Versão FINAL ContAgil em Polars (SELIC/TJLP mensais — sem SELIC diária)
# No WinPython, copie resumo_fluxos_polars.py (+ pasta scripts/) e rode:
python resumo_fluxos_polars.py \
  --pasta "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida" \
  --original "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx" \
  --selic "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\selic_mensal.xlsx" \
  --tjlp "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\tjlp_mensal.xlsx"
```

Auto-descoberta do STP (nessa ordem): `--arquivo-selic`, env
`CONTAGIL_SELIC`/`SELIC_STP`, caminho ContAgil Windows, `attachments/`,
`data/STP*.xlsx`, `data/selic_fatores_bacen.xlsx`. Sem STP local, o Bacen
SGS 11 é baixado automaticamente (use `--sem-selic-fatores` para o fallback
14,5% composto). Excel de operações também é procurado em
`/home/workdir/attachments/` e `attachments/`.

## Versão Web

```bash
python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv
streamlit run app.py
```

## Saídas

| Arquivo | Conteúdo |
|---------|----------|
| `output/fluxos_completos_final.csv` | Uma linha por parcela (colunas detalhadas) |
| `output/fluxos_completos_final.xlsx` | Resumo + **Por_Agente** + impacto mensal + amostra |
| `output/resumo_por_agente.csv` | Ranking: Qtd Contratos, Total Subsídio, Impacto Fiscal 2026 |
| `output/resumo_por_agente.xlsx` | Mesmo ranking em Excel |
| `output/impacto_fiscal_por_ano.xlsx` | Subsídio + impacto ContAgil agregados por ano de pagamento |
| `output/discriminativo_juros_reais_paises.xlsx` | Uma aba por país: taxa básica nominal BIS, IPC oficial, real mensal e acumulada no ano |
| `resumo_contratos.xlsx` | Por contrato: total subsídio, impacto e saldo final (pasta do CSV de entrada) |
| `resumo_por_ano.xlsx` | Por contrato × ano: total subsídio e impacto |
| `resumo_fluxos_avancado.xlsx` | Workbook ContAgil: Contratos (+ metadados), Por_Ano, Por_Agente, Impacto_Por_Ano, Totais |
| `resumo_fluxos_polars_final.xlsx` | Polars FINAL: Contratos, Por_Ano, Por_Agente, Impacto_Por_Ano, Totais_Gerais |
| `RELATORIO_EXECUTIVO.md` | Relatório markdown (totais + top agentes) |
| `grafico_interativo.html` | Impacto fiscal por ano (Plotly) |
| `output/fluxos_diarios_detalhados.xlsx` | Com `--fluxo-diario`: uma linha por dia entre parcelas |

Colunas do CSV detalhado: `contrato`, `Instituição Financeira`, `mes`,
`data_fluxo`, `saldo_fiscal`, `saldo_contrato`, `amortizacao`,
`taxa_selic_mensal`, `taxa_contrato_mensal` (só na 1ª parcela), `spread`,
`subsidio`, `impacto_fiscal`, `em_carencia`.

Colunas do Excel diário: as mesmas + `taxa_selic_diaria`,
`taxa_contrato_diaria`, `dia_parcela` (amortização só nesse dia).

## Metodologia

Para cada mês `p = 1 .. (carência + amortização)`:

- `data_fluxo` = dia 15 da contratação + `(p − 1)` meses (ContAgil)
- `em_carencia = p <= carência` → amortização = 0
- após a carência: `amortização = valor / n` (SAC)
- `taxa_contrato_efetiva` (mensal):
  - TAXA FIXA / demais: `(1 + juros)^(1/12) − 1`
  - TJLP / TLP: `(1 + 0,06)^(1/12) × (1 + juros)^(1/12) − 1`
- Dual balance:
  - `saldo_fiscal`: só principal (base do subsídio)
  - `saldo_contrato`: principal + juros do contrato
- `spread = (1 + (SELIC_m − taxa_contrato_m))^n`
- `subsídio = saldo_fiscal × (SELIC_m − taxa_contrato_m)` (antes da amortização)
- `impacto_fiscal` (`calcular_impacto_fiscal_real`):
  - STP ContAgil (coluna D):
    `subsídio × FATOR_30_06_2026 / fator(nearest data_parcela)`
    com `FATOR_30_06_2026 = 82.84819`
  - Bacen/outros: `subsídio × fator(nearest 30/06/2026) / fator(nearest data_parcela)`
  - sem fatores: `subsídio × (1 + SELIC_m)^(meses até 30/06/2026)`

SELIC anual de referência (taxas mensais / fallback): **14,5%**.
Fatores ContAgil: coluna A = data, **coluna D** = fator acumulado.
Capitalização: **na data da parcela** (nearest), até o fator de 30/06/2026.

**Correção de carência:** o script ContAgil original misturava
`data = contr + (carência + p)` com `em_carencia = p <= carência` no loop
`p = 1..n`, o que zerava amortização indevidamente. Aqui o cronograma
cobre `carência + n` meses.

**API ContAgil (`gerar_fluxos`):**

```python
df = pd.read_excel("operacoes_....xlsx", sheet_name="operacoes_indiretas_automaticas", header=5)
df_fluxos = gerar_fluxos(df, df)   # 2º arg = df_original (Instituição Financeira)
df_fluxos.to_excel("fluxos_completos_final.xlsx", index=False)
```

O 2º argumento também aceita DataFrame de fatores SELIC (STP/Bacen) ou a
taxa anual. Planilhas brutas (colunas em português) são normalizadas
automaticamente.

## Resultado do run completo (2009–2010)

Com fatores SELIC Bacen (SGS 11) no layout ContAgil (fator nomeado; data da parcela),
na ausência do arquivo STP local `STP-20260716182715078 (1).xlsx`:

| Indicador | Valor |
|-----------|-------|
| Contratos | 348.864 |
| Parcelas | 22.151.051 |
| Subsídio nominal | R$ 23,45 bi |
| Impacto fiscal 30/06/2026 | R$ 90,14 bi |
| Agentes | 72 |

Para usar o STP ContAgil da RFB (capitalização oficial da tabela):

```bash
python3 scripts/gerar_fluxos.py --download \
  --arquivo-selic "C:/Arquivos de Programas RFB/ContAgilAppBeta64/python_jep/winpython/STP-20260716182715078 (1).xlsx"
```

## Testes

```bash
PYTHONPATH=. python3 -m pytest tests/ -q
```
