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
# Amostra rápida (com agentes) — recomendado para validar
python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --stem fluxos_amostra

# Baixa contratos 2009–2010 (CSV aberto BNDES) e gera fluxos detalhados
python3 scripts/gerar_fluxos.py --download

# Excel local do portal de transparência (header=5)
python3 scripts/gerar_fluxos.py --excel caminho/operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx

# Com fatores SELIC ContAgil (STP-*.xlsx exportado da RFB/ContAgil)
python3 scripts/gerar_fluxos.py --download --arquivo-selic "caminho/STP-20260716182715078 (1).xlsx"

# Sem STP local: baixa SELIC diária do Bacen (SGS 11) e monta fatores acumulados
python3 scripts/gerar_fluxos.py --download --baixar-selic

# Só o ranking (CLI)
python3 scripts/resumo_por_agente.py --from-output
```

Auto-descoberta do STP (nessa ordem): `--arquivo-selic`, env
`CONTAGIL_SELIC`/`SELIC_STP`, caminho ContAgil Windows, `data/STP*.xlsx`,
`data/selic_fatores_bacen.xlsx`.

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

Colunas do CSV detalhado: `contrato`, `Instituição Financeira`, `mes`,
`data_fluxo`, `saldo`, `amortizacao`, `taxa_selic_mensal`,
`taxa_contrato_mensal`, `spread`, `subsidio`, `impacto_fiscal`, `em_carencia`.

## Metodologia

Para cada mês `p = 1 .. (carência + amortização)`:

- `data_fluxo` = dia 15 da contratação + `(p − 1)` meses (ContAgil)
- `em_carencia = p <= carência` → amortização = 0
- após a carência: `amortização = valor / n` (SAC)
- `taxa_*_mensal = (1 + taxa_aa)^(1/12) − 1`
- `spread = (1 + (SELIC_m − taxa_contrato_m))^n`
- `subsídio = saldo × (SELIC_m − taxa_contrato_m)`
- `impacto_fiscal` (`calcular_impacto_fiscal_real`):
  - com fatores (STP ContAgil ou Bacen): `subsídio × fator(nearest 30/06/2026) / fator(nearest data_fluxo)`
  - sem fatores: `subsídio × (1 + SELIC_m)^(meses até 30/06/2026)`

SELIC anual de referência (taxas mensais / fallback): **14,5%**.
Fatores ContAgil: coluna A = data, coluna C = fator acumulado.

**Correção de carência:** o script ContAgil original misturava
`data = contr + (carência + p)` com `em_carencia = p <= carência` no loop
`p = 1..n`, o que zerava amortização indevidamente. Aqui o cronograma
cobre `carência + n` meses.

## Testes

```bash
PYTHONPATH=. python3 -m pytest tests/ -q
```
