# SEC--data-analysys

Gera fluxos mensais completos (carência + amortização SAC) e impacto fiscal
a valor de 30/06/2026 das operações indiretas automáticas do BNDES (2009–2010),
com **resumo por agente financeiro** e interface web.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
# Amostra rápida (com agentes) — recomendado para validar
python scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --stem fluxos_amostra

# Baixa contratos 2009–2010 (CSV aberto BNDES) e gera fluxos + resumo por agente
python scripts/gerar_fluxos.py --download

# Excel local do portal de transparência (header=5)
python scripts/gerar_fluxos.py --excel caminho/operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx

# Só o ranking (CLI)
python scripts/resumo_por_agente.py --from-output
```

## Versão Web

```bash
# Gere o resumo antes (amostra ou download completo)
python scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv

streamlit run app.py
```

A UI mostra ranking por Instituição Financeira Credenciada, totais de
subsídio e impacto fiscal 2026, busca e download do CSV.

## Saídas

| Arquivo | Conteúdo |
|---------|----------|
| `output/fluxos_completos_corrigido.csv` | Uma linha por parcela |
| `output/fluxos_completos_corrigido.xlsx` | Resumo + **Por_Agente** + impacto mensal + amostra |
| `output/resumo_por_agente.csv` | Ranking: Qtd Contratos, Total Subsídio, Impacto Fiscal 2026 |
| `output/resumo_por_agente.xlsx` | Mesmo ranking em Excel |

## Metodologia

Para cada mês `p = 1 .. (carência + amortização)`:

- `data_fluxo = contratação + p meses`
- `em_carencia = p <= carência` → amortização = 0
- após a carência: `amortização = valor / n` (SAC)
- `subsídio = saldo × (SELIC/12 − juros/12)`
- `impacto = subsídio × (1 + SELIC/12)^(meses até 30/06/2026)`

SELIC anual: **14,5%**.

**Agente** = coluna `Instituição Financeira Credenciada` (Excel) /
`instituicao_financeira_credenciada` (CSV aberto). O vínculo é
`contrato → agente` — não use merge por índice no CSV de parcelas.

## Testes

```bash
PYTHONPATH=. python -m pytest tests/ -q
```
