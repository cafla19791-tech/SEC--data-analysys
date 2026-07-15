# SEC — Análise de dados (BNDES)

Scripts para gerar fluxos de amortização e impacto fiscal das **operações
indiretas automáticas** do BNDES (período 2009–2010), a valor de 30/06/2026.

## Setup

```bash
pip install -r requirements.txt
```

## Uso

Baixa os contratos 2009–2010 pela API do portal de dados abertos do BNDES,
gera o detalhe por parcela e um Excel de resumo:

```bash
python scripts/gerar_fluxos.py
```

Opções úteis:

```bash
# Teste rápido com 500 contratos
python scripts/gerar_fluxos.py --max-contratos 500

# Usar um Excel local no formato do portal (header na linha 6)
python scripts/gerar_fluxos.py --excel caminho/operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx

# Recalcular offsets do datastore (se a base BNDES mudar)
python scripts/gerar_fluxos.py --discover-offsets
```

## Saídas

| Arquivo | Conteúdo |
|---------|----------|
| `output/fluxos_gerados.csv` | Uma linha por parcela (Amortização, Saldo, Subsídio, Impacto_Fiscal_2026) |
| `output/fluxos_gerados.xlsx` | Resumo + agregação mensal do impacto + amostra de parcelas |
| `data/operacoes_2009_2010.parquet` | Cache local das operações baixadas |

## Metodologia (igual ao script de referência)

Para cada parcela `p = 1..n` após a carência:

- `amort = valor / n`
- `subsídio = (SELIC/12 − juros/12) × saldo_devedor`
- `impacto_2026 = subsídio × (1 + SELIC/12)^(meses até 30/06/2026)`

SELIC anual configurada: **14,5%** (`TAXA_SELIC_ANUAL` em `scripts/gerar_fluxos.py`).
