# SEC--data-analysys
#
# Gera fluxos mensais completos (carência + amortização SAC) e impacto fiscal
# a valor de 30/06/2026 das operações indiretas automáticas do BNDES (2009–2010).

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
# Baixa contratos 2009–2010 (CSV aberto BNDES) e gera fluxos
python scripts/gerar_fluxos.py --download

# Excel local do portal de transparência (header=5)
python scripts/gerar_fluxos.py --excel caminho/operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx

# Teste rápido
python scripts/gerar_fluxos.py --download --max-contratos 500
```

## Saídas

| Arquivo | Conteúdo |
|---------|----------|
| `output/fluxos_completos_corrigido.csv` | Uma linha por parcela (saldo, amortização, subsídio, impacto, em_carencia) |
| `output/fluxos_completos_corrigido.xlsx` | Resumo + impacto mensal + amostra de parcelas |
| `output/fluxos_completos_corrigido_stats.json` | Totais da execução |

O CSV detalhado (~1+ GB no período completo) não é versionado; o Excel de resumo cabe no git.

## Metodologia (script corrigido)

Para cada mês `p = 1 .. (carência + amortização)`:

- `data_fluxo = contratação + p meses`
- `em_carencia = p <= carência` → amortização = 0
- após a carência: `amortização = valor / n` (SAC)
- `subsídio = saldo × (SELIC/12 − juros/12)`
- `impacto = subsídio × (1 + SELIC/12)^(meses até 30/06/2026)`

SELIC anual: **14,5%**.

### Correção de carência

O script original misturava dois esquemas:

1. `data = contratação + (carência + p)` com loop `p = 1..n` (pula a carência nas datas)
2. `em_carencia = p <= carência` (zera amortização nas primeiras parcelas do loop)

Isso deixava saldo residual. A versão corrigida gera o cronograma completo
(`carência + n` meses) e só amortiza depois da carência.

## Testes

```bash
PYTHONPATH=. python -m pytest tests/ -q
```
