# SEC — Análise de dados / Fluxos de pagamento BNDES

Gera fluxos de pagamento (sistema SAC) e estimativa de subsídio implícito
(Selic de referência vs taxa do contrato) a partir das **operações indiretas
automáticas** do BNDES no período **2009–2010**.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

Baixar/filtrar dados abertos do BNDES e gerar fluxos:

```bash
python scripts/gerar_fluxos.py --download
```

Usar um CSV já filtrado em `data/`:

```bash
python scripts/gerar_fluxos.py --input data/operacoes_indiretas_automaticas_2009-2010.csv
```

Usar planilha Excel no formato de transparência (header na linha 6):

```bash
python scripts/gerar_fluxos.py --input caminho/para/operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx
```

Teste rápido (primeiros N contratos):

```bash
python scripts/gerar_fluxos.py --download --limit-contracts 100
```

## Saídas (`output/`)

| Arquivo | Conteúdo |
|---------|----------|
| `fluxos_gerados_corrigido.parquet` / `.csv` | Parcelas completas |
| `fluxos_gerados_corrigido.xlsx` | Mesmo conteúdo (se ≤ 1M linhas) |
| `fluxos_gerados_corrigido_resumo.xlsx` | Totais por contrato |

Colunas das parcelas: `Contrato_ID`, `Parcela`, `Data_Pagamento`,
`Valor_Amortizacao`, `Juros_Parcela`, `Saldo_Devedor`, `Subsídio`.

## Metodologia

- Amortização constante (SAC): `valor / n`
- Juros da parcela: `saldo × (taxa_aa / 12)`
- Subsídio implícito: `(Selic_aa/12 − taxa_aa/12) × saldo` (Selic padrão: **14,5% a.a.**)
- Data da parcela: `data_contratação + (carência + número_da_parcela)` meses

Fonte dos dados: [Portal de Dados Abertos do BNDES — Operações indiretas automáticas](https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento).
