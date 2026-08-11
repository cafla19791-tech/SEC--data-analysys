# SEC--data-analysys

Repositório de análises e demonstrativos econômicos/fiscais.

## Operações de Financiamento BNDES (Dados Abertos) — mensal

Fonte: [Operações de Financiamento](https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento) (a partir de 2002).

Discriminativo mensal direta × indireta (valor contratado), corrente e IPCA jun/2026.

- Script: `scripts/build_bndes_operacoes_financiamento_mensal.py`
- Excel: `output/bndes_operacoes_financiamento_mensal_ipca.xlsx`

```bash
python3 scripts/build_bndes_operacoes_financiamento_mensal.py
```

## Desembolsos BNDES mensais — Bases de Desembolso (jan/1995–jun/2026)

Discriminativo mensal a partir das Bases de Desembolso do Sistema BNDES.

- Script: `scripts/build_bndes_desembolsos_mensal.py`
- Excel: `output/bndes_desembolsos_mensal_direta_indireta_ipca.xlsx`

```bash
python3 scripts/build_bndes_desembolsos_mensal.py
```

IPCA: série `PRECOS12_IPCA12` do [Ipeadata](http://www.ipeadata.gov.br/).
