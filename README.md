# SEC--data-analysys

Repositório de análises e demonstrativos econômicos/fiscais.

## Desembolsos BNDES mensais — Direta × Indireta (jan/1995–jun/2026)

Discriminativo mensal a partir das Bases de Desembolso do Sistema BNDES, com valores correntes e atualizados pelo IPCA (Ipeadata) até jun/2026.

Colunas:

| Mês/ano | Direta corrente | Direta atual IPCA | Indireta corrente | Indireta atual IPCA |

- Script: `scripts/build_bndes_desembolsos_mensal.py`
- Excel: `output/bndes_desembolsos_mensal_direta_indireta_ipca.xlsx`
- Resumo: `output/bndes_desembolsos_mensal_direta_indireta_ipca.md`

```bash
python3 scripts/build_bndes_desembolsos_mensal.py
```

IPCA: série `PRECOS12_IPCA12` do [Ipeadata](http://www.ipeadata.gov.br/).
