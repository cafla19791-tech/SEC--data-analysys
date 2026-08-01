# Relatório Executivo - Subsídios BNDES

**Data:** 01/08/2026 01:25
**Total de Contratos:** 23
**Total de Parcelas:** 1,203
**Total Subsídio Nominal:** R$ 859,854.93
**Total Impacto Fiscal 2026:** R$ 4,232,480.09

## Principais Agentes

| Instituição Financeira Credenciada   |   Contratos |   Subsídio (R$) |   Impacto 2026 (R$) |
|:-------------------------------------|------------:|----------------:|--------------------:|
| BNDES                                |           3 |        251351   |         1.31772e+06 |
| BANCO DO BRASIL SA                   |           5 |        171860   |    874982           |
| BANCO SANTANDER BRASIL SA            |           2 |        145344   |    643867           |
| ITAU UNIBANCO SA                     |           2 |        112879   |    525274           |
| BANCO BRADESCO SA                    |           3 |         81317.3 |    385604           |

## Arquivos gerados

- `resumo_fluxos_polars_final.xlsx` — Contratos, Por_Ano, Por_Agente, Impacto_Por_Ano, Totais_Gerais
- `grafico_interativo.html` — gráfico interativo (Plotly)
- `grafico_top_subsidio.png` — top 10 contratos (Matplotlib)
- `RELATORIO_EXECUTIVO.md` — este relatório

Metodologia ContAgil (mensal): impacto = subsídio × fator_selic(30/06/2026)
/ fator_selic(mês da parcela), com `selic_mensal.xlsx` e `tjlp_mensal.xlsx`.
