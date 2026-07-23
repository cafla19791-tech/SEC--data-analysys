# Relatório Executivo - Subsídios BNDES

**Data:** 23/07/2026 01:26
**Total de Contratos:** 20
**Total de Parcelas:** 3,312
**Total Subsídio Nominal:** R$ 1,850,296.83
**Total Impacto Fiscal 2026:** R$ 8,858,951.38

## Principais Agentes

| Instituição Financeira Credenciada               |   Contratos |   Subsídio (R$) |   Impacto 2026 (R$) |
|:-------------------------------------------------|------------:|----------------:|--------------------:|
| BANCO DO BRASIL SA                               |           5 |          515579 |         2.62495e+06 |
| BANCO SANTANDER BRASIL SA                        |           2 |          446012 |         1.97578e+06 |
| ITAU UNIBANCO SA                                 |           2 |          343937 |         1.60246e+06 |
| BANCO BRADESCO SA                                |           3 |          253460 |         1.20068e+06 |
| BANCO REGIONAL DE DESENVOLVIMENTO DO EXTREMO SUL |           2 |          137539 |    647805           |

## Arquivos gerados

- `resumo_fluxos_polars_final.xlsx` — Contratos, Por_Ano, Por_Agente, Impacto_Por_Ano, Totais_Gerais
- `grafico_interativo.html` — gráfico interativo (Plotly)
- `grafico_top_subsidio.png` — top 10 contratos (Matplotlib)
- `RELATORIO_EXECUTIVO.md` — este relatório

Metodologia ContAgil (mensal): impacto = subsídio × fator_selic(30/06/2026)
/ fator_selic(mês da parcela), com `selic_mensal.xlsx` e `tjlp_mensal.xlsx`.
