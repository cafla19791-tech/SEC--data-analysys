# Relatório Executivo - Subsídios BNDES

**Data:** 20/07/2026 23:35
**Total de Contratos:** 20
**Total de Parcelas:** 3,312
**Total Subsídio Nominal:** R$ 1,850,296.83
**Total Impacto Fiscal 2026:** R$ 7,275,421.51

## Principais Agentes

| Instituição Financeira Credenciada               |   Contratos |   Subsídio (R$) |   Impacto 2026 (R$) |
|:-------------------------------------------------|------------:|----------------:|--------------------:|
| BANCO DO BRASIL SA                               |           5 |          515579 |         2.14037e+06 |
| BANCO SANTANDER BRASIL SA                        |           2 |          446012 |         1.63647e+06 |
| ITAU UNIBANCO SA                                 |           2 |          343937 |         1.32122e+06 |
| BANCO BRADESCO SA                                |           3 |          253460 |    987625           |
| BANCO REGIONAL DE DESENVOLVIMENTO DO EXTREMO SUL |           2 |          137539 |    533316           |

## Arquivos gerados

- `resumo_fluxos_polars_final.xlsx` — Contratos, Por_Ano, Por_Agente, Impacto_Por_Ano, Totais_Gerais
- `grafico_interativo.html` — gráfico interativo (Plotly)
- `grafico_top_subsidio.png` — top 10 contratos (Matplotlib)
- `RELATORIO_EXECUTIVO.md` — este relatório

Metodologia ContAgil: impacto = subsídio × fator_final / fator(data_fluxo),
com fator na coluna D do STP (FATOR_30_06_2026 = 82.84819) ou
série Bacen equivalente.
