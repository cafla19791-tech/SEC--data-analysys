# Análise: dívida, juros, estatais, BNDES e PIB per capita PPP

Planilha consolidada respondendo às 10 perguntas sobre DBGG, SELIC/BIS, estatais, desembolsos BNDES e PIB per capita PPP.

## Entregável

`output/analise_divida_juros_estatais_pib.xlsx`

| Aba | Conteúdo |
| --- | --- |
| `Q1_DBGG_fatores` | Evolução anual da DBGG 2002–2026 e fatores condicionantes |
| `Q2_SELIC_acumulada` | SELIC acumulada 02/01/2003–30/06/2026 (BCB SGS 11) |
| `Q3_BIS_taxas_paises` / `Q4_Ranking_taxas` | Taxas básicas acumuladas (BIS WS_CBPOL) e ranking |
| `Q5`–`Q7` | Dívida, juros/resultado financeiro e resultados de Petrobras, Eletrobras/Axia e demais estatais |
| `Q8_BNDES_IPCA` | Desembolsos do Sistema BNDES 2003–2026 atualizados pelo IPCA até 30/06/2026 |
| `Q9_rank_YYYY` | Ranking anual do PIB per capita PPP (constante 2021) |
| `Q10_Var_PPP_2002_2016` | Ranking da variação % 2002–2016 do PIB per capita PPP |

Resumo numérico: `output/respostas_resumo.md`.

## Como regenerar

Os dados brutos (BCB, BIS, CVM, BNDES, World Bank) devem estar em `data/`. Em seguida:

```bash
python3 scripts/build_analise_completa.py
```

## Limitações importantes

1. **DBGG fatores:** detalhamento completo de juros/emissões a partir de 2007 (Dbggindexp desde dez/2006). 2026 = último mês disponível.
2. **BIS:** o arquivo `WS_CBPOL` atual tem ~40 jurisdições com série mensal (não 49).
3. **Estatais:** Petrobras/Eletrobras via CVM DFP 2010–2025; “demais estatais” via BCB (DLSP líquida + NFSP), conceito que já exclui Petrobras, Eletrobras e bancos.
4. **BNDES:** as bases fornecidas cobrem o Sistema BNDES e **não discriminam** BNB/BASA.
5. **BNB/BASA:** sem série própria nos arquivos informados.
