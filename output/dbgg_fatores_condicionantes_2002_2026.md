# Discriminativo dos fatores condicionantes da DBGG (2002–2026) — em R$

Valores dos fatores em **R$ bilhões**.  
O efeito do PIB nominal **não altera o estoque em reais** (só a razão dívida/PIB); a coluna `PIB*` é o **equivalente em R$** do impacto sobre a razão (`pp/100 × PIB` do ano), para comparação de magnitudes.

**Conversão:** 2002–2017 usam fluxos oficiais em R$ da BCB NT 47; 2018–2026 convertem os p.p. das Notas de Estatísticas Fiscais pelo PIB implícito (`estoque ÷ %PIB`) da própria nota.

**2026:** acumulado até julho.

Arquivos: [`dbgg_fatores_condicionantes_2002_2026.csv`](dbgg_fatores_condicionantes_2002_2026.csv) · [`dbgg_fatores_condicionantes_2002_2026_reais.csv`](dbgg_fatores_condicionantes_2002_2026_reais.csv) (R$ milhões).

## Metodologia antiga (até 2007) — R$ bi

| Ano | Estoque | %PIB | Juros | Emissões | Reconhec. | Câmbio | PIB* |
|---:|---:|---:|---:|---:|---:|---:|---:|
| 2002 | 1.132,9 | 76,1 | 182,6 | −77,5 | 2,8 | 139,1 | −116,1 |
| 2003 | 1.228,6 | 71,5 | 179,2 | −20,1 | −14,4 | −49,0 | −173,5 |
| 2004 | 1.331,8 | 68,0 | 171,3 | −58,6 | 5,8 | −15,3 | −170,4 |
| 2005 | 1.453,6 | 67,0 | 206,6 | −46,0 | −5,0 | −33,8 | −143,2 |

## Metodologia atual — R$ bi

| Ano | Estoque | %PIB | Juros | Emissões | Reconhec. | Câmbio | Demais | PIB* |
|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 2006 | 1.336,6 | 55,5 | — | — | — | — | — | base |
| 2007 | 1.542,9 | 56,7 | 179,1 | 43,0 | 4,4 | −20,3 | | −176,9 |
| 2008 | 1.740,9 | 56,0 | 200,9 | −42,0 | 3,2 | 35,8 | | −223,8 |
| 2009 | 1.973,4 | 59,2 | 191,4 | 74,3 | 1,9 | −35,1 | | −123,3 |
| 2010 | 2.011,5 | 51,8 | 216,5 | −180,9 | 4,7 | −2,3 | | −322,3 |
| 2011 | 2.243,6 | 51,3 | 253,8 | −40,1 | 7,4 | 11,0 | | −258,0 |
| 2012 | 2.583,9 | 53,7 | 248,9 | 77,7 | 3,5 | 10,3 | | −226,2 |
| 2013 | 2.748,0 | 51,5 | 273,2 | −130,9 | 0,8 | 20,9 | | −277,5 |
| 2014 | 3.252,4 | 56,3 | 313,2 | 168,7 | 0,0 | 22,5 | | −225,3 |
| 2015 | 3.927,5 | 65,5 | 447,0 | 135,8 | 5,7 | 86,6 | | −125,9 |
| 2016 | 4.378,5 | 70,0 | 511,6 | −18,5 | 6,5 | −48,6 | | −168,9 |
| 2017 | 4.854,7 | 74,0 | 439,8 | 24,7 | 9,1 | 2,6 | | −209,9 |
| 2018 | 5.272,0 | 76,7 | 405,5 | | | 48,1 | −34,4 | −240,6 |
| 2019 | 5.500,1 | 75,8 | 406,3 | −195,9 | | 14,5 | | −283,0 |
| 2020 | 6.615,8 | 89,3 | 348,2 | 674,2 | | 96,3 | −7,4 | |
| 2021 | 7.000,0 | 80,3 | 505,6 | −183,1 | | 34,9 | | −1.072,2 |
| 2022 | 7.200,0 | 73,5 | 734,7 | −440,8 | | | −29,4 | −734,7 |
| 2023 | 8.100,0 | 74,3 | 817,6 | 65,4 | | −32,7 | | −566,9 |
| 2024 | 9.000,0 | 76,1 | 887,0 | −106,4 | 35,5 | 118,3 | −23,7 | −638,6 |
| 2025 | 10.000,0 | 78,7 | 1.130,9 | −38,1 | 25,4 | −63,5 | −25,4 | −724,3 |
| 2026* | 10.900,0 | 82,5 | 753,1 | 198,2 | | −39,6 | | −409,6 |

\*Até jul/2026. Estoques 2021–2026 das notas anuais/mensais (arredondados na fonte quando a nota cita “R$ X trilhões”).

## Destaques em reais

- **Juros** passam de ~R$ 180–270 bi/ano (2007–13) para **R$ 1,13 tri em 2025**.
- **2020:** emissões líquidas de **R$ 674 bi** (principal motor do salto).
- **2021:** efeito PIB equivalente a **−R$ 1,07 tri** na razão (forte diluição).
- **2022:** resgates líquidos de **−R$ 441 bi**.
- **2026 (até jul.):** juros **R$ 753 bi** + emissões **R$ 198 bi**, parcialmente compensados pelo PIB (**−R$ 410 bi** eq.).

## Fontes

- BCB Nota Técnica nº 47 (jul/2018), Tabelas 7 e 8 (R$ milhões oficiais)
- BCB Notas de Estatísticas Fiscais (2019–2026), p.p. convertidos pelo PIB implícito da nota
