# Simulação da DBGG — Selic anual = IPCA do ano + 0,37%

**Período:** Jan/2003 a Jul/2026 (estoque inicial: dez/2006).
**Gerado em:** 2026-08-31 02:53 UTC

## Resultado no último mês

- DBGG observada: **R$ 10,88 tri** (R$ 10.881,0 bi).
- DBGG simulada: **R$ 7,30 tri** (R$ 7.299,1 bi).
- Diferença (observada − simulada): **R$ 3,58 tri** (32,92% da DBGG observada).
- Estoque Selic observado: R$ 5,93 tri.
- Estoque Selic simulado: R$ 2,35 tri.
- Juros nominais da parcela Selic no período: observados R$ 4,70 tri; simulados R$ 1,12 tri; economia R$ 3,58 tri.

## Metodologia

Fonte da dívida: planilha especial do Banco Central [Dbggindexp.xlsx](https://www.bcb.gov.br/content/estatisticas/Documents/Tabelas_especiais/Dbggindexp.xlsx) — abas `DividaR$` (estoques), `JurosR$` (juros nominais mensais) e `PrimarioR$` (emissões líquidas). Arquivo usado: `data/Dbggindexp.xlsx`.

A coluna **Selic** do BCB reúne LFT, LFT-A, LFT-B, operações compromissadas (recompra e revenda), dívida bancária e securitizadas indexadas à Selic.

Hipótese: em cada ano civil a Selic acumulada no ano é igual ao IPCA acumulado no mesmo conjunto de meses **mais 0,37% proporcionais** (`spread × n/12`). A taxa mensal contrafactual é a equivalente composta, constante dentro do ano. Em 2026 o ano está incompleto (janeiro–junho): o IPCA e a Selic contrafactual são os acumulados do semestre, com o spread de 0,37% × 6/12 = 0,185%.

Os juros Selic observados (contabilidade BCB) são reescalonados pela razão entre as taxas mensais (SGS 4390 vs. contrafactual) e pela razão dos estoques Selic do mês anterior (efeito de estoque menor). As emissões líquidas por indexador permanecem as observadas. O resíduo de fechamento do BCB (`Δestoque − emissão − juros`) é replicado no cenário simulado, para não atribuir quebras estatísticas à Selic.

A DBGG simulada é a observada menos a diferença de estoque da parcela Selic. Não há efeitos de segunda ordem (PIB, câmbio, cupons dos prefixados, NTN-B, resultado primário ou composição das emissões). Em um equilíbrio geral os prefixados novos também sairiam mais baratos; esta simulação é, portanto, um **piso** para a redução da dívida.

Séries auxiliares: IPCA SGS 433 (% a.m.); Selic SGS 4390 (% a.m.); DBGG/PIB SGS 4513 (quando disponível).

## Série anual

Valores de estoque e DBGG no **último mês** de cada ano na amostra (dezembro; em 2026, junho). Juros e IPCA/Selic são acumulados no ano. Unidades: R$ bilhões (estoques e juros) e % a.a. (ou % no período, quando n < 12).

| Ano | n | IPCA | Selic obs. | Selic cf. | DBGG obs. | DBGG cf. | Δ DBGG | Juros Selic obs. | Juros cf. |
|----:|--:|-----:|-----------:|----------:|-----------:|---------:|-------:|-----------------:|----------:|
| 2003 | 12 | 9,30% | 23,33% | 9,67% | 1.112,3 | 1.031,7 | 80,6 | 136,9 | 56,3 |
| 2004 | 12 | 7,60% | 16,24% | 7,97% | 1.172,5 | 1.036,5 | 136,0 | 96,3 | 40,9 |
| 2005 | 12 | 5,69% | 19,04% | 6,06% | 1.241,2 | 1.020,7 | 220,4 | 111,7 | 27,3 |
| 2006 | 12 | 3,14% | 15,08% | 3,51% | 1.336,6 | 1.050,2 | 286,4 | 75,6 | 9,6 |
| 2007 | 12 | 4,46% | 11,85% | 4,83% | 1.542,9 | 1.204,8 | 338,1 | 63,6 | 12,0 |
| 2008 | 12 | 5,90% | 12,48% | 6,27% | 1.740,9 | 1.340,6 | 400,3 | 81,6 | 19,4 |
| 2009 | 12 | 4,31% | 9,92% | 4,68% | 1.973,4 | 1.509,4 | 464,0 | 85,1 | 21,3 |
| 2010 | 12 | 5,91% | 9,78% | 6,28% | 2.011,5 | 1.490,6 | 520,9 | 81,9 | 24,9 |
| 2011 | 12 | 6,50% | 11,62% | 6,87% | 2.243,6 | 1.649,5 | 594,1 | 95,8 | 22,6 |
| 2012 | 12 | 5,84% | 8,48% | 6,21% | 2.583,9 | 1.932,6 | 651,4 | 77,3 | 20,0 |
| 2013 | 12 | 5,91% | 8,21% | 6,28% | 2.748,0 | 2.034,7 | 713,2 | 85,8 | 24,0 |
| 2014 | 12 | 6,41% | 10,91% | 6,78% | 3.252,4 | 2.454,0 | 798,4 | 107,1 | 22,0 |
| 2015 | 12 | 10,67% | 13,29% | 11,04% | 3.927,5 | 3.014,4 | 913,2 | 170,4 | 55,6 |
| 2016 | 12 | 6,29% | 14,03% | 6,66% | 4.378,5 | 3.282,9 | 1.095,6 | 232,3 | 49,9 |
| 2017 | 12 | 2,95% | 9,96% | 3,32% | 4.854,7 | 3.594,1 | 1.260,6 | 195,1 | 30,1 |
| 2018 | 12 | 3,75% | 6,42% | 4,12% | 5.272,0 | 3.912,0 | 1.359,9 | 140,7 | 41,3 |
| 2019 | 12 | 4,31% | 5,95% | 4,68% | 5.500,1 | 4.046,1 | 1.454,0 | 153,2 | 59,2 |
| 2020 | 12 | 4,52% | 2,75% | 2,75% | 6.615,8 | 5.122,2 | 1.493,5 | 79,8 | 40,3 |
| 2021 | 12 | 10,06% | 4,44% | 4,44% | 6.966,9 | 5.405,0 | 1.561,9 | 138,6 | 70,2 |
| 2022 | 12 | 5,78% | 12,38% | 6,15% | 7.224,9 | 5.364,9 | 1.859,9 | 396,4 | 98,4 |
| 2023 | 12 | 4,62% | 13,03% | 4,99% | 8.079,3 | 5.836,5 | 2.242,8 | 464,9 | 82,1 |
| 2024 | 12 | 4,83% | 10,89% | 5,20% | 8.984,2 | 6.367,9 | 2.616,4 | 483,6 | 110,0 |
| 2025 | 12 | 4,26% | 14,33% | 4,63% | 10.017,9 | 6.805,2 | 3.212,7 | 697,5 | 101,2 |
| 2026 | 7 | 3,44% | 8,14% | 3,65% | 10.881,0 | 7.299,1 | 3.581,9 | 450,8 | 81,7 |

Δ DBGG = observada − simulada (R$ bilhões). Juros em R$ bilhões.

## Discriminativo das reduções por ano

A **redução no ano** é a economia de juros nominais da parcela Selic naquele exercício (fluxo). A **redução acumulada** é a diferença de estoque da DBGG no último mês do ano. Participação = redução do ano / soma das reduções do período.

| Ano | Selic obs. | Selic cf. | Juros Selic obs. | Juros cf. | Redução no ano | Redução acum. | Part. |
|----:|-----------:|----------:|-----------------:|----------:|---------------:|--------------:|------:|
| 2003 | 23,33% | 9,67% | 136,9 | 56,3 | 80,6 | 80,6 | 2,25% |
| 2004 | 16,24% | 7,97% | 96,3 | 40,9 | 55,4 | 136,0 | 1,55% |
| 2005 | 19,04% | 6,06% | 111,7 | 27,3 | 84,4 | 220,4 | 2,36% |
| 2006 | 15,08% | 3,51% | 75,6 | 9,6 | 66,0 | 286,4 | 1,84% |
| 2007 | 11,85% | 4,83% | 63,6 | 12,0 | 51,7 | 338,1 | 1,44% |
| 2008 | 12,48% | 6,27% | 81,6 | 19,4 | 62,2 | 400,3 | 1,74% |
| 2009 | 9,92% | 4,68% | 85,1 | 21,3 | 63,7 | 464,0 | 1,78% |
| 2010 | 9,78% | 6,28% | 81,9 | 24,9 | 56,9 | 520,9 | 1,59% |
| 2011 | 11,62% | 6,87% | 95,8 | 22,6 | 73,2 | 594,1 | 2,04% |
| 2012 | 8,48% | 6,21% | 77,3 | 20,0 | 57,3 | 651,4 | 1,60% |
| 2013 | 8,21% | 6,28% | 85,8 | 24,0 | 61,9 | 713,2 | 1,73% |
| 2014 | 10,91% | 6,78% | 107,1 | 22,0 | 85,2 | 798,4 | 2,38% |
| 2015 | 13,29% | 11,04% | 170,4 | 55,6 | 114,8 | 913,2 | 3,20% |
| 2016 | 14,03% | 6,66% | 232,3 | 49,9 | 182,4 | 1.095,6 | 5,09% |
| 2017 | 9,96% | 3,32% | 195,1 | 30,1 | 165,0 | 1.260,6 | 4,61% |
| 2018 | 6,42% | 4,12% | 140,7 | 41,3 | 99,4 | 1.359,9 | 2,77% |
| 2019 | 5,95% | 4,68% | 153,2 | 59,2 | 94,0 | 1.454,0 | 2,63% |
| 2020 | 2,75% | 2,75% | 79,8 | 40,3 | 39,5 | 1.493,5 | 1,10% |
| 2021 | 4,44% | 4,44% | 138,6 | 70,2 | 68,4 | 1.561,9 | 1,91% |
| 2022 | 12,38% | 6,15% | 396,4 | 98,4 | 298,0 | 1.859,9 | 8,32% |
| 2023 | 13,03% | 4,99% | 464,9 | 82,1 | 382,8 | 2.242,8 | 10,69% |
| 2024 | 10,89% | 5,20% | 483,6 | 110,0 | 373,6 | 2.616,4 | 10,43% |
| 2025 | 14,33% | 4,63% | 697,5 | 101,2 | 596,4 | 3.212,7 | 16,65% |
| 2026 | 8,14% | 3,65% | 450,8 | 81,7 | 369,1 | 3.581,9 | 10,30% |
| **Total** | | | 4.701,9 | 1.120,0 | 3.581,9 | 3.581,9 | 100,00% |

Valores em R$ bilhões. Sinal negativo = a Selic observada ficou abaixo de IPCA + spread (a simulação *aumenta* os juros naquele ano).

Em 2020–2021 a Selic observada ficou **abaixo** de IPCA + 0,37% (ciclo de juros reais negativos). Nesses anos a simulação *aumenta* os juros da parcela Selic e a diferença de estoque recua — o que confere o sinal do exercício.

## Premissas e limitações

- Só a remuneração da **parcela indexada à Selic** muda.
- Emissões líquidas (aba `PrimarioR$`) ficam iguais às históricas: o Tesouro não reduz a colocação de LFT/compromissadas além do efeito automático dos juros capitalizados.
- Prefixados, IPCA (NTN-B), câmbio, TR e TJLP/TLP não são reprecificados.
- O PIB usado na razão DBGG/PIB simulada é o mesmo da série oficial (SGS 4513).
