# Simulação da DBGG — Selic anual = IPCA do ano + 0,37%

**Período:** Jan/2007 a Jun/2026 (estoque inicial: dez/2006).
**Gerado em:** 2026-08-31 01:59 UTC

## Resultado no último mês

- DBGG observada: **R$ 10,81 tri** (R$ 10.809,5 bi).
- DBGG simulada: **R$ 8,36 tri** (R$ 8.364,3 bi).
- Diferença (observada − simulada): **R$ 2,45 tri** (22,62% da DBGG observada).
- Estoque Selic observado: R$ 5,86 tri.
- Estoque Selic simulado: R$ 3,42 tri.
- Juros nominais da parcela Selic no período: observados R$ 4,21 tri; simulados R$ 1,76 tri; economia R$ 2,45 tri.
- DBGG/PIB observada (SGS 4513): **68,48%**.
- DBGG/PIB simulada (mesmo PIB): **52,99%** (−15,49%).

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
| 2007 | 12 | 4,46% | 11,85% | 4,83% | 1.542,9 | 1.505,2 | 37,7 | 63,6 | 25,9 |
| 2008 | 12 | 5,90% | 12,48% | 6,27% | 1.740,9 | 1.660,1 | 80,8 | 81,6 | 38,5 |
| 2009 | 12 | 4,31% | 9,92% | 4,68% | 1.973,4 | 1.844,0 | 129,4 | 85,1 | 36,5 |
| 2010 | 12 | 5,91% | 9,78% | 6,28% | 2.011,5 | 1.845,4 | 166,1 | 81,9 | 45,1 |
| 2011 | 12 | 6,50% | 11,62% | 6,87% | 2.243,6 | 2.027,6 | 216,0 | 95,8 | 45,9 |
| 2012 | 12 | 5,84% | 8,48% | 6,21% | 2.583,9 | 2.334,0 | 249,9 | 77,3 | 43,3 |
| 2013 | 12 | 5,91% | 8,21% | 6,28% | 2.748,0 | 2.462,4 | 285,6 | 85,8 | 50,1 |
| 2014 | 12 | 6,41% | 10,91% | 6,78% | 3.252,4 | 2.908,3 | 344,1 | 107,1 | 48,6 |
| 2015 | 12 | 10,67% | 13,29% | 11,04% | 3.927,5 | 3.517,0 | 410,5 | 170,4 | 104,0 |
| 2016 | 12 | 6,29% | 14,03% | 6,66% | 4.378,5 | 3.818,8 | 559,6 | 232,3 | 83,1 |
| 2017 | 12 | 2,95% | 9,96% | 3,32% | 4.854,7 | 4.147,5 | 707,2 | 195,1 | 47,5 |
| 2018 | 12 | 3,75% | 6,42% | 4,12% | 5.272,0 | 4.487,0 | 785,0 | 140,7 | 62,9 |
| 2019 | 12 | 4,31% | 5,95% | 4,68% | 5.500,1 | 4.647,2 | 852,9 | 153,2 | 85,2 |
| 2020 | 12 | 4,52% | 2,75% | 4,89% | 6.615,8 | 5.787,5 | 828,3 | 79,8 | 104,5 |
| 2021 | 12 | 10,06% | 4,44% | 10,43% | 6.966,9 | 6.239,1 | 727,8 | 138,6 | 239,0 |
| 2022 | 12 | 5,78% | 12,38% | 6,15% | 7.224,9 | 6.252,0 | 972,9 | 396,4 | 151,3 |
| 2023 | 12 | 4,62% | 13,03% | 4,99% | 8.079,3 | 6.769,3 | 1.309,9 | 464,9 | 127,9 |
| 2024 | 12 | 4,83% | 10,89% | 5,20% | 8.984,2 | 7.351,1 | 1.633,1 | 483,6 | 160,4 |
| 2025 | 12 | 4,26% | 14,33% | 4,63% | 10.017,9 | 7.835,1 | 2.182,8 | 697,5 | 147,9 |
| 2026 | 6 | 3,36% | 6,84% | 3,55% | 10.809,5 | 8.364,3 | 2.445,2 | 379,3 | 117,0 |

Δ DBGG = observada − simulada (R$ bilhões). Juros em R$ bilhões.

Em 2020–2021 a Selic observada ficou **abaixo** de IPCA + 0,37% (ciclo de juros reais negativos). Nesses anos a simulação *aumenta* os juros da parcela Selic e a diferença de estoque recua — o que confere o sinal do exercício.

## Premissas e limitações

- Só a remuneração da **parcela indexada à Selic** muda.
- Emissões líquidas (aba `PrimarioR$`) ficam iguais às históricas: o Tesouro não reduz a colocação de LFT/compromissadas além do efeito automático dos juros capitalizados.
- Prefixados, IPCA (NTN-B), câmbio, TR e TJLP/TLP não são reprecificados.
- O PIB usado na razão DBGG/PIB simulada é o mesmo da série oficial (SGS 4513).
