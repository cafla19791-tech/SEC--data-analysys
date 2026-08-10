# Contratações do FNO (2011–2026) — Relatórios da Administração BASA

Portal: https://ri.bancoamazonia.com.br/informacoes-financeiras/central-de-resultados/

Arquivo: `output/fno_contratacoes_2011_2026_basa_ra.xlsx`

## Série (valores correntes e IPCA 30/06/2026)

| Ano | Corrente (R$ bi) | IPCA 30/06/2026 (R$ bi) | Cobertura | Conceito | Documento |
|----:|-----------------:|------------------------:|:----------|:---------|:----------|
| 2011 | 1.870 | 4.309 | ano completo | contratado | DF/RA 4T11 |
| 2012 | 4.300 | 9.400 | ano completo | contratado | DF/RA 4T12 |
| 2013 | 4.722 | 9.719 | ano completo | aplicado/contratado | DF/RA 4T13 |
| 2014 | 5.367 | 10.388 | ano completo | aplicado | DF/RA 4T14 |
| 2015 | 3.965 | 7.039 | ano completo | contratado | DF/RA 4T16 (tabela YoY) |
| 2016 | 2.334 | 3.811 | ano completo | contratado | DF/RA 4T16 |
| 2017 | 2.906 | 4.587 | ano completo | contratado | DF/RA 4T17 |
| 2018 | 4.636 | 7.059 | ano completo | contratado | DF/RA 4T18 |
| 2019 | 7.671 | 11.259 | ano completo | contratado | DF/RA 4T19 |
| 2020 | 10.500 | 14.932 | ano completo | contratado | RA 4T21 (comparativo YoY) |
| 2021 | 12.500 | 16.414 | ano completo | contratado/aplicado | RA 4T21 |
| 2022 | 12.000 | 14.419 | ano completo | contratado/aplicado | RA 4T22 |
| 2023 | 11.300 | 12.982 | ano completo | aplicado/disponibilizado | RA 4T23 |
| 2024 | 13.600 | 14.970 | ano completo | aplicado/contratado | RA 4T24 |
| 2025 | 17.800 | 18.657 | ano completo | aplicado/contratado | RA 4T25 |
| 2026 | 2.600 | 2.633 | 1T26 (parcial) | contratado (FNO no fomento) | RA 1T26 |

Soma 2011–2025 (anos cheios), correntes: R$ 115.47 bi
Soma 2011–2025 (anos cheios), IPCA 30/06/2026: R$ 159.94 bi

## Notas

- 2015: usa TOTAL de **contratações** da tabela do RA 2016 (R$ 3.964,9 mi); o RA 2015 cita R$ 5.068,4 mi em linguagem de **liberado**.
- 2020: o pacote DF 4T20 do RI não traz o total FNO; valor R$ 10,5 bi vem do comparativo do RA 2021 (MDR/SUDAM: R$ 10.486,0 mi).
- 2021: RA do exercício = R$ 12,5 bi; RA 2022 cita R$ 13,3 bi aplicados em 2021 (divergência documentada).
- 2026: apenas 1T26 — FNO ≈ R$ 2,6 bi no gráfico de fomento contratado.

Regenerar:

```bash
python3 scripts/build_fno_basa_ra.py
```
