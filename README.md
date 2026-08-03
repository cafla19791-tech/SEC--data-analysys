# Série histórica RTN — Mai/2026

Extrator da planilha oficial do Resultado do Tesouro Nacional:

`C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\serie_historica_mai26 (2).xlsx`

| Aba | Conteúdo |
|-----|----------|
| `1.1` | Valores **correntes** (R$ milhões) |
| `1.1-A` | Valores **constantes IPCA de Mai/2026** |

## ContAgil (WinPython)

```bat
cd /d "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
mkdir scripts 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Invoke-WebRequest -UseBasicParsing -Uri 'https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/serie-historica-mai26-d84a/scripts/extrair_serie_historica_rtn.py' -OutFile 'scripts\extrair_serie_historica_rtn.py'; Invoke-WebRequest -UseBasicParsing -Uri 'https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/serie-historica-mai26-d84a/extrair_rtn_mai26.bat' -OutFile 'extrair_rtn_mai26.bat'"
extrair_rtn_mai26.bat
```

Ou direto:

```bat
.\python.exe scripts\extrair_serie_historica_rtn.py "serie_historica_mai26 (2).xlsx" --constantes-ipca --out saida\rtn_anual_ipca_mai26.csv
```

## Resultado anual IPCA Mai/2026 (R$ bi)

| Ano | Primário | Juros | Nominal | Receita | Despesa |
|----:|--------:|------:|--------:|--------:|--------:|
| 2001 | 96.77 | -205.60 | -107.46 | 1181.28 | 896.24 |
| 2002 | 129.74 | -170.34 | -38.86 | 1288.11 | 948.63 |
| 2003 | 138.92 | -355.23 | -217.41 | 1251.04 | 911.91 |
| 2004 | 164.25 | -261.61 | -87.28 | 1379.70 | 1004.37 |
| 2005 | 163.65 | -397.87 | -224.63 | 1503.41 | 1092.47 |
| 2006 | 144.85 | -372.89 | -220.39 | 1606.23 | 1194.02 |
| 2007 | 165.89 | -340.44 | -169.42 | 1764.95 | 1308.11 |
| 2008 | 196.18 | -260.35 | -64.47 | 1933.60 | 1393.31 |
| 2009 | 101.53 | -386.00 | -276.62 | 1900.26 | 1490.48 |
| 2010 | 190.58 | -305.45 | -112.91 | 2250.65 | 1733.27 |
| 2011 | 212.23 | -415.32 | -200.52 | 2272.72 | 1685.49 |
| 2012 | 185.41 | -321.95 | -134.11 | 2338.83 | 1779.36 |
| 2013 | 147.71 | -381.79 | -227.66 | 2420.98 | 1899.47 |
| 2014 | -43.96 | -483.08 | -521.23 | 2360.99 | 2021.21 |
| 2015 | -206.54 | -704.68 | -904.37 | 2215.77 | 2058.92 |
| 2016 | -260.20 | -518.11 | -775.03 | 2145.85 | 2036.76 |
| 2017 | -195.53 | -537.30 | -723.50 | 2179.58 | 2015.32 |
| 2018 | -181.32 | -472.17 | -647.25 | 2263.84 | 2054.83 |
| 2019 | -138.30 | -454.45 | -583.55 | 2394.51 | 2110.64 |
| 2020 | -1060.26 | -379.09 | -1442.13 | 2080.73 | 2766.88 |
| 2021 | -47.40 | -529.98 | -578.45 | 2530.36 | 2114.82 |
| 2022 | 58.20 | -600.98 | -532.55 | 2775.94 | 2169.30 |
| 2023 | -258.08 | -704.76 | -1003.76 | 2699.75 | 2439.37 |
| 2024 | -47.24 | -938.77 | -988.70 | 2944.06 | 2422.56 |
| 2025 | -62.52 | -931.33 | -990.78 | 3037.79 | 2504.17 |

CSVs gerados: `output/rtn/rtn_anual_ipca_mai26_2001_2025.csv` e `output/rtn/rtn_anual_corrente_2001_2025.csv`.

Fonte oficial (mesmo boletim mai/26): [Tesouro Transparente / CKAN](https://www.tesourotransparente.gov.br/ckan/dataset/resultado-do-tesouro-nacional).

## Aba 1.2-A — total de **cada** item (A6–A177) por período

Cada célula = soma das colunas mensais **na mesma linha** (ex.: período 1 na linha 6 = `SOMA(B6:BU6)`).

| # | Período | Meses | Colunas Excel | Ex. fórmula A6 |
|--:|---------|------:|---------------|----------------|
| 1 | jan/97–dez/02 | 72 | B:BU | `SOMA(B6:BU6)` |
| 2 | jan/03–mai/16 | 161 | BV:HZ | `SOMA(BV6:HZ6)` |
| 3 | jun/16–dez/18 | 31 | IA:JE | `SOMA(IA6:JE6)` |
| 4 | jan/19–dez/22 | 48 | JF:LA | `SOMA(JF6:LA6)` |
| 5 | jan/23–mai/26 | 41 | LB:MP | `SOMA(LB6:MP6)` |

```bat
.\python.exe scripts\somar_aba_1_2A_periodos.py "serie_historica_mai26 (2).xlsx" --out saida\totais_1_2A_por_item.xlsx
```

Arquivo: `output/rtn/totais_1_2A_por_item.xlsx` (172 itens × 5 períodos, R$ mi e R$ bi, IPCA Mai/2026).

Exemplo — linha 6 `1. RECEITA TOTAL` (R$ mi):

| jan/97–dez/02 | jan/03–mai/16 | jun/16–dez/18 | jan/19–dez/22 | jan/23–mai/26 |
|-------------:|-------------:|-------------:|-------------:|-------------:|
| 6.433.212,26 | 26.103.160,84 | 5.685.250,90 | 9.781.534,92 | 10.011.944,15 |
