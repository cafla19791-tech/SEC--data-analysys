# Extrações ANP — diesel, gasolina, GLP e QAV (2010–2026)

## 1) Importação (volumes + dispêndios)

Fonte: [planilha importações/exportações](https://docs.google.com/spreadsheets/d/1PsjXu8XIGahgJVRIO1MJ1IS3RjEnCe85)

```bash
python3 extrair_volumes_importacao_derivados.py
```

| Arquivo | Conteúdo |
|---------|----------|
| `volumes_dispendios_diesel_gasolina_glp_qav_2010_2026.csv` | Volume (b) + dispêndio (US$ FOB) mensal |
| `volumes_dispendios_diesel_gasolina_glp_qav_resumo_anual.csv` | Totais anuais |
| `volumes_dispendios_diesel_gasolina_glp_qav_2010_2026.xlsx` | Longo, matrizes e pivôs |

## 2) Produção nacional nas refinarias (volumes)

Fonte: [planilha produção nacional](https://docs.google.com/spreadsheets/d/1X_DHZxEJe4iP02gwGy6y24eHiMTy1bCd)

Escopo: **Brasil — todas as refinárias** (visão `REFINARIA = (Tudo)`).  
Unidade: **barris**.

```bash
python3 extrair_volumes_producao_derivados.py
```

| Arquivo | Conteúdo |
|---------|----------|
| `producao_diesel_gasolina_glp_qav_2010_2026.csv` | Produção mensal (barris) |
| `producao_diesel_gasolina_glp_qav_resumo_anual.csv` | Totais anuais |
| `producao_diesel_gasolina_glp_qav_2010_2026.xlsx` | Longo, matriz e pivôs por produto |

Produtos em ambas: `ÓLEO DIESEL`, `GASOLINA A`, `GLP`, `QUEROSENE DE AVIAÇÃO`.
