# Volumes e dispêndios — diesel, gasolina, GLP e QAV (2010–2026)

Extração mensal a partir da planilha ANP (Google Sheets):

- **Volumes**: Importação de derivados por produto (barris)
- **Dispêndios**: Dispêndio com importação por produto (US$ FOB)

Produtos: `ÓLEO DIESEL`, `GASOLINA A`, `GLP`, `QUEROSENE DE AVIAÇÃO`  
Período: **cada mês de 2010 a 2026**

## Como gerar

```bash
python3 extrair_volumes_importacao_derivados.py
```

## Saídas

| Arquivo | Conteúdo |
|---------|----------|
| `volumes_dispendios_diesel_gasolina_glp_qav_2010_2026.csv` | Ano, mês, produto, volume (b), dispêndio (US$ FOB) |
| `volumes_dispendios_diesel_gasolina_glp_qav_resumo_anual.csv` | Totais anuais por produto |
| `volumes_dispendios_diesel_gasolina_glp_qav_2010_2026.xlsx` | Longo, resumo, matrizes e pivôs Vol/Disp por produto |

Fonte: `data/anp/anp_importacoes_exportacoes_barris.xlsx`  
Link: https://docs.google.com/spreadsheets/d/1PsjXu8XIGahgJVRIO1MJ1IS3RjEnCe85
