# Output — BIS WS_LONG_CPI

| Arquivo | Conteúdo |
|---------|----------|
| `cpi_mensal_por_pais.xlsx` | 1 aba/país: Mês \| Índice (2010=100) \| YoY (%) \| Inflação acumulada (%) |
| `cpi_inflacao_acumulada_periodos.xlsx` | Rankings de inflação acumulada em 6 períodos |
| `pdf/*.pdf` | Conversões LibreOffice |

Fonte: `https://data.bis.org/static/bulk/WS_LONG_CPI_csv_col.zip` (63 países).

Regenerar:

```bash
export PYTHONPATH=bis-cpi/src
python3 -m bis_cpi.cli download --dir .
python3 -m bis_cpi.cli excel-mensal --out output/cpi/cpi_mensal_por_pais.xlsx
python3 -m bis_cpi.cli excel-periodos --out output/cpi/cpi_inflacao_acumulada_periodos.xlsx
python3 -m bis_cpi.cli para-pdf output/cpi/cpi_mensal_por_pais.xlsx --outdir output/cpi/pdf
python3 -m bis_cpi.cli para-pdf output/cpi/cpi_inflacao_acumulada_periodos.xlsx --outdir output/cpi/pdf
```
