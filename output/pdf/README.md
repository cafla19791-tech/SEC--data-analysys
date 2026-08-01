# PDFs gerados a partir das planilhas CBPOL

Convertidos com LibreOffice (`soffice --convert-to pdf`).

| PDF | Origem |
|-----|--------|
| `cbpol_taxas_acumuladas_periodos.pdf` | ranking diário (6 períodos) |
| `cbpol_taxas_acumuladas_periodos_mensal.pdf` | ranking mensal (6 períodos) |
| `cbpol_taxas_mensais_compostas.pdf` | séries mensais (1 aba/país) |
| `cbpol_taxas_diarias_compostas_desde_2000.pdf` | séries diárias desde 2000 |
| `cbpol_taxas_diarias_compostas.pdf` | séries diárias completas (~31 MB) |
| `*_indice.pdf` | só abas Legenda + Índice |

## Regenerar

```bash
# Linux (com LibreOffice)
bis-cli para-pdf output/cbpol_taxas_acumuladas_periodos.xlsx --outdir output/pdf
```

```bat
REM ContAgil / Windows — requer LibreOffice instalado
cd bis-mcp
bis_cli.bat para-pdf ..\cbpol_taxas_acumuladas_periodos.xlsx --outdir ..\pdf
```
