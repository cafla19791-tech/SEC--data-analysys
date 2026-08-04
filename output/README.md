# Output — BIS CBPOL diário

## Arquivos

| Arquivo | Conteúdo |
|---------|----------|
| `cbpol_taxas_diarias_compostas.xlsx` | Histórico diário completo (todas as datas BIS) |
| `cbpol_taxas_diarias_compostas_desde_2000.xlsx` | Mesmo layout, a partir de **2000-01-01** (recomendado ContAgil) |
| `cbpol_taxas_acumuladas_periodos.xlsx` | Ranking por país da taxa acumulada em 6 períodos (diário, sem sáb/dom) |
| `cbpol_taxas_mensais_compostas.xlsx` | 1 aba/país com série mensal composta (1/12) |
| `cbpol_taxas_acumuladas_periodos_mensal.xlsx` | Ranking por períodos com taxas **mensais** |
| `pdf/*.pdf` | Versões PDF das planilhas (LibreOffice) |

Cada aba de país tem 5 colunas:

1. **Dia**
2. **Taxa (% a.d.)** — `(1 + taxa_aa/100)^(1/252) - 1` (ano com 252 dias úteis)
3. **Taxa acumulada (%)** — juros compostos desde o início da série
4. **Taxa acumulada mês (%)** — só no último dia do mês; compostos do mês
5. **Taxa acumulada ano (%)** — só no último dia do ano; compostos do ano

Abas auxiliares: `00_Legenda`, `01_Indice`.

## Regenerar

```bat
cd bis-mcp
bis_cli.bat excel-diario --csv ..\WS_CBPOL_csv_flat.csv --out ..\output\cbpol_taxas_diarias_compostas.xlsx
bis_cli.bat excel-diario --csv ..\WS_CBPOL_csv_flat.csv --from 2000-01-01 --out ..\output\cbpol_taxas_diarias_compostas_desde_2000.xlsx
bis_cli.bat excel-periodos --csv ..\WS_CBPOL_csv_flat.csv --out ..\output\cbpol_taxas_acumuladas_periodos.xlsx
bis_cli.bat excel-mensal --csv ..\WS_CBPOL_csv_flat.csv --out ..\output\cbpol_taxas_mensais_compostas.xlsx
bis_cli.bat excel-periodos --freq M --csv ..\WS_CBPOL_csv_flat.csv --out ..\output\cbpol_taxas_acumuladas_periodos_mensal.xlsx
```
