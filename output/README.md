# Output — BIS CBPOL diário

## Arquivos

| Arquivo | Conteúdo |
|---------|----------|
| `cbpol_taxas_diarias_compostas.xlsx` | Histórico diário completo (todas as datas BIS) |
| `cbpol_taxas_diarias_compostas_desde_2000.xlsx` | Mesmo layout, a partir de **2000-01-01** (recomendado ContAgil) |

Cada aba de país tem 3 colunas:

1. **Dia**
2. **Taxa (% a.d.)** — conversão ContAgil `(1 + taxa_aa/100)^(1/365) - 1`
3. **Taxa acumulada (%)** — juros compostos: `fator *= (1 + taxa_ad)`

Abas auxiliares: `00_Legenda`, `01_Indice`.

## Regenerar

```bat
cd bis-mcp
bis_cli.bat excel-diario --csv ..\WS_CBPOL_csv_flat.csv --out ..\output\cbpol_taxas_diarias_compostas.xlsx
bis_cli.bat excel-diario --csv ..\WS_CBPOL_csv_flat.csv --from 2000-01-01 --out ..\output\cbpol_taxas_diarias_compostas_desde_2000.xlsx
```
