# SEC--data-analysys

Ferramentas de análise de dados para uso no **ContAgil** (WinPython) e Cursor Cloud Agents.

## bis-mcp — taxas de política do BIS (`WS_CBPOL`)

Pacote para o CSV flat do BIS:

`C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\WS_CBPOL_csv_flat.csv`

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bis-cbpol-mcp-41ca/bis-mcp/baixar_bis_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_bis.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bis.ps1 -DownloadCsv
```

```bat
cd bis-mcp
bis_cli.bat serie brasil --last 12
bis_cli.bat compare BR,US,XM
bis_cli.bat serie selic --local --last 24
bis_cli.bat excel-diario --from 2000-01-01 --out ..\output\cbpol_taxas_diarias_compostas_desde_2000.xlsx
```

Excel gerado (1 aba/país: Dia | Taxa % a.d. | Taxa acumulada compostos):

- [`output/cbpol_taxas_diarias_compostas_desde_2000.xlsx`](output/cbpol_taxas_diarias_compostas_desde_2000.xlsx) (recomendado)
- [`output/cbpol_taxas_diarias_compostas.xlsx`](output/cbpol_taxas_diarias_compostas.xlsx) (histórico completo)
- [`output/cbpol_taxas_acumuladas_periodos.xlsx`](output/cbpol_taxas_acumuladas_periodos.xlsx) (ranking por períodos, diário)
- [`output/cbpol_taxas_mensais_compostas.xlsx`](output/cbpol_taxas_mensais_compostas.xlsx) (séries mensais)
- [`output/cbpol_taxas_acumuladas_periodos_mensal.xlsx`](output/cbpol_taxas_acumuladas_periodos_mensal.xlsx) (ranking por períodos, mensal)

Docs: [`bis-mcp/README.md`](bis-mcp/README.md) · cadastro Cloud: [`bis-mcp/CADASTRO_MCP_CLOUD.md`](bis-mcp/CADASTRO_MCP_CLOUD.md)
