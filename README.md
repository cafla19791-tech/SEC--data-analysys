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
```

Docs: [`bis-mcp/README.md`](bis-mcp/README.md) · cadastro Cloud: [`bis-mcp/CADASTRO_MCP_CLOUD.md`](bis-mcp/CADASTRO_MCP_CLOUD.md)
