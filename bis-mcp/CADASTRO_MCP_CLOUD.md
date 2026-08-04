# Cadastro do bis-mcp no Cursor Cloud

## O que este MCP faz

Acessa **taxas de política monetária do BIS** (`WS_CBPOL`):

1. Séries por país via SDMX (BR/SELIC, US, Eurozona...)
2. Snapshot comparativo da última taxa
3. Download / leitura do `WS_CBPOL_csv_flat.csv` (ContAgil)

## A) Registrar no cursor.com/agents

1. Abra [https://cursor.com/agents](https://cursor.com/agents)
2. Dropdown **MCP** -> Add custom MCP / stdio
3. Preencha:

| Campo | Valor |
|-------|--------|
| **Name** | `bis-mcp` |
| **Transport** | `stdio` |
| **Command** | `bash` |
| **Args** | `./bis-mcp/run_mcp.sh` |
| **Env** | `BIS_USER_AGENT` = `SEC-data-analysys-bis-mcp/0.1 (seu@email.com)` |

4. Novo agente neste repo. Exemplos:

```text
Qual a meta Selic mensal BIS dos últimos 12 meses?
Compare BR, US e Eurozona (última taxa).
Baixe WS_CBPOL_csv_flat.csv para /tmp.
```

## B) CLI no ContAgil WinPython (sem Cursor Desktop)

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bis-cbpol-mcp-41ca/bis-mcp/baixar_bis_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_bis.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bis.ps1 -DownloadCsv
```

Depois:

```bat
cd bis-mcp
bis_cli.bat catalog
bis_cli.bat serie brasil --last 12
bis_cli.bat compare BR,US,XM
bis_cli.bat serie selic --local --last 24
```

## Arquitetura (resumo)

```
Cursor / CLI
    -> bis-mcp tools
        -> stats.bis.org/api/v1/data/WS_CBPOL/{FREQ}.{AREAS}   (SDMX CSV)
        -> data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip      (bulk)
        -> local WS_CBPOL_csv_flat.csv                         (ContAgil)
```
