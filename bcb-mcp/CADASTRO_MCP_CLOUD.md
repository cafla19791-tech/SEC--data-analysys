# Cadastro do bcb-mcp no Cursor Cloud

## O que este MCP faz

Acessa dados abertos do **Banco Central do Brasil**:

1. Series SGS (Selic, CDI, IPCA, cambio...)
2. PTAX USD/BRL (OLINDA)
3. Expectativas Focus (OLINDA)

## A) Registrar no cursor.com/agents

1. Abra [https://cursor.com/agents](https://cursor.com/agents)
2. Dropdown **MCP** -> Add custom MCP / stdio
3. Preencha:

| Campo | Valor |
|-------|--------|
| **Name** | `bcb-mcp` |
| **Transport** | `stdio` |
| **Command** | `bash` |
| **Args** | `./bcb-mcp/run_mcp.sh` |
| **Env** | `BCB_USER_AGENT` = `SEC-data-analysys-bcb-mcp/0.1 (seu@email.com)` |

4. Novo agente neste repo. Exemplos:

```text
Qual a Selic dos ultimos 30 dias?
Serie do IPCA de 2020 a 2025.
PTAX do dolar na ultima semana.
Expectativas Focus para IPCA.
```

## B) CLI no ContAgil WinPython (sem Cursor Desktop)

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bcb-mcp-f342/bcb-mcp/baixar_bcb_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_bcb.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bcb.ps1
```

Depois:

```bat
cd bcb-mcp
bcb_cli.bat catalog
bcb_cli.bat serie selic --last 10
bcb_cli.bat serie ipca --from 2020-01-01 --to 2025-12-31
bcb_cli.bat ptax --days 7
bcb_cli.bat expectativas Selic --top 10
```

## Arquitetura (resumo)

```
Cursor / CLI
    -> bcb-mcp tools
        -> api.bcb.gov.br/dados/serie/bcdata.sgs.{CODE}/dados  (SGS)
        -> olinda.bcb.gov.br/.../PTAX                          (cambio)
        -> olinda.bcb.gov.br/.../Expectativas                  (Focus)
```
