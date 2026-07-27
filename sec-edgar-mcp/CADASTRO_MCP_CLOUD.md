# Cadastro do sec-edgar-mcp no Cursor Cloud

## O que este MCP faz

Acessa a SEC (EDGAR) para:

1. Resolver ticker -> CIK  
2. Listar formularios (10-K, 10-Q, 8-K, 20-F...) com links  
3. Extrair fatos financeiros XBRL (receita, lucro liquido, ativos...)

## A) Registrar no cursor.com/agents

1. Abra [https://cursor.com/agents](https://cursor.com/agents)
2. Dropdown **MCP** -> Add custom MCP / stdio
3. Preencha:

| Campo | Valor |
|-------|--------|
| **Name** | `sec-edgar-mcp` |
| **Transport** | `stdio` |
| **Command** | `bash` |
| **Args** | `./sec-edgar-mcp/run_mcp.sh` |
| **Env** | `SEC_USER_AGENT` = `SEC-data-analysys/0.1 (cafla19791@gmail.com)` |

4. Novo agente neste repo. Exemplos de pergunta:

```text
Liste os ultimos 5 formularios 10-K da Apple na SEC.
Qual o NetIncomeLoss da Petrobras (PBR) nos ultimos anos (XBRL)?
Qual o CIK da Coca-Cola (KO)?
```

## B) CLI no ContAgil WinPython (sem Cursor Desktop)

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/sec-edgar-mcp-f342/sec-edgar-mcp/baixar_sec_edgar_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_sec_edgar.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_sec_edgar.ps1
```

Depois:

```bat
cd sec-edgar-mcp
sec_edgar_cli.bat lookup AAPL
sec_edgar_cli.bat filings AAPL --form 10-K --limit 5
sec_edgar_cli.bat facts KO
sec_edgar_cli.bat concept PBR NetIncomeLoss
sec_edgar_cli.bat concept PBR NetIncomeLoss --taxonomy ifrs-full --annual
```

## Arquitetura (resumo)

```
Cursor / CLI
    -> sec-edgar-mcp tools
        -> data.sec.gov/submissions/CIK....json   (historico de filings)
        -> data.sec.gov/api/xbrl/companyfacts/... (XBRL)
        -> www.sec.gov/files/company_tickers.json (ticker->CIK)
        -> www.sec.gov/Archives/edgar/data/...    (HTML/PDF do relatorio)
```

Nao e feed pago da NYSE; e o repositorio oficial de divulgacoes da SEC.
