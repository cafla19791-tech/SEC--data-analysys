# Cadastro exato do `nyse-mcp` no Cursor Cloud (browser)

Use isto se a empresa **bloqueia Cursor Desktop**. VS Code local **não** registra o MCP do Cursor.

## A) Registrar MCP no Cloud Agent (recomendado)

1. Abra **[https://cursor.com/agents](https://cursor.com/agents)** e faça login.
2. Abra (ou crie) um agente no repositório `SEC--data-analysys`.
3. No topo / painel do agente, abra o dropdown **MCP**.
4. Clique em **Add custom MCP** / **Add server** (stdio).
5. Preencha **exatamente**:

| Campo | Valor |
|-------|--------|
| **Name** | `nyse-mcp` |
| **Transport** | `stdio` |
| **Command** | `bash` |
| **Args** | `./nyse-mcp/run_mcp.sh` |
| **Env** (opcional) | `MARKET_DATA_PROVIDER` = `yahoo` |

6. Salve. Confirme que o environment do Cloud Agent já instala deps
   (há `.cursor/environment.json` no repo; se o setup falhar, peça ao agente:
   `cd nyse-mcp && python3 -m venv .venv && pip install -e .`).
7. Inicie um **novo** agente (ou recarregue) e pergunte, por exemplo:

```text
Qual a cotação da JPM na NYSE?
Mostre fundamentals da XOM.
Busque o ticker da Coca-Cola.
```

### Personal vs Team

- **Personal MCP:** dropdown MCP em `cursor.com/agents`.
- **Team / shared:** admin em **Dashboard → Integrations & MCP**.

Docs: [Cloud Agent capabilities → MCP tools](https://cursor.com/docs/cloud-agent/capabilities).

## B) Sem registrar MCP (CLI no Cloud Agent)

No chat do Cloud Agent, peça:

```text
Rode no VM:
  cd nyse-mcp && python3 -m venv .venv && source .venv/bin/activate && pip install -e .
  nyse-mcp-cli quote JPM
  nyse-mcp-cli fundamentals XOM
e me explique o JSON.
```

## C) CLI no ContAgil WinPython (sem clonar o repo)

**Não use** `cd caminho\para\...` — isso era só um exemplo genérico.

Cole **exatamente** no CMD/PowerShell do WinPython:

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$b="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/nyse-mcp-winpython-cli-f342"
Invoke-WebRequest "$b/nyse-mcp/baixar_nyse_mcp_winpython.ps1" -OutFile baixar_nyse_mcp_winpython.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_nyse_mcp_winpython.ps1 -Symbol JPM
```

Isso cria `winpython\nyse-mcp\` e testa a cotação. Depois:

```bat
cd nyse-mcp
nyse_mcp_cli.bat quote JPM
nyse_mcp_cli.bat history XOM --period 1y
nyse_mcp_cli.bat search "Coca Cola"
nyse_mcp_cli.bat status
```

> Se `pip` ou `finance.yahoo.com` forem bloqueados na rede da RFB, use só o caminho **A** (Cloud Agent), onde a saída de rede costuma ser mais ampla.

## O que NÃO fazer

- Não esperar que o VS Code “ative” o MCP do Cursor.
- Não depender só de `.cursor/mcp.json` nos Cloud Agents (não é lido sozinho).
- Não usar transport SSE no Cloud Agent (não suportado).
