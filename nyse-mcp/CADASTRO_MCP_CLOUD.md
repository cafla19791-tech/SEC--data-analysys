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

## C) CLI no Windows (VS Code terminal / ContAgil WinPython)

1. Clone ou baixe o repo (branch com `nyse-mcp/`).
2. No PowerShell:

```powershell
cd caminho\para\SEC--data-analysys\nyse-mcp
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_e_rodar_cli.ps1 -Symbol JPM
```

3. Depois:

```bat
nyse_mcp_cli.bat quote JPM
nyse_mcp_cli.bat history XOM --period 1y
nyse_mcp_cli.bat search "Coca Cola"
nyse_mcp_cli.bat status
```

Com Python do ContAgil:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_e_rodar_cli.ps1 `
  -Python "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\python.exe" `
  -Symbol PBR
```

> Se `pip` ou `finance.yahoo.com` forem bloqueados na rede da RFB, use só o caminho **A** (Cloud Agent), onde a saída de rede costuma ser mais ampla.

## O que NÃO fazer

- Não esperar que o VS Code “ative” o MCP do Cursor.
- Não depender só de `.cursor/mcp.json` nos Cloud Agents (não é lido sozinho).
- Não usar transport SSE no Cloud Agent (não suportado).
