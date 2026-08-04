# Usar o nyse-mcp **sem Cursor Desktop**

Se a empresa bloqueia software não homologado, você **não precisa** de Cursor local nem de **Settings → Tools & MCP**.
**VS Code local também não registra o MCP do Cursor.**

O fluxo correto é **Cursor Cloud Agent** (browser) + registro do MCP no dashboard web.

## Caminho recomendado (Cloud Agent + stdio)

Passo a passo completo (tabela de campos): **[CADASTRO_MCP_CLOUD.md](./CADASTRO_MCP_CLOUD.md)**.

Resumo:

1. Abra [cursor.com/agents](https://cursor.com/agents)
2. No dropdown **MCP**, adicione um servidor custom **stdio**:
   - **Name:** `nyse-mcp`
   - **Command:** `bash`
   - **Args:** `./nyse-mcp/run_mcp.sh`
   - **Env (opcional):** `MARKET_DATA_PROVIDER=yahoo`
3. Garanta que o environment do Cloud Agent instala deps (já há [`.cursor/environment.json`](../.cursor/environment.json)).
4. Inicie um novo agente neste repo e pergunte, por exemplo:  
   `Qual a cotação da JPM na NYSE?`

Personal MCP: dropdown em `cursor.com/agents`.  
Team/shared MCP: admin em **Dashboard → Integrations & MCP**.

Documentação oficial: [Cloud Agent capabilities → MCP tools](https://cursor.com/docs/cloud-agent/capabilities).

## Alternativa imediata (sem registrar MCP)

### Cloud Agent / Linux

```bash
cd nyse-mcp
source .venv/bin/activate   # ou: python3 -m venv .venv && pip install -e .
nyse-mcp-cli quote JPM
nyse-mcp-cli history XOM --period 1y
nyse-mcp-cli fundamentals AAPL
nyse-mcp-cli search "Coca Cola"
nyse-mcp-cli status
```

Ou peça ao Cloud Agent: *“rode `nyse-mcp-cli quote JPM` e me explique o resultado”*.

### Windows (VS Code terminal / ContAgil WinPython)

```powershell
cd caminho\para\SEC--data-analysys\nyse-mcp
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_e_rodar_cli.ps1 -Symbol JPM
.\nyse_mcp_cli.bat quote XOM
```

Scripts: `setup_e_rodar_cli.ps1`, `nyse_mcp_cli.bat`.

## HTTP remoto (quando houver host homologado)

Cursor Cloud recomenda **HTTP** para MCP remoto (SSE/`mcp-remote` não são suportados).

No ambiente onde o servidor puder ficar exposto:

```bash
export MCP_TRANSPORT=streamable-http
export MCP_HOST=0.0.0.0
export MCP_PORT=8000
bash ./nyse-mcp/run_mcp.sh
# endpoint: http://<host>:8000/mcp
```

No dashboard MCP do Cloud Agent, cadastre o URL HTTP (não o stdio).

> Em redes corporativas, hospedar HTTP costuma exigir aprovação de infra/segurança. A CLI + stdio no Cloud Agent costuma ser o caminho mais simples.

## O que NÃO fazer

- Não depender de Cursor Desktop / Tools & MCP no notebook bloqueado.
- Não esperar que só commitar `.cursor/mcp.json` ative o MCP nos Cloud Agents.
- Não usar transport SSE no Cloud Agent (não suportado).
