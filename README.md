# SEC--data-analysys

Includes a skeleton MCP server for US equity market data that works **without Cursor Desktop** (Cloud Agents / corporate locked-down machines):

- [`nyse-mcp/`](./nyse-mcp/) — MCP tools + CLI (`nyse-mcp-cli`)
- [`nyse-mcp/CLOUD_SETUP.md`](./nyse-mcp/CLOUD_SETUP.md) — setup via [cursor.com/agents](https://cursor.com/agents)
- [`.cursor/mcp.json`](./.cursor/mcp.json) — reference stdio config (IDE); Cloud Agents need dashboard registration
- [`.cursor/environment.json`](./.cursor/environment.json) — installs deps in Cloud Agent VMs
