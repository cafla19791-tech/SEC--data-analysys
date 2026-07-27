# bcb-mcp

MCP + CLI para **dados abertos do Banco Central do Brasil** (SGS e OLINDA).

Fontes publicas (sem API key):

- [SGS](https://www3.bcb.gov.br/sgspub/) — series temporais (`api.bcb.gov.br`)
- [OLINDA](https://olinda.bcb.gov.br/) — PTAX e Expectativas Focus
- Portal: [dadosabertos.bcb.gov.br](https://dadosabertos.bcb.gov.br/)

## Tools (MCP)

| Tool | Uso |
|------|-----|
| `list_known_series` | Aliases locais (selic, cdi, ipca, dolar...) |
| `get_sgs_series` | Serie SGS por alias ou codigo |
| `get_ptax` | Cotacao USD/BRL (OLINDA) |
| `get_expectativas` | Expectativas Focus anuais |

## Cadastro no Cursor Cloud

Veja **[CADASTRO_MCP_CLOUD.md](./CADASTRO_MCP_CLOUD.md)**.

## CLI (ContAgil / sem MCP)

```bash
cd bcb-mcp
export PYTHONPATH=src
python -m bcb_mcp.cli catalog
python -m bcb_mcp.cli serie selic --last 5
python -m bcb_mcp.cli serie ipca --from 2020-01-01 --to 2025-12-31
python -m bcb_mcp.cli serie 11 --from 01/01/2024 --to 31/12/2024
python -m bcb_mcp.cli ptax --days 10
python -m bcb_mcp.cli expectativas IPCA --top 10
```

WinPython ContAgil: use `baixar_bcb_winpython.ps1`.

## Aliases uteis

| Alias | Codigo SGS | Descricao |
|-------|------------|-----------|
| `selic` | 11 | Selic diaria |
| `selic_meta` | 432 | Meta Selic |
| `cdi` | 12 | CDI diario |
| `ipca` | 433 | IPCA mensal |
| `dolar` | 1 | PTAX dolar venda (SGS) |
| `igpm` | 189 | IGP-M mensal |

Qualquer codigo numerico do SGS tambem funciona.

## Limites

- Series diarias: janela maxima ~10 anos por request (o cliente fatia automaticamente).
- Rate limit nao documentado; o cliente faz throttle leve.
- Preferir `--last` ou janelas curtas quando possivel.
