# sec-edgar-mcp

MCP + CLI para **relatorios / filings da SEC (EDGAR)** e fatos financeiros XBRL.

Fonte: API publica [data.sec.gov](https://data.sec.gov/) (sem API key).  
Obrigatorio: header `User-Agent` com nome do app + e-mail (`SEC_USER_AGENT`).

## Tools (MCP)

| Tool | Uso |
|------|-----|
| `lookup_ticker` | Ticker -> CIK |
| `get_company_profile` | Nome, SIC, exchanges |
| `list_filings` | Lista 10-K / 10-Q / 8-K... com URLs |
| `get_company_facts` | Receita, lucro, ativos (XBRL) |
| `get_concept` | Serie de um conceito us-gaap / ifrs-full |
| `get_concept_range` | Serie anual mesclando us-gaap + IFRS |
| `get_total_debt` | Borrowings + LeaseLiabilities (preenche via 20-F) |
| `extract_filing_concepts` | Extrai tags do iXBRL do ultimo 20-F/10-K |

## Cadastro no Cursor Cloud

Veja **[CADASTRO_MCP_CLOUD.md](./CADASTRO_MCP_CLOUD.md)**.

Resumo stdio:

| Campo | Valor |
|-------|--------|
| Name | `sec-edgar-mcp` |
| Command | `bash` |
| Args | `./sec-edgar-mcp/run_mcp.sh` |
| Env | `SEC_USER_AGENT=SEC-data-analysys/0.1 (seu@email.com)` |

## CLI (ContAgil / sem MCP)

```bash
cd sec-edgar-mcp
export PYTHONPATH=src
export SEC_USER_AGENT='SEC-data-analysys/0.1 (seu@email.com)'
python -m sec_edgar_mcp.cli lookup AAPL
python -m sec_edgar_mcp.cli filings AAPL --form 10-K --limit 5
python -m sec_edgar_mcp.cli facts AAPL
python -m sec_edgar_mcp.cli concept AAPL NetIncomeLoss
python -m sec_edgar_mcp.cli series PBR NetIncomeLoss --from 2008 --to 2025
python -m sec_edgar_mcp.cli debt PBR --from 2016 --to 2025
python -m sec_edgar_mcp.cli filing-xbrl PBR --concepts Borrowings,LeaseLiabilities --form 20-F
```

WinPython ContAgil: use `baixar_sec_edgar_winpython.ps1` (mesmo padrao do nyse-mcp v3).

### Divuda total (PBR)

`debt` soma `ifrs-full:Borrowings` + `ifrs-full:LeaseLiabilities`.  
Quando o `companyfacts` ainda nao tem o FY (ex.: 2025), baixa o iXBRL do 20-F mais recente e preenche.  
A "divida bruta" dos releases da Petrobras costuma ser so `Borrowings` (Finance Debt).

## Limites

- Rate limit SEC: ~10 req/s (o cliente faz throttle).
- Sem User-Agent valido -> HTTP 403.
- `companyfacts` traz numeros XBRL; o HTML/PDF completo do 10-K esta em `documentUrl` / `indexUrl`.
