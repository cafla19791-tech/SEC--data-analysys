# tesouro-mcp

MCP + CLI para **estatisticas fiscais do Tesouro Nacional** (Resultado do Tesouro Nacional / RTN).

Fontes publicas (sem API key):

- [ARIA Series Temporais](https://apiapex.tesouro.gov.br/aria/v1/series-temporais/docs) — series mensais RTN
- [Grandes Numeros](https://grandesnumeros.tesouro.gov.br) — indicadores da capa
- [Tesouro Transparente / CKAN](https://www.tesourotransparente.gov.br/ckan/) — planilhas e metadados

## Tools (MCP)

| Tool | Uso |
|------|-----|
| `list_temas` | Temas 10 / 13 / 20 |
| `list_known_aliases` | Aliases (`resultado_primario`, `receita_total`...) |
| `list_series` / `search_series` | Catalogo ARIA |
| `get_serie` | Serie mensal por alias/codigo |
| `get_resultado_fiscal` | Consulta RTN por tema |
| `get_grandes_numeros` | Headlines (primario, DPF...) |
| `ckan_package_search` / `ckan_package_show` | Datasets abertos |

## CLI (ContAgil)

```bash
cd tesouro-mcp
export PYTHONPATH=src
python -m tesouro_mcp.cli aliases
python -m tesouro_mcp.cli serie resultado_primario --from 2024-01 --to 2025-12
python -m tesouro_mcp.cli serie receita_total --from 01/2024 --to 12/2024
python -m tesouro_mcp.cli headline
python -m tesouro_mcp.cli ckan-show resultado-do-tesouro-nacional
```

WinPython: `baixar_tesouro_winpython.ps1`.

## Temas RTN

| Codigo | Conteudo |
|--------|----------|
| 10 | Resultado fiscal do Governo Central (receitas, despesas, primario) |
| 13 | Investimento do Governo Federal |
| 20 | Custeio administrativo |

Unidade tipica das series: **R$ milhoes**.

## Cadastro Cloud

Veja [CADASTRO_MCP_CLOUD.md](./CADASTRO_MCP_CLOUD.md).
