# tesouro-mcp

MCP + CLI para **estatisticas fiscais do Tesouro Nacional** (Resultado do Tesouro Nacional / RTN) e coletor anual de divida/resultado.

Fontes publicas (sem API key):

- [ARIA Series Temporais](https://apiapex.tesouro.gov.br/aria/v1/series-temporais/docs) — series mensais RTN
- [Grandes Numeros](https://grandesnumeros.tesouro.gov.br) — indicadores da capa
- [Tesouro Transparente / CKAN](https://www.tesourotransparente.gov.br/ckan/) — planilhas e metadados
- [BCB SGS](https://dadosabertos.bcb.gov.br/) — DBGG (13761) e desembolsos BNDES (7415)

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
| `collect_annual_table` | Tabela anual DBGG/RTN/DPF/BNDES (+ DGT/fundos via CSV) |

## CLI (ContAgil)

```bash
cd tesouro-mcp
export PYTHONPATH=src
python -m tesouro_mcp.cli aliases
python -m tesouro_mcp.cli serie resultado_primario --from 2024-01 --to 2025-12
python -m tesouro_mcp.cli serie receita_total --from 01/2024 --to 12/2024
python -m tesouro_mcp.cli headline
python -m tesouro_mcp.cli ckan-show resultado-do-tesouro-nacional
python -m tesouro_mcp.cli coletar-anual --from 2001 --to 2025 --out tabela_anual.csv \
  --dgt data/templates/dgt_renuncias_anual.csv \
  --fundos data/templates/fundos_constitucionais_anual.csv

# Serie historica RTN (XLSX do boletim, ex. mai/26) — correntes ou IPCA:
python -m tesouro_mcp.cli rtn-xlsx ../serie_historica_mai26.xlsx --constantes-ipca --out rtn_ipca.csv
python -m tesouro_mcp.cli coletar-anual --from 2001 --to 2025 \
  --rtn-xlsx ../serie_historica_mai26.xlsx --constantes-ipca --out tabela_anual_ipca.csv
```

WinPython: `baixar_tesouro_winpython.ps1`. No PowerShell use `.\tesouro_cli.bat` (com `.\`) de dentro da pasta `tesouro-mcp`.

### Coletor anual

Preenche automaticamente (R$ bilhoes, valores correntes):

- DBGG 1/jan e 31/dez (BCB 13761)
- Resultado primario / juros nominais / resultado nominal (RTN)
- Emissoes e resgates da DPF (XLSX Tesouro Transparente)
- Desembolsos BNDES (BCB 7415)

Colunas **sem API aberta** — cole nos templates e passe `--dgt` / `--fundos`:

- Renuncias DGT: `data/templates/dgt_renuncias_anual.csv`
- FNO / FNE / FCO: `data/templates/fundos_constitucionais_anual.csv`

Instrucoes: [`data/templates/INSTRUCOES_COLA_DGT_FUNDOS.md`](./data/templates/INSTRUCOES_COLA_DGT_FUNDOS.md).

## Temas RTN

| Codigo | Conteudo |
|--------|----------|
| 10 | Resultado fiscal do Governo Central (receitas, despesas, primario) |
| 13 | Investimento do Governo Federal |
| 20 | Custeio administrativo |

Unidade tipica das series mensais: **R$ milhoes**. Tabela anual do coletor: **R$ bilhoes**.

## Cadastro Cloud

Veja [CADASTRO_MCP_CLOUD.md](./CADASTRO_MCP_CLOUD.md).
