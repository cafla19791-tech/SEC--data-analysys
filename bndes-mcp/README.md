# bndes-mcp

Consulta pública de **operações de financiamento do BNDES** por CNPJ/CPF (portal Transparência → gateway Solr) e exportação para Excel (ContAgil WinPython).

## API usada

```
GET https://gateway.apis.bndes.gov.br/operacoes/web/select
  ?q=documentoClienteIndex:(CNPJ)
  &tdoc=tipoDocumento:cnpj
  &rows=10000
  &wt=json
```

## ContAgil (PowerShell)

Cole **só** estes comandos no PowerShell do ContAgil:

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bndes-operacoes-mcp-41ca/bndes-mcp/baixar_bndes_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_bndes.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bndes.ps1
```

Embraer (CNPJ `07.689.002/0001-89`):

```powershell
.\bndes-mcp\bndes_cli.bat cnpj 07689002000189 --out embraer_bndes.xlsx
```

## Excel gerado

Abas:

| Aba | Conteúdo |
|-----|----------|
| Resumo | Totais e período |
| Por_Ano / Por_Produto / Por_Situacao / Por_Tope | Agregados |
| Operacoes | 1 linha por documento Solr |
| Subcreditos | Expansão do XML `subcreditos` |

**Nota:** várias operações EXIM pós-embarque **não publicam** `valorContratacao` no documento agregado; o CNPJ pode aparecer mascarado (`XXXXXXXXX00189`) em parte das linhas.

## CLI local

```bash
PYTHONPATH=src python -m bndes_mcp.cli cnpj 07689002000189 --out output/embraer_bndes.xlsx
PYTHONPATH=src python -m bndes_mcp.cli resumo 07689002000189
PYTHONPATH=src python -m bndes_mcp.cli json-para-excel arquivo.json --out saida.xlsx
```
