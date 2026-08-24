# BIS Data Portal — Excel por tema, aba por país

Planilhas geradas a partir de https://data.bis.org/bulkdownload

Há **um arquivo Excel por tópico** do portal. Em cada arquivo:

| Aba | Conteúdo |
|-----|----------|
| `Capa` | tema, dataflows, recorte aplicado e citação do BIS |
| `Indice` | países e agregados, com último período e aba correspondente |
| `Comparativo` | último valor disponível por país |
| `BR`, `US`, `5A`… | séries daquele país/jurisdição no tema |

Códigos numéricos (`5A`, `5C`, `5R`…) são agregados (todas as economias, zona do euro, avançadas etc.).

## Como gerar

```bash
python3 scripts/bis_bulk_excel.py
python3 scripts/bis_bulk_excel.py --topics CBPOL,CPI,CREDIT_GAP
```

Os ZIP oficiais ficam em `data/raw/bis/` (não versionados). Os `.xlsx` saem nesta pasta.

Conjuntos muito grandes (LBS, CBS, títulos, câmbio diário) entram com recorte de frequência (A/Q/M quando existir) e, se preciso, janela recente — o detalhe está na aba `Capa` de cada arquivo.
