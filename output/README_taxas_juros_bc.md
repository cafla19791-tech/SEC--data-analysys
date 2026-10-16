# Taxas de juros — Meta Selic (BCB) e bancos centrais (BIS)

Arquivo principal: `taxas_juros_bc_paises_1996.xlsx`

## Conteúdo

- **Indice** — lista das abas
- **Meta_Selic_BCB** — Meta Selic e períodos de vigência (Copom), desde 01/07/1996
- **Uma aba por país** (49) — taxas diárias e mensais desde 01/01/1996
- **Fontes** — URLs das fontes

## Como regenerar

```bash
# Requer data/bis/WS_CBPOL_csv_flat.csv (baixado automaticamente se ausente)
python3 scripts/gerar_planilha_taxas_bc.py --output output/taxas_juros_bc_paises_1996.xlsx
```

## Fontes

- BCB: https://www.bcb.gov.br/controleinflacao/historicotaxasjuros
- BIS WS_CBPOL: https://data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip

Nota: a planilha original em formato colunar no Drive estoura o limite de colunas do Excel
(corta em maio/1988). Este gerador usa o CSV flat completo do BIS.
