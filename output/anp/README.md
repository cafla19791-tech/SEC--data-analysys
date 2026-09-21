# Volumes importados de derivados de petróleo (ANP) — barris

Fonte: planilha ANP compartilhada (Google Sheets), seção
**Importação de derivados de petróleo por produto - 2000-2026 (b)**.

Arquivo local: `data/anp/anp_importacoes_exportacoes_barris.xlsx`  
Link: https://docs.google.com/spreadsheets/d/1PsjXu8XIGahgJVRIO1MJ1IS3RjEnCe85

Unidade: **barris (b)**. Atualização da planilha: 28/08/2026.

## Como gerar

```bash
python3 extrair_volumes_importacao_derivados.py
```

## Saídas

| Arquivo | Conteúdo |
|---------|----------|
| `volumes_importacao_derivados_barris_longo.csv` | Ano, mês, produto, volume (barris) |
| `volumes_importacao_derivados_barris_resumo_anual.csv` | Totais anuais por produto |
| `volumes_importacao_derivados_barris.xlsx` | Pivô total + uma aba por derivado |

Produtos: ASFALTO, COQUE, GASOLINA A, GASOLINA DE AVIAÇÃO, GLP, LUBRIFICANTE,
NAFTA, OUTROS NÃO ENERGÉTICOS, PARAFINA, QUEROSENE DE AVIAÇÃO, QUEROSENE ILUMINANTE,
SOLVENTE, ÓLEO COMBUSTÍVEL, ÓLEO DIESEL.
