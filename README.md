# SEC--data-analysys

Análises e planilhas de dados fiscais / crédito / subsídios.

## Orçamento de Subsídios da União (OSU) 2025

Fonte oficial (MPO):

https://www.gov.br/planejamento/pt-br/assuntos/avaliacao-de-politicas-publicas/arquivos/orcamento-de-subsidios-da-uniao/osu_2025-anexos-publicacao.xlsx/@@download/file

| Arquivo | Descrição |
|---|---|
| `OSU_2025_Anexos.xlsx` | Planilha limpa com índice, Tabelas 1–8 e resumos |
| `data/osu_2025/OSU_2025_anexos_fonte.xlsx` | Arquivo oficial baixado do gov.br |
| `data/osu_2025/download_osu_2025.py` | Script para baixar a fonte |
| `scripts/build_osu_2025_anexos.py` | Regenera `OSU_2025_Anexos.xlsx` a partir da fonte |

### Regenerar

```bash
python data/osu_2025/download_osu_2025.py
python scripts/build_osu_2025_anexos.py
```

### Destaques 2024 (valores nominais)

- **Total de subsídios:** R$ 678,4 bi
- **Tributários:** 83,1% · **Financeiros:** 9,6% · **Creditícios:** 7,3%
- Maiores itens: Simples Nacional, Agricultura e Agroindústria, IRPF isentos

## Renúncia fiscal SUDAM/SUDENE

Arquivo: `RENUNCIA FISCAL SUDAM-SUDENE (1).xlsx` (período 2015–2023, com atualização IPCA).
