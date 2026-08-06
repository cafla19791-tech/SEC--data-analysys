# SEC--data-analysys

Análises e planilhas de dados fiscais / crédito / subsídios.

## Orçamento de Subsídios da União (OSU) 2025

Fonte oficial (MPO):

https://www.gov.br/planejamento/pt-br/assuntos/avaliacao-de-politicas-publicas/arquivos/orcamento-de-subsidios-da-uniao/osu_2025-anexos-publicacao.xlsx/@@download/file

| Arquivo | Descrição |
|---|---|
| `OSU_2025_Anexos.xlsx` | Planilha limpa com índice, Tabelas 1–8 e resumos |
| `OSU_2025_Discriminativo_IPCA.xlsx` | Discriminativo com valores atualizados pelo IPCA até jun/2026 |
| `data/osu_2025/OSU_2025_anexos_fonte.xlsx` | Arquivo oficial baixado do gov.br |
| `data/osu_2025/download_osu_2025.py` | Script para baixar a fonte |
| `scripts/build_osu_2025_anexos.py` | Regenera `OSU_2025_Anexos.xlsx` a partir da fonte |
| `scripts/build_osu_discriminativo_ipca.py` | Regenera o discriminativo com IPCA (SIDRA 1737) |

### Regenerar

```bash
python data/osu_2025/download_osu_2025.py
python scripts/build_osu_2025_anexos.py
python scripts/build_osu_discriminativo_ipca.py
```

### Metodologia IPCA

- Fluxo do ano-calendário Y tratado como 31/12/Y
- `fator(Y) = IPCA(jun/2026) / IPCA(dez/Y)` (último IPCA publicado; jul/2026 sai em 11/08/2026)
- Fonte do índice: IBGE SIDRA tabela 1737

### Destaques 2024

- **Nominal:** R$ 678,4 bi · **IPCA jun/2026:** R$ 731,1 bi
- Tributários 83,1% · Financeiros 9,6% · Creditícios 7,3%
- Maiores itens (IPCA): Simples Nacional, Agricultura e Agroindústria, IRPF isentos
- **Acumulado 2003–2024 em R$ de jun/2026:** ~R$ 10,7 tri

## Renúncia fiscal SUDAM/SUDENE

Arquivo: `RENUNCIA FISCAL SUDAM-SUDENE (1).xlsx` (período 2015–2023, com atualização IPCA).
