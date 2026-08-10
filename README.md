# SEC--data-analysys

Repositório de análises e demonstrativos econômicos/fiscais.

## Contratações do FNE (2000–2025)

Demonstrativo do valor total das contratações do Fundo Constitucional de Financiamento do Nordeste (FNE), em valores correntes e atualizados pelo IPCA para 30/06/2026.

- Script: `scripts/build_fne_contratacoes.py`
- Excel: `output/fne_contratacoes_2000_2025_ipca.xlsx`
- Resumo: `output/fne_contratacoes_resumo.md`

### Fontes

- Portal BNB/ETENE – Relatórios FNE: https://bnb.gov.br/web/guest/etene/relatorios-fne
- Preferência pelo **Relatório de Gestão do FNE** de cada exercício
- Lacunas do content-set de Gestão (notadamente 2006–2013 e 2018–2020) preenchidas com **Relatórios de Resultados e Impactos / Atividades (RFNE)** no DSpace do BNB e, para conciliação 2007–2011, Demonstrações Financeiras do BNB
- IPCA: BCB SGS 433 (`data/raw/bcb_series/433_ipca.json`), índice médio do ano da contratação atualizado para o índice de jun/2026

### Como regenerar

```bash
python3 scripts/build_fne_contratacoes.py
```
