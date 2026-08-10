# SEC--data-analysys

Repositório de análises e demonstrativos econômicos/fiscais.

## Contratações do FNO (2000–2021)

Demonstrativo a partir da Figura 3 do livro *O impacto do FNO em dados e ciência* (Banco da Amazônia), p. 79.

- Script: `scripts/build_fno_contratacoes.py`
- Excel: `output/fno_contratacoes_2000_2021_ipca.xlsx`
- Resumo: `output/fno_contratacoes_resumo.md`

### Conteúdo do Excel

1. **Demonstrativo_reversao_IGPDI** — reverte o IGP-DI (rótulo do gráfico) e atualiza pelo IPCA até 30/06/2026  
2. **Alternativa_grafico_como_corrente** — trata os níveis do gráfico como correntes (mais coerente com totais oficiais) e atualiza pelo IPCA  
3. **Oficial_parcial_IPCA** — totais oficiais publicados (anos disponíveis) atualizados pelo IPCA  
4. **Confronto_oficial** / **Metodologia**

### Como regenerar

```bash
python3 scripts/build_fno_contratacoes.py
```
