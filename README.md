# SEC--data-analysys

Repositório de análises e demonstrativos econômicos/fiscais.

## Contratações do FNO — Relatórios da Administração BASA (2011–2026)

Série extraída da [Central de Resultados](https://ri.bancoamazonia.com.br/informacoes-financeiras/central-de-resultados/) do Banco da Amazônia (RA/DF 4T; 2026 = 1T26), com atualização pelo IPCA até 30/06/2026.

- Script: `scripts/build_fno_basa_ra.py`
- Excel: `output/fno_contratacoes_2011_2026_basa_ra.xlsx`
- Resumo: `output/fno_contratacoes_2011_2026_basa_ra.md`

```bash
python3 scripts/build_fno_basa_ra.py
```

## Contratações do FNO — Figura 3 do livro BASA (2000–2021)

Demonstrativo a partir da Figura 3 do livro *O impacto do FNO em dados e ciência* (Banco da Amazônia), p. 79.

- Script: `scripts/build_fno_contratacoes.py`
- Excel: `output/fno_contratacoes_2000_2021_ipca.xlsx`
- Resumo: `output/fno_contratacoes_resumo.md`

### Conteúdo do Excel (livro)

1. **Demonstrativo_reversao_IGPDI** — reverte o IGP-DI (rótulo do gráfico) e atualiza pelo IPCA até 30/06/2026  
2. **Alternativa_grafico_como_corrente** — trata os níveis do gráfico como correntes (mais coerente com totais oficiais) e atualiza pelo IPCA  
3. **Oficial_parcial_IPCA** — totais oficiais publicados (anos disponíveis) atualizados pelo IPCA  
4. **Confronto_oficial** / **Metodologia**

```bash
python3 scripts/build_fno_contratacoes.py
```
