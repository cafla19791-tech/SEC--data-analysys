# SEC--data-analysys

Repositório de análises e demonstrativos econômicos/fiscais.

## Discriminativo — Benefícios tributários selecionados (2003–2025)

Nove gastos tributários (OSU 2025 Anexos para 2003–2024; projeção PLDO 2025 para 2025), em valores correntes e atualizados pelo IPCA até 30/06/2026:

1. Desenvolvimento Regional  
2. Entidades Sem Fins Lucrativos – Imunes e Isentas  
3. Pesquisas Científicas e Tecnológicas  
4. Informática e Automação  
5. Zona Franca de Manaus  
6. Cultura e Audiovisual  
7. Regime Automotivo (Setor Automotivo)  
8. Água Mineral  
9. Fundos Constitucionais (gasto tributário)

- Script: `scripts/build_osu_beneficios_tributarios.py`
- Excel: `output/osu_beneficios_tributarios_2003_2025_ipca.xlsx`
- Resumo: `output/osu_beneficios_tributarios_2003_2025_ipca.md`

```bash
python3 scripts/build_osu_beneficios_tributarios.py
```

## OSU 2025 — Fundos Constitucionais creditícios (FNE + FNO + FCO)

Benefício **creditício** implícito agregado dos fundos constitucionais, 2003–2024 (distinto do gasto tributário “Fundos Constitucionais” acima).

- Script: `scripts/build_osu_fundos_constitucionais.py`
- Excel: `output/osu_2025_fundos_constitucionais.xlsx`
- Resumo: `output/osu_2025_fundos_constitucionais.md`

```bash
python3 scripts/build_osu_fundos_constitucionais.py
```
