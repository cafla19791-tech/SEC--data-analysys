# SEC--data-analysys

Repositório de análises e demonstrativos econômicos/fiscais.

## OSU 2025 — Fundos Constitucionais (FNE + FNO + FCO)

Extração dos anexos do Orçamento de Subsídios da União (MPO): benefício creditício implícito agregado dos fundos constitucionais, 2003–2024.

- Script: `scripts/build_osu_fundos_constitucionais.py`
- Excel: `output/osu_2025_fundos_constitucionais.xlsx`
- Resumo: `output/osu_2025_fundos_constitucionais.md`

```bash
python3 scripts/build_osu_fundos_constitucionais.py
```

**Atenção:** a série mede *custo de subsídio*, não volume contratado. O rateio regional do OSU usa chave fixa 20% Norte / 60% Nordeste / 20% Centro-Oeste.
