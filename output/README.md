Gere o CSV detalhado e o resumo por agente com:

```bash
# amostra
python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --stem fluxos_amostra

# período completo
python3 scripts/gerar_fluxos.py --download
```

Arquivos principais:

- `fluxos_completos_final.xlsx` — abas Resumo, **Por_Agente**, Impacto_Mensal, Amostra_Parcelas
- `resumo_por_agente.csv` / `.xlsx` — ranking por Instituição Financeira Credenciada
- `fluxos_amostra.xlsx` — amostra rápida com colunas detalhadas ContAgil

Colunas detalhadas: Instituição Financeira, taxa_selic_mensal, taxa_contrato_mensal,
spread, subsidio, impacto_fiscal, em_carencia.

Interface web: `streamlit run app.py`
