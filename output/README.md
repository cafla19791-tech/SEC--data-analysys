Gere o CSV detalhado e o resumo por agente com:

```bash
# amostra
python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --stem fluxos_amostra

# período completo
python3 scripts/gerar_fluxos.py --download
```

Arquivos principais:

- `fluxos_completos_final.xlsx` — abas Resumo, **Por_Agente**, Impacto_Mensal, Amostra_Parcelas (run completo 2009–2010)
- `resumo_por_agente.csv` / `.xlsx` — ranking por Instituição Financeira Credenciada
- `amostra_fluxos_detalhados.xlsx` — primeiras 1.000 parcelas com colunas ContAgil
- `fluxos_amostra.xlsx` — amostra rápida (20 contratos de exemplo)

Colunas detalhadas: Instituição Financeira, taxa_selic_mensal, taxa_contrato_mensal,
spread, subsidio, impacto_fiscal, em_carencia.

Run completo (SELIC composta 14,5% a.a.): 348.864 contratos · 22.151.051 parcelas ·
subsídio R$ 23,45 bi · impacto fiscal 2026 R$ 164,45 bi · 72 agentes.

Interface web: `streamlit run app.py`
