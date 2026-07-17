Gere o CSV detalhado e o resumo por agente com:

```bash
# amostra (com fatores SELIC Bacen)
python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv --baixar-selic --stem fluxos_amostra

# período completo com fatores ContAgil/Bacen
python3 scripts/gerar_fluxos.py --download --baixar-selic
```

Arquivos principais:

- `fluxos_completos_final.xlsx` — abas Resumo, **Por_Agente**, Impacto_Mensal, Amostra_Parcelas (run completo 2009–2010)
- `resumo_por_agente.csv` / `.xlsx` — ranking por Instituição Financeira Credenciada
- `amostra_fluxos_detalhados.xlsx` — primeiras parcelas com colunas ContAgil
- `fluxos_amostra.xlsx` — amostra rápida (20 contratos de exemplo)

Colunas detalhadas: Instituição Financeira, taxa_selic_mensal, taxa_contrato_mensal,
spread, subsidio, impacto_fiscal, em_carencia.

Run completo (fatores SELIC Bacen SGS 11 → ContAgil): 348.864 contratos ·
22.151.051 parcelas · subsídio R$ 23,45 bi · impacto fiscal 2026 R$ 90,17 bi ·
72 agentes.

Interface web: `streamlit run app.py`
