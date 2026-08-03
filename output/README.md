Gere o CSV detalhado e o resumo por agente com:

```bash
# amostra
python scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv

# período completo
python scripts/gerar_fluxos.py --download
```

Arquivos principais:

- `fluxos_completos_corrigido.xlsx` — abas Resumo, **Por_Agente**, Impacto_Mensal, Amostra_Parcelas
- `resumo_por_agente.csv` / `.xlsx` — ranking por Instituição Financeira Credenciada
- `amostra_fluxos.xlsx` — primeiras 1.000 parcelas (export opcional)

Interface web: `streamlit run app.py`
