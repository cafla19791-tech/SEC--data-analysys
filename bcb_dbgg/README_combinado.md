# DBGG contrafactual combinado

Simulação da Dívida Bruta do Governo Geral (BCB `Dbggindexp`) sob três premissas simultâneas:

1. **Sem gastos tributários** (OSU 2025): Desenvolvimento Regional; Pesquisas Científicas e Inovação Tecnológica; Informática e Automação
2. **SELIC a 4% a.a.** constante de jan/2007 ao fim da amostra
3. **Emissões líquidas de 2020 iguais a zero** (todos os indexadores)

## Fontes

- BCB: https://www.bcb.gov.br/content/estatisticas/Documents/Tabelas_especiais/Dbggindexp.xlsx
- MPO/OSU: anexos do Orçamento de Subsídios da União 2025

## Método

Ver `assumption` em `dbgg_gt_selic4_sem_emissoes_2020.json` e script `simulate_combined_counterfactual.py`.

## Saídas

- `dbgg_gt_selic4_sem_emissoes_2020.xlsx`
- `dbgg_gt_selic4_sem_emissoes_2020.json`
