# Análise — Lucro líquido Petrobras (últimos 10 anos)

<!-- AUTO-TABLE:START -->
## Lucro líquido — PETROBRAS - PETROLEO BRASILEIRO SA (PBR)

**Métrica:** Net Income / Lucro Líquido atribuível aos acionistas
**CIK:** 0001119639

| Ano | Lucro líquido (R$) | Lucro líquido (US$) | YoY (R$) | YoY (US$) | FX médio |
|----:|-------------------:|--------------------:|---------:|----------:|---------:|
| 2015 | -R$ 28,21 bi | -US$ 8.45 bi | — | — | 3.3387 |
| 2016 | -R$ 16,85 bi | -US$ 4.84 bi | +40.3% | +42.7% | 3.4833 |
| 2017 | -R$ 0,29 bi | -US$ 0.09 bi | +98.3% | +98.1% | 3.1925 |
| 2018 | R$ 26,22 bi | US$ 7.17 bi | +9126.2% | +7982.4% | 3.6558 |
| 2019 | R$ 40,06 bi | US$ 10.15 bi | +52.8% | +41.5% | 3.9461 |
| 2020 | R$ 5,89 bi | US$ 1.14 bi | -85.3% | -88.8% | 5.1578 |
| 2021 | R$ 107,24 bi | US$ 19.88 bi | +1722.2% | +1641.9% | 5.3956 |
| 2022 | R$ 189,18 bi | US$ 36.62 bi | +76.4% | +84.3% | 5.1655 |
| 2023 | R$ 124,30 bi | US$ 24.88 bi | -34.3% | -32.1% | 4.9953 |
| 2024 | R$ 40,59 bi | US$ 7.53 bi | -67.3% | -69.7% | 5.3920 |

## Notas metodológicas

- **USD:** Valores em USD extraídos da SEC EDGAR CompanyFacts (IFRS, 20-F FY).
- **BRL:** Valores em R$ estimados como USD × média anual da taxa de câmbio USD/BRL (BCB SGS série 1, venda). Podem diferir ligeiramente dos valores oficiais em reais publicados na CVM/DFP.
- **YoY:** variação percentual ano a ano; base negativa usa denominador em módulo.

<!-- AUTO-TABLE:END -->

## Visão geral das tendências

A série anual de lucro líquido atribuível aos acionistas (IFRS, formulário **20-F** na SEC) mostra três fases claras:

1. **2015–2017 — crise e recuperação inicial:** prejuízos elevados em 2015–2016, ligados à queda do petróleo, desvalorização do real, impairments e desdobramentos da Operação Lava Jato / ajustes de governança. Em 2017 o resultado quase zera (ligeiro prejuízo atribuível).
2. **2018–2019 — normalização:** retorno consistente ao lucro com preços do Brent mais firmes, disciplina de custos e desalavancagem.
3. **2020–2024 — choque, boom e normalização:** 2020 sofre o choque COVID (demanda e preço); 2021–2022 registram lucros recordes com Brent elevado e diferencial de preços no mercado doméstico; 2023–2024 recuam a partir da base excepcional de 2022 (queda de preços, maior carga tributária / participação governamental e efeitos de preços de realização).

Em dólares, o lucro sai de cerca de **−US$ 8,5 bi (2015)** para o pico de **~US$ 36,6 bi (2022)** e depois modera para **~US$ 7,5 bi (2024)**.

## Principais eventos que impactaram os resultados

| Período | Evento / contexto | Efeito no lucro |
|--------:|-------------------|-----------------|
| 2015–2016 | Colapso do petróleo, Real fraco, impairments e custos da crise de governança | Prejuízos bilionários |
| 2017 | Estabilização operacional; preço ainda moderado | Resultado próximo de zero |
| 2018–2019 | Política de preços alinhada ao mercado internacional; venda de ativos; redução de dívida | Lucros sólidos (~US$ 7–10 bi) |
| 2020 | COVID-19: destruição de demanda e Brent em mínimos | Lucro comprimido (~US$ 1,1 bi) |
| 2021 | Reabertura global + alta de commodities | Salto para ~US$ 19,9 bi |
| 2022 | Guerra na Ucrânia / choque energético; Brent elevado | Recorde ~US$ 36,6 bi |
| 2023 | Normalização de preços; mudanças em política de preços e tributos | Recuo YoY, ainda elevado (~US$ 24,9 bi) |
| 2024 | Brent mais baixo vs. 2022–23; mix de realização e custos | Nova desaceleração (~US$ 7,5 bi) |

## Leitura da variação YoY

- As variações percentuais em anos que saem de prejuízo para lucro (ex.: 2017→2018) ou vice-versa são muito amplas — interprete a **magnitude absoluta** (US$ / R$ bilhões) junto com o %.
- Em **R$**, parte da variação YoY também reflete o câmbio médio anual (PTAX/BCB). Por isso o script reporta YoY em ambas as moedas.
- O valor em reais é **estimado** (USD × câmbio médio anual). Para fechamento contábil oficial em BRL, use as DFPs na CVM; a SEC (20-F) é a fonte autoritativa em USD neste projeto.

## Como atualizar no futuro

```bash
# Recarrega SEC + BCB e regenera CSV/JSON/gráficos/tabela
python extract_petrobras_net_income.py --years 10 --refresh
```

Sem `--refresh`, o script reutiliza o cache em `data/raw/` e `data/usdbrl_annual_avg.json` (útil offline / CI).
