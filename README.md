# SEC--data-analysys

Ferramentas para análise de dados financeiros a partir da **SEC EDGAR**, com foco inicial no **lucro líquido da Petrobras** (últimos 10 anos).

## Petrobras — Lucro líquido (Net Income)

Script principal:

```bash
pip install -r requirements.txt
python extract_petrobras_net_income.py --years 10
```

Atualizar dados (força download SEC + câmbio BCB):

```bash
python extract_petrobras_net_income.py --years 10 --refresh
```

### Saídas

| Arquivo | Conteúdo |
|---------|----------|
| `data/petrobras_net_income.csv` | Série anual: R$, USD, FX, YoY |
| `data/petrobras_net_income.json` | Mesmo conteúdo + metadados |
| `data/raw/petrobras_CIK0001119639_companyfacts.json` | Cache CompanyFacts SEC |
| `data/usdbrl_annual_avg.json` | Médias anuais USD/BRL (BCB) |
| `reports/petrobras_net_income_table.md` | Tabela formatada |
| `reports/petrobras_net_income_analysis.md` | Tendências e eventos |
| `reports/petrobras_net_income_chart.png` | Gráfico matplotlib |
| `reports/petrobras_net_income_chart.html` | Gráfico interativo plotly |

### Metodologia

- **USD:** `ProfitLossAttributableToOwnersOfParent` (IFRS) nos formulários **20-F** (FY), via CompanyFacts (`CIK 0001119639`).
- **R$:** USD × média anual da taxa de câmbio USD/BRL (BCB SGS série 1 — dólar venda).
- **YoY:** variação percentual ano a ano (denominador em módulo quando a base é negativa).

### Gráfico (código sugerido)

O módulo `petrobras/charts.py` já gera PNG/HTML. Snippet equivalente:

```python
import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_csv("data/petrobras_net_income.csv")
years = df["year"]
usd_bi = df["net_income_usd"] / 1e9

fig, ax = plt.subplots(figsize=(10, 5))
colors = ["#c0392b" if v < 0 else "#1f6f4a" for v in usd_bi]
ax.bar(years, usd_bi, color=colors)
ax.axhline(0, color="black", lw=0.8)
ax.set_title("Petrobras — Lucro líquido anual (US$ bi)")
ax.set_xlabel("Ano")
ax.set_ylabel("US$ bilhões")
plt.tight_layout()
plt.savefig("reports/petrobras_net_income_chart.png", dpi=150)
```

### Testes

```bash
pytest -q
```

### User-Agent SEC

A SEC exige um User-Agent identificável (`Nome email@dominio`). Padrão do projeto: `SEC-Data-Analysis cafla19791@gmail.com` (override com `--user-agent`).
