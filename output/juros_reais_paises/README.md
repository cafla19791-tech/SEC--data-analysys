# Discriminativo — taxas básicas de juros reais por país

Gere a planilha com os bulk downloads do [BIS Data Portal](https://data.bis.org/bulkdownload):

```bash
python3 scripts/discriminativo_juros_reais_paises.py
python3 scripts/discriminativo_juros_reais_paises.py --ano-inicio 2000
```

Saída: `output/discriminativo_juros_reais_paises.xlsx`

- **Capa** — fontes, fórmula de Fisher e marker
- **Resumo** — último mês e real acumulada do último ano completo
- **Anual** — comparação país × ano da real acumulada
- **Uma aba por país** — Mês/ano, taxa básica nominal (% a.a.), índice de inflação oficial (2010=100), inflação no mês, taxa real no mês; linha **ACUMULADO AAAA** após dezembro

Fórmula: `r_m = (1+i_aa)^(1/12) / (1 + IPC_t/IPC_{t-1} − 1) − 1` e `R_ano = Π(1+r_m) − 1`.
