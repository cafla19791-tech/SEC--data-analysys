# Discriminativo — taxas básicas de juros reais por país

Gere a planilha com os bulk downloads do [BIS Data Portal](https://data.bis.org/bulkdownload):

```bash
python3 scripts/discriminativo_juros_reais_paises.py
python3 scripts/discriminativo_juros_reais_paises.py --ano-inicio 1995
```

Saída: `output/discriminativo_juros_reais_paises.xlsx`

- **Capa** — fontes, fórmula de Fisher e marker
- **Resumo** — último mês e real acumulada do último ano completo
- **Anual** — comparação país × ano da real acumulada
- **Uma aba por país** — Mês/ano, taxa básica nominal (% a.a.), índice de inflação oficial (2010=100), inflação no mês, taxa real no mês; linha **ACUMULADO AAAA** após dezembro

Fórmula: `r_m = (1+i_aa)^(1/12) / (1 + IPC_t/IPC_{t-1} − 1) − 1` e `R_ano = Π(1+r_m) − 1`.

## Ranking anual (1995–2026)

```bash
python3 scripts/discriminativo_ranking_juros_reais.py
```

Saída: `output/discriminativo_ranking_juros_reais.xlsx`

- **Capa** / **Resumo** (pódio + posição do Brasil) / **Brasil**
- **Uma aba por ano** (1995–2026): posição, país, real acumulada, inflação acumulada, nominal composta, meses e cobertura
- Anos com dezembro e 12 meses entram no ranking oficial; 2026 é parcial (acumulado até o último mês BIS)

## PDF e HTML (navegação entre abas)

O PDF **não** tem abas nativas. O conversor gera:

- **PDF** com página Índice, marcadores (sumário do leitor) e links anterior/próxima
- **HTML** com faixa de abas clicáveis, equivalente ao Excel

```bash
python3 scripts/discriminativo_para_pdf.py --entrada output/discriminativo_ranking_juros_reais.xlsx
python3 scripts/discriminativo_para_pdf.py --entrada output/discriminativo_juros_reais_paises.xlsx
```
