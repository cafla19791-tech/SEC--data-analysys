# SEC Data Analysis

Script Python para analisar dados financeiros de empresas públicas usando a API SEC EDGAR (XBRL CompanyFacts) ou arquivos CSV locais.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso rápido

```bash
# Demonstração offline (dados de exemplo)
python analyze_finance.py --demo

# Empresa específica no CSV de exemplo
python analyze_finance.py --csv data/sample_companies.csv --company MSFT

# Dados ao vivo da SEC EDGAR
python analyze_finance.py --ticker AAPL
python analyze_finance.py --ticker MSFT --years 7 --export-json resultado.json

# Comparar várias empresas
python analyze_finance.py --compare AAPL MSFT TSLA
```

A SEC exige um `User-Agent` identificável. Ajuste se necessário:

```bash
python analyze_finance.py --ticker AAPL --user-agent "SeuNome seu@email.com"
```

## O que o script calcula

| Indicador | Descrição |
|-----------|-----------|
| Receita / Lucro líquido | Séries anuais (formulários 10-K) |
| Margem líquida e operacional | Lucratividade sobre receita |
| ROE / ROA | Retorno sobre patrimônio e ativos |
| Dívida / PL | Alavancagem |
| Liquidez corrente | Ativo circulante / passivo circulante |
| CAGR de receita | Crescimento composto no período |

## CSV local

O arquivo deve ter pelo menos a coluna `year` e métricas numéricas. Colunas suportadas:

`ticker`, `name`, `year`, `revenue`, `net_income`, `total_assets`, `total_liabilities`, `equity`, `operating_income`, `current_assets`, `current_liabilities`, `cash`

Veja `data/sample_companies.csv` como referência.

## Estrutura

```
analyze_finance.py          # CLI principal
financial_analyzer/
  sec_client.py             # Cliente SEC EDGAR
  metrics.py                # Indicadores e crescimento
  report.py                 # Relatório texto / exportação
data/
  sample_companies.csv      # Dados de demonstração
```
