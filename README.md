# SEC — Fluxos de contratos (SAC + carência + subsídio Selic)

Gera fluxos mensais de contratos no sistema SAC, incluindo meses de carência,
e estima o **subsídio implícito** (Selic de referência vs taxa do contrato).

Baseado no script ContAgil (`python_jep` / WinPython).

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

### Dados de exemplo (smoke test)

```bash
python scripts/gerar_fluxos_contratos.py --input data/sample_contratos.csv
```

### Pasta com vários Excel ContAgil / transparência

```bash
python scripts/gerar_fluxos_contratos.py --pasta-dados /caminho/para/dados
```

### Excel único (header na linha 6)

```bash
python scripts/gerar_fluxos_contratos.py --input operacoes.xlsx --excel-header 5
```

### Download do Portal de Dados Abertos do BNDES (2009–2010)

```bash
python scripts/gerar_fluxos_contratos.py --download
```

Teste rápido (N contratos):

```bash
python scripts/gerar_fluxos_contratos.py --download --limit-contracts 100
```

### Série Selic (Excel STP) — opcional

```bash
python scripts/gerar_fluxos_contratos.py --input data/sample_contratos.csv \
  --arquivo-selic STP-....xlsx --selic 14.5
```

O arquivo STP é carregado para lookup de fatores; o subsídio usa `--selic`
(padrão 14,5% a.a.), como no script original.

## Saídas (`output/`)

| Arquivo | Conteúdo |
|---------|----------|
| `fluxos_0.csv`, `fluxos_1.csv`, … | Lotes de parcelas (default 50.000 linhas) |
| `fluxos_resumo.xlsx` | Totais por contrato |

Colunas dos lotes: `contrato`, `mes`, `data_fluxo`, `saldo`, `amortizacao`,
`taxa_mensal`, `subsidio`, `em_carencia`.

## Metodologia

Para cada mês `t = 0 .. carência+amortização−1`:

| Item | Fórmula |
|------|---------|
| Data | `data_contratação + (t+1)` meses |
| Taxa mensal (TAXA FIXA) | `(1 + juros)^(1/12) − 1` |
| Taxa mensal (TJLP/TLP) | `(1 + 0,06 + juros)^(1/12) − 1` |
| Amortização | `valor / n` só após a carência; 0 na carência |
| Subsídio | `saldo × (Selic_aa/12 − taxa_mensal)` |

Selic padrão: **14,5% a.a.**

Fonte aberta: [Portal de Dados Abertos do BNDES — Operações indiretas automáticas](https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento).
