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

### Pasta ContAgil (vários Excel) — equivalente a `pasta_dados`

```bash
python scripts/gerar_fluxos_contratos.py \
  --pasta-dados "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\dados" \
  --pasta-saida "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\saida"
```

### Excel STP Selic + impacto capitalizado — equivalente a `arquivo_selic`

```bash
python scripts/gerar_fluxos_contratos.py \
  --pasta-dados "...\winpython\dados" \
  --arquivo-selic "...\winpython\STP-20260607091256352_para_05062026.xlsx" \
  --data-impacto 2026-06-30 \
  --pasta-saida "...\winpython\saida"
```

Com `--arquivo-selic`, o script usa o mesmo `fator_rapido` do ContAgil e
grava a coluna `impacto` (= subsídio capitalizado até a data de referência).

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

## Saídas (`output/`)

| Arquivo | Conteúdo |
|---------|----------|
| `fluxos_0.csv`, `fluxos_1.csv`, … | Lotes de parcelas (default 50.000 linhas) |
| `fluxos_resumo.xlsx` | Totais por contrato |

Colunas dos lotes: `contrato`, `mes`, `data_fluxo`, `saldo`, `amortizacao`,
`taxa_mensal`, `subsidio`, `em_carencia` (+ `impacto` se Selic STP informada).

## Metodologia

Para cada mês `t = 0 .. carência+amortização−1`:

| Item | Fórmula |
|------|---------|
| Data | `data_contratação + (t+1)` meses |
| Taxa mensal (TAXA FIXA) | `(1 + juros)^(1/12) − 1` |
| Taxa mensal (TJLP/TLP) | `(1 + 0,06 + juros)^(1/12) − 1` |
| Amortização | `valor / n` só após a carência; 0 na carência |
| Subsídio | `saldo × (Selic_aa/12 − taxa_mensal)` |
| Impacto (opcional) | `subsídio × fator(data_impacto) / fator(data_fluxo)` |

Selic padrão no subsídio nominal: **14,5% a.a.** (igual ao ContAgil).

Fonte aberta: [Portal de Dados Abertos do BNDES — Operações indiretas automáticas](https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento).
