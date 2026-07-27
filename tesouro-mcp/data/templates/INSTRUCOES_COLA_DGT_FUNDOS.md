# Como colar DGT e FNO/FNE/FCO

O coletor anual (`tesouro_cli.bat coletar-anual`) preenche automaticamente:

- DBGG 1/jan e 31/dez (BCB)
- Resultado primario, juros e nominal (RTN/ARIA)
- Emissoes e resgates da DPF (Tesouro Transparente)
- Desembolsos BNDES (BCB)

As colunas abaixo **dependem de voce colar valores** nos CSVs deste diretorio (unidade: **R$ bilhoes**, valores correntes).

## 1) Renuncias (DGT / Receita Federal)

Arquivo: `dgt_renuncias_anual.csv`

| Coluna | Conteudo tipico no DGT |
|--------|-------------------------|
| `renuncia_desenv_regional_R$bi` | SUDAM/SUDENE, ZFM, incentivos regionais |
| `renuncia_imunes_isentas_R$bi` | Entidades imunes/isentas |
| `renuncia_automotivo_R$bi` | Setor automotivo |
| `renuncia_cultura_audiovisual_R$bi` | Cultura / audiovisual |
| `renuncia_inovacao_R$bi` | Inovacao (Lei do Bem etc.) |

Fonte sugerida: Demonstrativo de Gastos Tributarios (bases efetivas ou PLOA) da RFB.

Exemplo de linha preenchida:

```csv
2020,45.2,12.1,4.0,1.5,8.3
```

## 2) Fundos constitucionais

Arquivo: `fundos_constitucionais_anual.csv`

| Coluna | Banco administrador |
|--------|---------------------|
| `financ_FNO_BASA_R$bi` | Banco da Amazonia (FNO) |
| `financ_FNE_BNB_R$bi` | Banco do Nordeste (FNE) |
| `financ_FCO_BB_R$bi` | Banco do Brasil (FCO) |

Use **contratacoes** ou **desembolsos** de forma consistente (indique no nome do arquivo se mudar). Fonte: relatorios de gestao MDR / bancos administradores.

Exemplo:

```csv
2023,12.5,40.1,10.2
```

## 3) Rodar o coletor com merge

```powershell
# 1) Se ainda nao instalou (cria a pasta tesouro-mcp e o .bat):
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/tesouro-mcp-f342/tesouro-mcp/baixar_tesouro_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_tesouro.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_tesouro.ps1

# 2) Rodar o coletor (PowerShell exige .\):
cd tesouro-mcp
.\tesouro_cli.bat coletar-anual --from 2001 --to 2025 --out tabela_anual.csv --dgt data\templates\dgt_renuncias_anual.csv --fundos data\templates\fundos_constitucionais_anual.csv
```

Celulas vazias / `n/d` sao ignoradas no merge (mantem nulo na tabela final).
