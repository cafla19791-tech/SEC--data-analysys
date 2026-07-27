# Cadastro do tesouro-mcp no Cursor Cloud

## O que este MCP faz

Consulta **estatisticas fiscais do Tesouro Nacional**:

1. Series mensais do Boletim Resultado do Tesouro Nacional (RTN)
2. Grandes numeros (resultado primario, estoque DPF...)
3. Datasets abertos no Tesouro Transparente (CKAN)
4. Coletor anual (DBGG + RTN + emissoes/resgates DPF + BNDES; merge DGT/FNO-FNE-FCO via CSV)

## A) Registrar no cursor.com/agents

| Campo | Valor |
|-------|--------|
| **Name** | `tesouro-mcp` |
| **Transport** | `stdio` |
| **Command** | `bash` |
| **Args** | `./tesouro-mcp/run_mcp.sh` |
| **Env** | `TESOURO_USER_AGENT` = `SEC-data-analysys-tesouro-mcp/0.1 (seu@email.com)` |

Exemplos:

```text
Qual o resultado primario do governo central em 2024-2025?
Serie mensal de receita liquida desde 2020.
Quais os grandes numeros fiscais atuais do Tesouro?
```

## B) ContAgil WinPython

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/tesouro-mcp-f342/tesouro-mcp/baixar_tesouro_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_tesouro.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_tesouro.ps1
```

```powershell
cd tesouro-mcp
.\tesouro_cli.bat aliases
.\tesouro_cli.bat serie resultado_primario --from 2024-01 --to 2025-12
.\tesouro_cli.bat headline resultado_primario
.\tesouro_cli.bat search "Receita"
.\tesouro_cli.bat coletar-anual --from 2001 --to 2025 --out tabela_anual.csv --dgt data\templates\dgt_renuncias_anual.csv --fundos data\templates\fundos_constitucionais_anual.csv
```

No PowerShell use sempre `.\tesouro_cli.bat` (o ponto-barra). Se o arquivo nao existir, rode o instalador acima primeiro.

Cole DGT / FNO-FNE-FCO em `data\templates\` (ver `INSTRUCOES_COLA_DGT_FUNDOS.md`).

## Arquitetura

```
Cursor / CLI
  -> tesouro-mcp
       -> apiapex.tesouro.gov.br/aria/.../resultado-fiscal  (RTN)
       -> grandesnumeros.tesouro.gov.br                     (headlines)
       -> tesourotransparente.gov.br/ckan/api               (datasets)
       -> api.bcb.gov.br/dados/serie                        (DBGG / BNDES)
       -> CSV overlay local                                 (DGT / FNO-FNE-FCO)
```
