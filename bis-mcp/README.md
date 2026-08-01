# bis-mcp

MCP + CLI para **taxas de política monetária do BIS** (`WS_CBPOL` / `WS_CBPOL_csv_flat`).

Fontes públicas (sem API key):

- [SDMX REST](https://stats.bis.org/api/v1/data/WS_CBPOL) — séries filtradas por país
- [Bulk flat CSV](https://data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip) — `WS_CBPOL_csv_flat.csv`
- Portal: [Central bank policy rates](https://data.bis.org/topics/CBPOL)

Para o Brasil, a série BIS corresponde à **meta Selic** (fim do período).

## ContAgil WinPython

O arquivo que você apontou:

```text
C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\WS_CBPOL_csv_flat.csv
```

é o dump flat do BIS. Instale a CLI na mesma pasta:

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bis-cbpol-mcp-41ca/bis-mcp/baixar_bis_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_bis.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bis.ps1
```

Se o CSV ainda não existir, use `-DownloadCsv`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bis.ps1 -DownloadCsv
```

Depois:

```bat
cd bis-mcp
bis_cli.bat catalog
bis_cli.bat serie brasil --last 12
bis_cli.bat serie BR,US,XM --from 2020-01
bis_cli.bat compare BR,US,XM,GB,JP
bis_cli.bat serie selic --local --last 24
bis_cli.bat download --dir ..
```

`--local` lê o `WS_CBPOL_csv_flat.csv` irmão da pasta `winpython` (sem rede SDMX).

O CSV flat descompactado tem **~450 MB**. Para o ContAgil, prefira SDMX (`serie` sem `--local`) ou extraia só os países necessários:

```bat
bis_cli.bat download --dir ..
bis_cli.bat extract BR,US,XM --out ..\cbpol_BR_US_XM.csv
bis_cli.bat serie BR --local --csv ..\cbpol_BR_US_XM.csv --last 24
```

## Tools (MCP)

| Tool | Uso |
|------|-----|
| `list_known_areas` | Aliases locais (brasil/selic, us/fed, euro...) |
| `get_policy_rates` | Série por país / lista de países |
| `compare_latest` | Snapshot da última taxa |
| `download_flat_csv` | Baixa o ZIP flat e extrai o CSV |
| `local_csv_info` | Localiza o CSV ContAgil |

## Cadastro no Cursor Cloud

Veja **[CADASTRO_MCP_CLOUD.md](./CADASTRO_MCP_CLOUD.md)**.

## CLI (Linux / cloud)

```bash
cd bis-mcp
export PYTHONPATH=src
python -m bis_mcp.cli catalog
python -m bis_mcp.cli serie brasil --last 5
python -m bis_mcp.cli compare BR,US,XM
python -m bis_mcp.cli download --dir /tmp
python -m bis_mcp.cli serie BR --local --csv /tmp/WS_CBPOL_csv_flat.csv --last 3
```

## Aliases úteis

| Alias | Código | Descrição |
|-------|--------|-----------|
| `brasil` / `selic` | `BR` | Meta Selic (BIS policy rate) |
| `us` / `fed` | `US` | Estados Unidos |
| `euro` / `ecb` | `XM` | Zona do euro |
| `uk` / `gb` | `GB` | Reino Unido |
| `jp` | `JP` | Japão |

Qualquer código ISO de 2 letras do BIS também funciona.

## Limites

- Preferir `--last` ou janelas curtas; diário (`--freq D`) gera payloads maiores.
- O CSV flat completo tem histórico longo (desde ~1945) e ~450 MB descompactado; `--local` no arquivo inteiro leva alguns segundos.
- Cite a fonte nacional + BIS ao republicar os dados.
