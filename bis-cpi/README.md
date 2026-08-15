# bis-cpi — BIS WS_LONG_CPI (ContAgil)

Relatórios no mesmo espírito do CBPOL: **1 aba por país**, rankings de **inflação acumulada** por períodos e **PDF**.

Fonte: [WS_LONG_CPI_csv_col.zip](https://data.bis.org/static/bulk/WS_LONG_CPI_csv_col.zip) (63 países; índice 2010=100 e YoY %).

## ContAgil (PowerShell)

```powershell
cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
$u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bis-long-cpi-reports-b311/bis-cpi/baixar_cpi_winpython.ps1"
Invoke-WebRequest "$u`?v=1" -OutFile baixar_cpi.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_cpi.ps1 -DownloadCsv
```

```bat
cd bis-cpi
bis_cpi_cli.bat excel-mensal --out ..\cpi_mensal_por_pais.xlsx
bis_cpi_cli.bat excel-periodos --out ..\cpi_inflacao_acumulada_periodos.xlsx
bis_cpi_cli.bat para-pdf ..\cpi_mensal_por_pais.xlsx --outdir ..\pdf
bis_cpi_cli.bat para-pdf ..\cpi_inflacao_acumulada_periodos.xlsx --outdir ..\pdf
```

Só PDFs prontos:

```powershell
Invoke-WebRequest "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bis-long-cpi-reports-b311/bis-cpi/baixar_pdfs_cpi.ps1?v=1" -OutFile baixar_pdfs_cpi.ps1 -Headers @{"Cache-Control"="no-cache"}
powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_pdfs_cpi.ps1
```

## Método

- **Mensal por país:** `Mês | Índice (2010=100) | Variação YoY (%) | Inflação acumulada (%)`
- **Acumulada na série:** `(Índice_t / Índice_primeiro − 1) × 100`
- **Ranking por período:** `(Índice_fim / Índice_início − 1) × 100` nos mesmos recortes do CBPOL (1995–2002, 2003–abr/2016, mai/2016–2018, 2019–2022, 2023–jun/2026, 2003–jun/2026)

Saídas versionadas: `output/cpi/`.
