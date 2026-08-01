# Baixa bis-mcp no ContAgil WinPython (sem .venv).
# ASCII-only. BAT sem BOM.
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bis-cbpol-mcp-41ca/bis-mcp/baixar_bis_winpython.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_bis.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bis.ps1
#
# Opcional: tambem baixa WS_CBPOL_csv_flat.csv nesta pasta:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bis.ps1 -DownloadCsv

param(
    [string]$Ref = "cursor/bis-cbpol-mcp-41ca",
    [string]$Python = "",
    [string]$UserAgent = "SEC-data-analysys-bis-mcp/0.1 (cafla19791@gmail.com)",
    [switch]$DownloadCsv
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref/bis-mcp"
$dest = Join-Path $Root "bis-mcp"
$headers = @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" }

Write-Host "============================================================"
Write-Host " Baixando bis-mcp -> $dest"
Write-Host "============================================================"

$files = @(
    "pyproject.toml",
    "README.md",
    "run_mcp.sh",
    "src/bis_mcp/__init__.py",
    "src/bis_mcp/cli.py",
    "src/bis_mcp/providers.py",
    "src/bis_mcp/server.py",
    "src/bis_mcp/excel_diario.py",
    "src/bis_mcp/excel_periodos.py",
    "src/bis_mcp/excel_mensal.py",
    "src/bis_mcp/excel_format.py",
    "src/bis_mcp/pdf_export.py"
)

if (-not (Test-Path -LiteralPath $dest)) {
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
}

foreach ($rel in $files) {
    $out = Join-Path $dest ($rel -replace "/", "\")
    $dir = Split-Path -Parent $out
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    Write-Host "  $rel"
    Invoke-WebRequest -Uri "$baseRaw/$rel`?v=1" -OutFile $out -UseBasicParsing -Headers $headers
}

$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}
if (-not $py -or -not (Test-Path -LiteralPath $py)) {
    throw "python.exe nao encontrado em $Root"
}
Write-Host "Python: $py"

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$batPath = Join-Path $dest "bis_cli.bat"
$bat = @"
@echo off
setlocal
cd /d "%~dp0"
if "%BIS_USER_AGENT%"=="" set BIS_USER_AGENT=$UserAgent
if "%BIS_CBPOL_CSV%"=="" if exist "%~dp0..\WS_CBPOL_csv_flat.csv" set BIS_CBPOL_CSV=%~dp0..\WS_CBPOL_csv_flat.csv
set PYTHONPATH=%~dp0src;%PYTHONPATH%
"$py" -m bis_mcp.cli %*
exit /b %ERRORLEVEL%
"@
[System.IO.File]::WriteAllText($batPath, $bat, $utf8NoBom)
Write-Host "  bis_cli.bat"

$env:BIS_USER_AGENT = $UserAgent
$env:PYTHONPATH = (Join-Path $dest "src") + ";" + $env:PYTHONPATH
$csvSibling = Join-Path $Root "WS_CBPOL_csv_flat.csv"
if (Test-Path -LiteralPath $csvSibling) {
    $env:BIS_CBPOL_CSV = $csvSibling
}

Write-Host "Instalando httpx (se preciso)..."
& $py -m pip install "httpx>=0.28.0"
if ($LASTEXITCODE -ne 0) { throw "pip install httpx falhou" }

if ($DownloadCsv) {
    Write-Host "Baixando WS_CBPOL_csv_flat.csv para $Root ..."
    & $py -m bis_mcp.cli download --dir $Root
    if ($LASTEXITCODE -ne 0) { throw "download CSV falhou (BIS/rede?)" }
    $env:BIS_CBPOL_CSV = Join-Path $Root "WS_CBPOL_csv_flat.csv"
}

Write-Host ""
Write-Host "Testes..."
Write-Host "--- catalog ---"
& $py -m bis_mcp.cli catalog
if ($LASTEXITCODE -ne 0) { throw "catalog falhou" }
Write-Host "--- serie brasil --last 3 ---"
& $py -m bis_mcp.cli serie brasil --last 3
if ($LASTEXITCODE -ne 0) { throw "serie falhou (BIS/rede?)" }
Write-Host "--- compare BR,US,XM ---"
& $py -m bis_mcp.cli compare BR,US,XM
if ($LASTEXITCODE -ne 0) { throw "compare falhou" }

Write-Host ""
Write-Host "OK. Comandos:"
Write-Host "  cd bis-mcp"
Write-Host "  .\bis_cli.bat catalog"
Write-Host "  .\bis_cli.bat serie brasil --last 12"
Write-Host "  .\bis_cli.bat serie BR,US,XM --from 2020-01 --to 2026-07"
Write-Host "  .\bis_cli.bat compare BR,US,XM,GB,JP"
Write-Host "  .\bis_cli.bat download --dir .."
Write-Host "  .\bis_cli.bat extract BR,US,XM --out ..\cbpol_BR_US_XM.csv"
Write-Host "  .\bis_cli.bat excel-diario --out ..\cbpol_taxas_diarias_compostas.xlsx"
Write-Host "  .\bis_cli.bat excel-mensal --out ..\cbpol_taxas_mensais_compostas.xlsx"
Write-Host "  .\bis_cli.bat excel-periodos --out ..\cbpol_taxas_acumuladas_periodos.xlsx"
Write-Host "  .\bis_cli.bat excel-periodos --freq M --out ..\cbpol_taxas_acumuladas_periodos_mensal.xlsx"
Write-Host "  .\bis_cli.bat serie selic --local --last 12"
Write-Host ""
Write-Host "CSV ContAgil esperado (~450MB descompactado):"
Write-Host "  $csvSibling"
