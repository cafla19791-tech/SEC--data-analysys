# Baixa tesouro-mcp no ContAgil WinPython (sem .venv).
# ASCII-only. BAT sem BOM.
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/tesouro-mcp-f342/tesouro-mcp/baixar_tesouro_winpython.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_tesouro.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_tesouro.ps1

param(
    [string]$Ref = "cursor/tesouro-mcp-f342",
    [string]$Python = "",
    [string]$UserAgent = "SEC-data-analysys-tesouro-mcp/0.1 (cafla19791@gmail.com)"
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref/tesouro-mcp"
$dest = Join-Path $Root "tesouro-mcp"
$headers = @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" }

Write-Host "============================================================"
Write-Host " Baixando tesouro-mcp -> $dest"
Write-Host "============================================================"

$files = @(
    "pyproject.toml",
    "README.md",
    "run_mcp.sh",
    "src/tesouro_mcp/__init__.py",
    "src/tesouro_mcp/cli.py",
    "src/tesouro_mcp/providers.py",
    "src/tesouro_mcp/bcb_client.py",
    "src/tesouro_mcp/collector.py",
    "src/tesouro_mcp/server.py",
    "data/templates/dgt_renuncias_anual.csv",
    "data/templates/fundos_constitucionais_anual.csv",
    "data/templates/INSTRUCOES_COLA_DGT_FUNDOS.md"
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
$batPath = Join-Path $dest "tesouro_cli.bat"
$bat = @"
@echo off
setlocal
cd /d "%~dp0"
if "%TESOURO_USER_AGENT%"=="" set TESOURO_USER_AGENT=$UserAgent
set PYTHONPATH=%~dp0src;%PYTHONPATH%
"$py" -m tesouro_mcp.cli %*
exit /b %ERRORLEVEL%
"@
[System.IO.File]::WriteAllText($batPath, $bat, $utf8NoBom)
Write-Host "  tesouro_cli.bat"

$env:TESOURO_USER_AGENT = $UserAgent
$env:PYTHONPATH = (Join-Path $dest "src") + ";" + $env:PYTHONPATH

Write-Host "Instalando httpx + openpyxl (se preciso)..."
& $py -m pip install "httpx>=0.28.0" "openpyxl>=3.1.0"
if ($LASTEXITCODE -ne 0) { throw "pip install falhou" }

Write-Host ""
Write-Host "Testes..."
Write-Host "--- aliases ---"
& $py -m tesouro_mcp.cli aliases
if ($LASTEXITCODE -ne 0) { throw "aliases falhou" }
Write-Host "--- serie resultado_primario --from 2025-01 --to 2025-03 ---"
& $py -m tesouro_mcp.cli serie resultado_primario --from 2025-01 --to 2025-03
if ($LASTEXITCODE -ne 0) { throw "serie falhou (Tesouro/rede?)" }
Write-Host "--- headline ---"
& $py -m tesouro_mcp.cli headline resultado_primario

Write-Host ""
Write-Host "OK. Comandos (PowerShell - use .\ ):"
Write-Host "  cd tesouro-mcp"
Write-Host "  .\tesouro_cli.bat aliases"
Write-Host "  .\tesouro_cli.bat serie resultado_primario --from 2024-01 --to 2025-12"
Write-Host "  .\tesouro_cli.bat serie receita_total --from 01/2024 --to 12/2024"
Write-Host "  .\tesouro_cli.bat headline"
Write-Host "  .\tesouro_cli.bat search primario"
Write-Host "  .\tesouro_cli.bat ckan-show resultado-do-tesouro-nacional"
Write-Host "  .\tesouro_cli.bat coletar-anual --from 2001 --to 2025 --out tabela_anual.csv --dgt data\templates\dgt_renuncias_anual.csv --fundos data\templates\fundos_constitucionais_anual.csv"
Write-Host ""
Write-Host "Cole DGT/FNO-FNE-FCO em data\templates\ (ver INSTRUCOES_COLA_DGT_FUNDOS.md)"
Write-Host "Se 'tesouro_cli.bat' nao for reconhecido: cd para a pasta tesouro-mcp e use .\tesouro_cli.bat"
