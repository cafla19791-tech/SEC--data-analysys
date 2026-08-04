# Baixa bcb-mcp no ContAgil WinPython (sem .venv).
# ASCII-only. BAT sem BOM.
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bcb-mcp-f342/bcb-mcp/baixar_bcb_winpython.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_bcb.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bcb.ps1

param(
    [string]$Ref = "cursor/bcb-mcp-f342",
    [string]$Python = "",
    [string]$UserAgent = "SEC-data-analysys-bcb-mcp/0.1 (cafla19791@gmail.com)"
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref/bcb-mcp"
$dest = Join-Path $Root "bcb-mcp"
$headers = @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" }

Write-Host "============================================================"
Write-Host " Baixando bcb-mcp -> $dest"
Write-Host "============================================================"

$files = @(
    "pyproject.toml",
    "README.md",
    "run_mcp.sh",
    "src/bcb_mcp/__init__.py",
    "src/bcb_mcp/cli.py",
    "src/bcb_mcp/providers.py",
    "src/bcb_mcp/server.py"
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
$batPath = Join-Path $dest "bcb_cli.bat"
$bat = @"
@echo off
setlocal
cd /d "%~dp0"
if "%BCB_USER_AGENT%"=="" set BCB_USER_AGENT=$UserAgent
set PYTHONPATH=%~dp0src;%PYTHONPATH%
"$py" -m bcb_mcp.cli %*
exit /b %ERRORLEVEL%
"@
[System.IO.File]::WriteAllText($batPath, $bat, $utf8NoBom)
Write-Host "  bcb_cli.bat"

$env:BCB_USER_AGENT = $UserAgent
$env:PYTHONPATH = (Join-Path $dest "src") + ";" + $env:PYTHONPATH

Write-Host "Instalando httpx (se preciso)..."
& $py -m pip install "httpx>=0.28.0"
if ($LASTEXITCODE -ne 0) { throw "pip install httpx falhou" }

Write-Host ""
Write-Host "Testes..."
Write-Host "--- catalog ---"
& $py -m bcb_mcp.cli catalog
if ($LASTEXITCODE -ne 0) { throw "catalog falhou" }
Write-Host "--- serie selic --last 3 ---"
& $py -m bcb_mcp.cli serie selic --last 3
if ($LASTEXITCODE -ne 0) { throw "serie falhou (BCB/rede?)" }

Write-Host ""
Write-Host "OK. Comandos:"
Write-Host "  cd bcb-mcp"
Write-Host "  .\bcb_cli.bat catalog"
Write-Host "  .\bcb_cli.bat serie selic --last 10"
Write-Host "  .\bcb_cli.bat serie ipca --from 2020-01-01 --to 2025-12-31"
Write-Host "  .\bcb_cli.bat ptax --days 7"
Write-Host "  .\bcb_cli.bat expectativas IPCA --top 10"
