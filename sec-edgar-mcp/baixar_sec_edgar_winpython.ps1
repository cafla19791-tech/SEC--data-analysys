# Baixa sec-edgar-mcp no ContAgil WinPython (sem .venv).
# ASCII-only. BAT sem BOM.
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/sec-edgar-mcp-f342/sec-edgar-mcp/baixar_sec_edgar_winpython.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_sec_edgar.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_sec_edgar.ps1

param(
    [string]$Symbol = "AAPL",
    [string]$Ref = "cursor/sec-edgar-mcp-f342",
    [string]$Python = "",
    [string]$UserAgent = "SEC-data-analysys/0.1 (cafla19791@gmail.com)"
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref/sec-edgar-mcp"
$dest = Join-Path $Root "sec-edgar-mcp"
$headers = @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" }

Write-Host "============================================================"
Write-Host " Baixando sec-edgar-mcp -> $dest"
Write-Host "============================================================"

$files = @(
    "pyproject.toml",
    "README.md",
    "run_mcp.sh",
    "src/sec_edgar_mcp/__init__.py",
    "src/sec_edgar_mcp/cli.py",
    "src/sec_edgar_mcp/providers.py",
    "src/sec_edgar_mcp/server.py"
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
$batPath = Join-Path $dest "sec_edgar_cli.bat"
$bat = @"
@echo off
setlocal
cd /d "%~dp0"
if "%SEC_USER_AGENT%"=="" set SEC_USER_AGENT=$UserAgent
set PYTHONPATH=%~dp0src;%PYTHONPATH%
"$py" -m sec_edgar_mcp.cli %*
exit /b %ERRORLEVEL%
"@
[System.IO.File]::WriteAllText($batPath, $bat, $utf8NoBom)
Write-Host "  sec_edgar_cli.bat"

$env:SEC_USER_AGENT = $UserAgent
$env:PYTHONPATH = (Join-Path $dest "src") + ";" + $env:PYTHONPATH

Write-Host "Instalando httpx (se preciso)..."
& $py -m pip install "httpx>=0.28.0"
if ($LASTEXITCODE -ne 0) { throw "pip install httpx falhou" }

Write-Host ""
Write-Host "Testes..."
Write-Host "--- lookup $Symbol ---"
& $py -m sec_edgar_mcp.cli lookup $Symbol
if ($LASTEXITCODE -ne 0) { throw "lookup falhou" }
Write-Host "--- filings $Symbol --form 10-K ---"
& $py -m sec_edgar_mcp.cli filings $Symbol --form 10-K --limit 3
if ($LASTEXITCODE -ne 0) { throw "filings falhou (SEC/rede?)" }
Write-Host "--- facts $Symbol ---"
& $py -m sec_edgar_mcp.cli facts $Symbol --limit 3

Write-Host ""
Write-Host "OK. Comandos:"
Write-Host "  cd sec-edgar-mcp"
Write-Host "  .\sec_edgar_cli.bat lookup AAPL"
Write-Host "  .\sec_edgar_cli.bat filings AAPL --form 10-K --limit 5"
Write-Host "  .\sec_edgar_cli.bat facts KO"
Write-Host "  .\sec_edgar_cli.bat concept PBR NetIncomeLoss"
Write-Host "  .\sec_edgar_cli.bat series PBR NetIncomeLoss --from 2008 --to 2025"
Write-Host "  .\sec_edgar_cli.bat debt PBR --from 2016 --to 2025"
Write-Host "  .\sec_edgar_cli.bat filing-xbrl PBR --form 20-F"
