# Baixa bndes-mcp no ContAgil WinPython (sem .venv).
# ASCII-only. BAT sem BOM.
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bndes-operacoes-mcp-41ca/bndes-mcp/baixar_bndes_winpython.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_bndes.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_bndes.ps1
#
# Depois (Embraer exemplo):
#   .\bndes-mcp\bndes_cli.bat cnpj 07689002000189 --out embraer_bndes.xlsx

param(
    [string]$Ref = "cursor/bndes-operacoes-mcp-41ca",
    [string]$Python = "",
    [string]$UserAgent = "SEC-data-analysys-bndes-mcp/0.1 (cafla19791@gmail.com)"
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref/bndes-mcp"
$dest = Join-Path $Root "bndes-mcp"
$headers = @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" }

Write-Host "============================================================"
Write-Host " Baixando bndes-mcp -> $dest"
Write-Host "============================================================"

$files = @(
    "pyproject.toml",
    "README.md",
    "src/bndes_mcp/__init__.py",
    "src/bndes_mcp/cli.py",
    "src/bndes_mcp/providers.py",
    "src/bndes_mcp/server.py",
    "src/bndes_mcp/excel_export.py",
    "src/bndes_mcp/excel_format.py"
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
$batPath = Join-Path $dest "bndes_cli.bat"
$bat = @"
@echo off
setlocal
cd /d "%~dp0"
if "%BNDES_USER_AGENT%"=="" set BNDES_USER_AGENT=$UserAgent
set PYTHONPATH=%~dp0src;%PYTHONPATH%
"$py" -m bndes_mcp.cli %*
exit /b %ERRORLEVEL%
"@
[System.IO.File]::WriteAllText($batPath, $bat, $utf8NoBom)
Write-Host "BAT: $batPath"

Write-Host "Checando dependencias (httpx/pandas/openpyxl)..."
& $py -c "import httpx,pandas,openpyxl; print('ok', httpx.__version__, pandas.__version__)"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Instalando httpx pandas openpyxl xlsxwriter..."
    & $py -m pip install --user "httpx>=0.28" "pandas>=2" "openpyxl>=3.1" "xlsxwriter>=3.1"
}

Write-Host ""
Write-Host "Pronto. Exemplos:"
Write-Host "  .\bndes-mcp\bndes_cli.bat cnpj 07689002000189 --out embraer_bndes.xlsx"
Write-Host "  .\bndes-mcp\bndes_cli.bat resumo 07689002000189"
Write-Host "============================================================"
