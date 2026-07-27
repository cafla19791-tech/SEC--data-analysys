# Instala o nyse-mcp (venv) e testa a CLI no Windows.
# Nao precisa de Cursor Desktop. Usa Yahoo Finance (cotacao atrasada, sem API key).
#
# Uso (PowerShell):
#   cd caminho\para\SEC--data-analysys\nyse-mcp
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_e_rodar_cli.ps1
#
# Opcional: passar um ticker
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_e_rodar_cli.ps1 -Symbol JPM
#
# Opcional: apontar para o python do ContAgil WinPython
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_e_rodar_cli.ps1 `
#     -Python "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\python.exe"

param(
    [string]$Symbol = "JPM",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

Write-Host "============================================================"
Write-Host " nyse-mcp CLI — setup Windows (sem Cursor Desktop)"
Write-Host " Pasta: $Root"
Write-Host "============================================================"

function Find-Python {
    param([string]$Preferred)

    if ($Preferred -and (Test-Path -LiteralPath $Preferred)) {
        return $Preferred
    }

    $candidates = @(
        (Join-Path $Root "python.exe"),
        "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\python.exe"
    )

    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c)) { return $c }
    }

    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -notmatch "WindowsApps") {
        return $cmd.Source
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { return "py" }

    throw "python.exe nao encontrado. Informe -Python caminho\python.exe"
}

$basePython = Find-Python -Preferred $Python
Write-Host "Python base: $basePython"

$venvDir = Join-Path $Root ".venv"
$venvPython = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host ""
    Write-Host "Criando venv em .venv ..."
    if ($basePython -eq "py") {
        & py -3 -m venv $venvDir
    } else {
        & $basePython -m venv $venvDir
    }
}

if (-not (Test-Path -LiteralPath $venvPython)) {
    throw "Falha ao criar .venv\Scripts\python.exe"
}

Write-Host "Instalando pacote nyse-mcp (editavel) ..."
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -e "$Root"
if ($LASTEXITCODE -ne 0) {
    throw "pip install -e . falhou (exit $LASTEXITCODE). Rede da empresa pode bloquear PyPI."
}

if (-not $env:MARKET_DATA_PROVIDER) {
    $env:MARKET_DATA_PROVIDER = "yahoo"
}

Write-Host ""
Write-Host "Testes rapidos (provider=$env:MARKET_DATA_PROVIDER) ..."
Write-Host "--- status ---"
& $venvPython -m nyse_mcp.cli status
Write-Host ""
Write-Host "--- quote $Symbol ---"
& $venvPython -m nyse_mcp.cli quote $Symbol
Write-Host ""
Write-Host "--- fundamentals $Symbol ---"
& $venvPython -m nyse_mcp.cli fundamentals $Symbol

Write-Host ""
Write-Host "OK. Exemplos manuais:"
Write-Host "  .\.venv\Scripts\python.exe -m nyse_mcp.cli quote JPM"
Write-Host "  .\.venv\Scripts\python.exe -m nyse_mcp.cli history XOM --period 1y"
Write-Host "  .\.venv\Scripts\python.exe -m nyse_mcp.cli search `"Coca Cola`""
Write-Host "  .\.venv\Scripts\python.exe -m nyse_mcp.cli status"
Write-Host ""
Write-Host "Atalho BAT: .\nyse_mcp_cli.bat quote JPM"
Write-Host "Se a empresa bloquear finance.yahoo.com, o quote falha — use Cloud Agent + MCP."
