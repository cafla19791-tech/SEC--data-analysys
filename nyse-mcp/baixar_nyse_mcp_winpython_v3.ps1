# Baixa nyse-mcp no ContAgil WinPython (v3).
# Sem .venv: o WinPython ContAgil nao tem venvlauncher.exe.
# Instala deps no proprio python.exe da pasta winpython.
#
# Cole EXATO:
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/nyse-mcp-winpython-cli-f342/nyse-mcp/baixar_nyse_mcp_winpython_v3.ps1"
#   Invoke-WebRequest "$u`?v=20260727d" -OutFile baixar_nyse_mcp_v3.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_nyse_mcp_v3.ps1 -Symbol JPM

param(
    [string]$Symbol = "JPM",
    [string]$Ref = "cursor/nyse-mcp-winpython-cli-f342",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$baseCdn = "https://cdn.jsdelivr.net/gh/cafla19791-tech/SEC--data-analysys@$Ref/nyse-mcp"
$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref/nyse-mcp"
$dest = Join-Path $Root "nyse-mcp"
$headers = @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" }

Write-Host "============================================================"
Write-Host " Baixando nyse-mcp (v3 sem-venv) -> $dest"
Write-Host " Ref: $Ref"
Write-Host "============================================================"

$files = @(
    "pyproject.toml",
    "README.md",
    "run_mcp.sh",
    ".env.example",
    "src/nyse_mcp/__init__.py",
    "src/nyse_mcp/cli.py",
    "src/nyse_mcp/providers.py",
    "src/nyse_mcp/server.py"
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
    $ok = $false
    try {
        Invoke-WebRequest -Uri "$baseCdn/$rel" -OutFile $out -UseBasicParsing -Headers $headers
        $ok = $true
    } catch {
        $ok = $false
    }
    if (-not $ok) {
        Write-Host "    fallback: raw.githubusercontent.com"
        Invoke-WebRequest -Uri "$baseRaw/$rel`?v=20260727d" -OutFile $out -UseBasicParsing -Headers $headers
    }
}

# Resolve python do WinPython (pasta atual)
$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}
if (-not $py -or -not (Test-Path -LiteralPath $py)) {
    throw "python.exe nao encontrado em $Root"
}
Write-Host "Python: $py"

# Gera BAT que usa o python do WinPython (sem .venv)
$batPath = Join-Path $dest "nyse_mcp_cli.bat"
$bat = @"
@echo off
setlocal
cd /d "%~dp0"
if "%MARKET_DATA_PROVIDER%"=="" set MARKET_DATA_PROVIDER=yahoo
set PYTHONPATH=%~dp0src;%PYTHONPATH%
"$py" -m nyse_mcp.cli %*
exit /b %ERRORLEVEL%
"@
# CMD.exe nao aceita BOM no .bat (aparece '´╗┐@echo' e falha a 1a linha)
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($batPath, $bat, $utf8NoBom)
Write-Host "  nyse_mcp_cli.bat (WinPython direto, sem BOM)"

# Gera setup LOCAL (sem venv)
$setupPath = Join-Path $dest "setup_e_rodar_cli.ps1"
$setup = @"
param(
    [string]`$Symbol = "JPM",
    [string]`$Python = "$py"
)

`$ErrorActionPreference = "Stop"
`$Root = if (`$PSScriptRoot) { `$PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath `$Root

Write-Host "============================================================"
Write-Host " nyse-mcp CLI - setup Windows [local-v3-sem-venv]"
Write-Host " Pasta: `$Root"
Write-Host "============================================================"

if (-not (Test-Path -LiteralPath `$Python)) {
    throw "Python nao encontrado: `$Python"
}
Write-Host "Python: `$Python"

`$env:PYTHONPATH = (Join-Path `$Root "src") + ";" + `$env:PYTHONPATH
if (-not `$env:MARKET_DATA_PROVIDER) { `$env:MARKET_DATA_PROVIDER = "yahoo" }

Write-Host "Garantindo pip..."
& `$Python -m ensurepip --upgrade 2>`$null
& `$Python -m pip install --upgrade pip
if (`$LASTEXITCODE -ne 0) {
    Write-Host "Aviso: pip upgrade falhou; tentando install mesmo assim."
}

Write-Host "Instalando dependencias (yfinance, mcp, httpx) no WinPython..."
& `$Python -m pip install "yfinance>=0.2.54" "httpx>=0.28.0" "mcp[cli]>=1.6.0"
if (`$LASTEXITCODE -ne 0) {
    throw "pip install deps falhou (exit `$LASTEXITCODE). Rede pode bloquear PyPI."
}

# Pacote local via PYTHONPATH (src/) - nao precisa pip install -e
Write-Host ""
Write-Host "Testes rapidos (provider=`$env:MARKET_DATA_PROVIDER) ..."
Write-Host "--- status ---"
& `$Python -m nyse_mcp.cli status
if (`$LASTEXITCODE -ne 0) { throw "status falhou" }
Write-Host ""
Write-Host "--- quote `$Symbol ---"
& `$Python -m nyse_mcp.cli quote `$Symbol
if (`$LASTEXITCODE -ne 0) { throw "quote falhou (rede/Yahoo bloqueado?)" }
Write-Host ""
Write-Host "--- fundamentals `$Symbol ---"
& `$Python -m nyse_mcp.cli fundamentals `$Symbol

Write-Host ""
Write-Host "OK. Use:"
Write-Host "  .\nyse_mcp_cli.bat quote JPM"
Write-Host "  .\nyse_mcp_cli.bat history XOM --period 1y"
"@

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($setupPath, $setup, $utf8Bom)
Write-Host "  setup_e_rodar_cli.ps1 (gerado local [local-v3-sem-venv])"

$check = Get-Content -LiteralPath $setupPath -Raw
if ($check -notmatch 'local-v3-sem-venv') {
    throw "setup gerado invalido"
}

Write-Host ""
Write-Host "Rodando setup da CLI..."
& powershell -NoProfile -ExecutionPolicy Bypass -File $setupPath -Symbol $Symbol -Python $py
if ($LASTEXITCODE -ne 0) {
    throw "setup falhou (exit $LASTEXITCODE)"
}

Write-Host ""
Write-Host "Pronto. Comandos:"
Write-Host "  cd nyse-mcp"
Write-Host "  .\nyse_mcp_cli.bat quote JPM"
Write-Host "  .\nyse_mcp_cli.bat history XOM --period 1y"
Write-Host "  .\nyse_mcp_cli.bat search `"Coca Cola`""
Write-Host "  .\nyse_mcp_cli.bat status"
