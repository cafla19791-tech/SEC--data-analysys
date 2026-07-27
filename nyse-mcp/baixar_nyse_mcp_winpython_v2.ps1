# Baixa nyse-mcp no ContAgil WinPython (v2 - anti-cache).
# Setup .ps1 e gerado LOCALMENTE (ASCII), nao depende do raw antigo em cache.
#
# Cole EXATO no PowerShell:
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://cdn.jsdelivr.net/gh/cafla19791-tech/SEC--data-analysys@cursor/nyse-mcp-winpython-cli-f342/nyse-mcp/baixar_nyse_mcp_winpython_v2.ps1"
#   Invoke-WebRequest $u -OutFile baixar_nyse_mcp_v2.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_nyse_mcp_v2.ps1 -Symbol JPM

param(
    [string]$Symbol = "JPM",
    [string]$Ref = "cursor/nyse-mcp-winpython-cli-f342",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

# jsDelivr costuma furar cache corporativo do raw.githubusercontent.com
$base = "https://cdn.jsdelivr.net/gh/cafla19791-tech/SEC--data-analysys@$Ref/nyse-mcp"
$dest = Join-Path $Root "nyse-mcp"

Write-Host "============================================================"
Write-Host " Baixando nyse-mcp (v2 anti-cache) -> $dest"
Write-Host " Ref: $Ref"
Write-Host "============================================================"

# Apenas fontes Python / metadados (sem .ps1 remoto)
$files = @(
    "pyproject.toml",
    "README.md",
    "nyse_mcp_cli.bat",
    "run_mcp.sh",
    ".env.example",
    "src/nyse_mcp/__init__.py",
    "src/nyse_mcp/cli.py",
    "src/nyse_mcp/providers.py",
    "src/nyse_mcp/server.py"
)

$headers = @{
    "Cache-Control" = "no-cache"
    "Pragma"        = "no-cache"
}

if (-not (Test-Path -LiteralPath $dest)) {
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
}

foreach ($rel in $files) {
    $url = "$base/$rel"
    $out = Join-Path $dest ($rel -replace "/", "\")
    $dir = Split-Path -Parent $out
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    Write-Host "  $rel"
    try {
        Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing -Headers $headers
    } catch {
        # Fallback raw GitHub com commit/branch + query anti-cache
        $fb = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref/nyse-mcp/$rel`?v=20260727c"
        Write-Host "    fallback: raw.githubusercontent.com"
        Invoke-WebRequest -Uri $fb -OutFile $out -UseBasicParsing -Headers $headers
    }
}

# Gera setup LOCAL em ASCII puro (nunca baixa o .ps1 antigo do cache)
$setupPath = Join-Path $dest "setup_e_rodar_cli.ps1"
$setup = @'
param(
    [string]$Symbol = "JPM",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

Write-Host "============================================================"
Write-Host " nyse-mcp CLI - setup Windows (sem Cursor Desktop) [local-v2]"
Write-Host " Pasta: $Root"
Write-Host "============================================================"

function Find-Python {
    param([string]$Preferred)
    if ($Preferred -and (Test-Path -LiteralPath $Preferred)) { return $Preferred }
    $candidates = @(
        (Join-Path $Root "python.exe"),
        "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython\python.exe"
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c)) { return $c }
    }
    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -notmatch "WindowsApps") { return $cmd.Source }
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { return "py" }
    throw "python.exe nao encontrado. Informe -Python caminho\python.exe"
}

$basePython = Find-Python -Preferred $Python
Write-Host "Python base: $basePython"

$venvDir = Join-Path $Root ".venv"
$venvPython = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "Criando venv em .venv ..."
    if ($basePython -eq "py") { & py -3 -m venv $venvDir }
    else { & $basePython -m venv $venvDir }
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

if (-not $env:MARKET_DATA_PROVIDER) { $env:MARKET_DATA_PROVIDER = "yahoo" }

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
Write-Host "OK. Use: .\nyse_mcp_cli.bat quote JPM"
'@

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($setupPath, $setup, $utf8Bom)
Write-Host "  setup_e_rodar_cli.ps1 (gerado local [local-v2])"

# Verificacao anti-cache
$check = Get-Content -LiteralPath $setupPath -Raw
if ($check -notmatch '\[local-v2\]') {
    throw "setup gerado invalido - abortando"
}
if ($check -match [char]0x2014) {
    throw "setup contem em-dash - abortando"
}

$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}

Write-Host ""
Write-Host "Rodando setup da CLI..."
if ($py) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $setupPath -Symbol $Symbol -Python $py
} else {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $setupPath -Symbol $Symbol
}
$code = $LASTEXITCODE
if ($code -ne 0) {
    throw "setup_e_rodar_cli.ps1 falhou (exit $code)"
}

Write-Host ""
Write-Host "Pronto. Comandos:"
Write-Host "  cd nyse-mcp"
Write-Host "  .\nyse_mcp_cli.bat quote JPM"
Write-Host "  .\nyse_mcp_cli.bat history XOM --period 1y"
Write-Host "  .\nyse_mcp_cli.bat search `"Coca Cola`""
Write-Host "  .\nyse_mcp_cli.bat status"
