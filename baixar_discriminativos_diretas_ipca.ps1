# Baixa discriminativos DIRETAS+IPCA no ContAgil WinPython.
# ASCII-only. Rode na pasta winpython:
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/discriminativos-diretas-ipca-e4e9/baixar_discriminativos_diretas_ipca.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_discriminativos.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_discriminativos.ps1

param(
    [string]$Ref = "cursor/discriminativos-diretas-ipca-e4e9",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref"
$headers = @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" }

Write-Host "============================================================"
Write-Host " Baixando discriminativos DIRETAS+IPCA -> $Root"
Write-Host "============================================================"

$files = @(
    "discriminativos_diretas_ipca.bat",
    "discriminativos_diretas_ipca.py",
    "scripts/__init__.py",
    "scripts/discriminativos_diretas_ipca.py",
    "scripts/calcular_diretas_ipca_selic.py",
    "scripts/gerar_fluxos.py"
)

foreach ($rel in $files) {
    $out = Join-Path $Root ($rel -replace "/", "\")
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

Write-Host "Instalando pandas/openpyxl/httpx/requests (se preciso)..."
& $py -m pip install "pandas>=2.0" "openpyxl>=3.1" "httpx>=0.28" "requests>=2.28"
if ($LASTEXITCODE -ne 0) { throw "pip install falhou" }

# Sanity check dos arquivos baixados
$need = @(
    "discriminativos_diretas_ipca.bat",
    "scripts\__init__.py",
    "scripts\discriminativos_diretas_ipca.py",
    "scripts\calcular_diretas_ipca_selic.py",
    "scripts\gerar_fluxos.py"
)
foreach ($rel in $need) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path -LiteralPath $p)) { throw "Arquivo faltando apos download: $rel" }
    if ((Get-Item -LiteralPath $p).Length -lt 50) { throw "Arquivo suspeito (muito pequeno): $rel" }
}
Write-Host "Checagem de arquivos: OK"

if (-not (Test-Path -LiteralPath (Join-Path $Root "saida"))) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root "saida") | Out-Null
}

Write-Host ""
Write-Host "OK. Arquivos prontos."
Write-Host "1) Deixe a planilha OPERACOES DIRETAS nesta pasta winpython"
Write-Host "2) Rode:"
Write-Host "     .\discriminativos_diretas_ipca.bat"
Write-Host "   (no cmd: discriminativos_diretas_ipca.bat)"
Write-Host "3) Saida: saida\DISCRIMINATIVOS_DIRETAS_IPCA.xlsx"
