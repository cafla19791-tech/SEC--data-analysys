# Baixa contagil_fluxos_seguro no ContAgil WinPython (sec_scripts).
# Apenas ASCII neste arquivo (evita erro de encoding no PowerShell Windows).
#
# No CMD (Prompt de Comando), rode:
#
#   cd /d "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest 'https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/numerar-contratos-indiretas-e4e9/baixar_contagil_fluxos_seguro.ps1?v=3' -OutFile baixar_fluxos.ps1 -Headers @{'Cache-Control'='no-cache'}"
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_fluxos.ps1
#
# Depois (teste rapido - 50 contratos):
#   contagil_fluxos_seguro.bat "dados\Operacoes Indiretas 2002.xlsx" 50
#
# Massa completa:
#   contagil_fluxos_seguro.bat

param(
    [string]$Ref = "cursor/numerar-contratos-indiretas-e4e9",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$bust = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref"
$headers = @{
    "Cache-Control" = "no-cache"
    "Pragma" = "no-cache"
}

Write-Host "============================================================"
Write-Host " Baixando contagil_fluxos_seguro -> $Root"
Write-Host " Pasta de codigo: sec_scripts (nao Scripts)"
Write-Host " Cache-bust: $bust"
Write-Host "============================================================"

$map = @(
    @{ Remote = "contagil_fluxos_seguro.bat"; Local = "contagil_fluxos_seguro.bat" },
    @{ Remote = "contagil_fluxos_seguro.py"; Local = "contagil_fluxos_seguro.py" },
    @{ Remote = "scripts/__init__.py"; Local = "sec_scripts/__init__.py" },
    @{ Remote = "scripts/contagil_fluxos_seguro.py"; Local = "sec_scripts/contagil_fluxos_seguro.py" },
    @{ Remote = "scripts/gerar_fluxos.py"; Local = "sec_scripts/gerar_fluxos.py" }
)

foreach ($item in $map) {
    $out = Join-Path $Root ($item.Local -replace "/", [string][char]92)
    $dir = Split-Path -Parent $out
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    if (Test-Path -LiteralPath $out) {
        Remove-Item -LiteralPath $out -Force
    }
    $uri = "$baseRaw/$($item.Remote)?v=$bust"
    Write-Host ("  {0} -> {1}" -f $item.Remote, $item.Local)
    Invoke-WebRequest -Uri $uri -OutFile $out -UseBasicParsing -Headers $headers
}

$pyMod = Join-Path $Root "sec_scripts\contagil_fluxos_seguro.py"
$txt = Get-Content -LiteralPath $pyMod -Raw
if ($txt -match "from scripts\.gerar_fluxos import") {
    throw "Download trouxe versao ANTIGA (import scripts.*). Rode o baixar de novo."
}
if ($txt -notmatch "_load_gerar_fluxos") {
    throw "Download incompleto: falta _load_gerar_fluxos em contagil_fluxos_seguro.py"
}
Write-Host "Versao contagil_fluxos_seguro.py: OK"

$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}
if (-not $py -or -not (Test-Path -LiteralPath $py)) {
    throw "python.exe nao encontrado em $Root"
}
Write-Host "Python: $py"

Write-Host "Instalando pandas/openpyxl/numpy/requests (se preciso)..."
& $py -m pip install "pandas>=2.0" "openpyxl>=3.1" "numpy>=1.24" "requests>=2.28"
if ($LASTEXITCODE -ne 0) { throw "pip install falhou" }

$need = @(
    "contagil_fluxos_seguro.bat",
    "contagil_fluxos_seguro.py",
    "sec_scripts\__init__.py",
    "sec_scripts\contagil_fluxos_seguro.py",
    "sec_scripts\gerar_fluxos.py"
)
foreach ($rel in $need) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path -LiteralPath $p)) { throw "Arquivo faltando apos download: $rel" }
    if ((Get-Item -LiteralPath $p).Length -lt 50) { throw "Arquivo suspeito (muito pequeno): $rel" }
}

$fatores = Join-Path $Root "fator_acumulado_SELIC_TJLP_TLP.xlsx"
if (-not (Test-Path -LiteralPath $fatores)) {
    Write-Host ""
    Write-Host "AVISO: falta fator_acumulado_SELIC_TJLP_TLP.xlsx nesta pasta."
    Write-Host "Coloque o arquivo de fatores antes de rodar o .bat"
} else {
    Write-Host "Fatores: OK ($fatores)"
}

$saidaDir = Join-Path $Root "saida"
if (-not (Test-Path -LiteralPath $saidaDir)) {
    New-Item -ItemType Directory -Force -Path $saidaDir | Out-Null
}

Write-Host ""
Write-Host "OK. Proximo passo (recomendado - teste rapido):"
Write-Host '  contagil_fluxos_seguro.bat "dados\Operacoes Indiretas 2002.xlsx" 50'
Write-Host ""
Write-Host "Massa completa (demora; arquivos grandes em saida):"
Write-Host "  contagil_fluxos_seguro.bat"
