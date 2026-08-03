# Baixa discriminativos DIRETAS+IPCA no ContAgil WinPython.
# Usa pasta sec_scripts (evita colisao com WinPython\Scripts).
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/discriminativos-diretas-ipca-e4e9/baixar_discriminativos_diretas_ipca.ps1"
#   Invoke-WebRequest "$u`?v=5" -OutFile baixar_discriminativos.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_discriminativos.ps1

param(
    [string]$Ref = "cursor/discriminativos-diretas-ipca-e4e9",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

# Cache-bust unico por execucao
$bust = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref"
$headers = @{
    "Cache-Control" = "no-cache"
    "Pragma" = "no-cache"
}

Write-Host "============================================================"
Write-Host " Baixando discriminativos DIRETAS+IPCA -> $Root"
Write-Host " Pasta de codigo: sec_scripts (nao Scripts)"
Write-Host " Cache-bust: $bust"
Write-Host "============================================================"

$map = @(
    @{ Remote = "discriminativos_diretas_ipca.bat"; Local = "discriminativos_diretas_ipca.bat" },
    @{ Remote = "discriminativos_diretas_ipca.py"; Local = "discriminativos_diretas_ipca.py" },
    @{ Remote = "scripts/__init__.py"; Local = "sec_scripts/__init__.py" },
    @{ Remote = "scripts/discriminativos_diretas_ipca.py"; Local = "sec_scripts/discriminativos_diretas_ipca.py" },
    @{ Remote = "scripts/calcular_diretas_ipca_selic.py"; Local = "sec_scripts/calcular_diretas_ipca_selic.py" },
    @{ Remote = "scripts/gerar_fluxos.py"; Local = "sec_scripts/gerar_fluxos.py" }
)

foreach ($item in $map) {
    $out = Join-Path $Root ($item.Local -replace "/", "\")
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

# Bloqueia versao antiga com import quebrado
$calc = Join-Path $Root "sec_scripts\calcular_diretas_ipca_selic.py"
$calcText = Get-Content -LiteralPath $calc -Raw
if ($calcText -match "from scripts\.gerar_fluxos import _excel_tem_colunas_contratos") {
    throw "Download trouxe versao ANTIGA de calcular_diretas_ipca_selic.py (import scripts.*). Tente de novo."
}
if ($calcText -notmatch "_load_gerar_fluxos") {
    throw "Download incompleto: falta _load_gerar_fluxos em calcular_diretas_ipca_selic.py"
}
Write-Host "Versao calcular_diretas_ipca_selic.py: OK (sem import scripts.* quebrado)"

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

$need = @(
    "discriminativos_diretas_ipca.bat",
    "discriminativos_diretas_ipca.py",
    "sec_scripts\__init__.py",
    "sec_scripts\discriminativos_diretas_ipca.py",
    "sec_scripts\calcular_diretas_ipca_selic.py",
    "sec_scripts\gerar_fluxos.py"
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
Write-Host "OK. Proximo comando:"
Write-Host ".\discriminativos_diretas_ipca.bat"
