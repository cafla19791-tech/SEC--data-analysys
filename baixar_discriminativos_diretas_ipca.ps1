# Baixa discriminativos DIRETAS+IPCA no ContAgil WinPython.
# Usa pasta sec_scripts (evita colisao com WinPython\Scripts).
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/discriminativos-diretas-ipca-e4e9/baixar_discriminativos_diretas_ipca.ps1"
#   Invoke-WebRequest "$u`?v=3" -OutFile baixar_discriminativos.ps1 -Headers @{"Cache-Control"="no-cache"}
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
Write-Host " Pasta de codigo: sec_scripts (nao Scripts)"
Write-Host "============================================================"

# remote path -> local path (repo scripts/* vira sec_scripts/*)
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
    Write-Host ("  {0} -> {1}" -f $item.Remote, $item.Local)
    Invoke-WebRequest -Uri "$baseRaw/$($item.Remote)`?v=4" -OutFile $out -UseBasicParsing -Headers $headers
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

# Smoke: import entrypoint
Write-Host "Teste rapido de import..."
& $py -c "import importlib.util; from pathlib import Path; p=Path(r'$Root')/'discriminativos_diretas_ipca.py'; print('runner', p.exists(), 'sec', (Path(r'$Root')/'sec_scripts'/'discriminativos_diretas_ipca.py').exists())"
if ($LASTEXITCODE -ne 0) { throw "teste de import falhou" }

if (-not (Test-Path -LiteralPath (Join-Path $Root "saida"))) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root "saida") | Out-Null
}

Write-Host ""
Write-Host "OK. Agora rode SOMENTE esta linha:"
Write-Host "  .\discriminativos_diretas_ipca.bat"
Write-Host "Saida: saida\DISCRIMINATIVOS_DIRETAS_IPCA.xlsx"
