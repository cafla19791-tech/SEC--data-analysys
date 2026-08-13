# Baixa numeracao de contratos INDIRETAS (N-AAAA) no ContAgil WinPython.
# Usa pasta sec_scripts (evita colisao com WinPython\Scripts).
#
# Cole no PowerShell (ja dentro da pasta winpython):
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/numerar-contratos-indiretas-e4e9/baixar_numerar_contratos_indiretas.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_numerar.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_numerar.ps1
#   .\numerar_contratos_indiretas.bat

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
Write-Host " Baixando numerar contratos INDIRETAS -> $Root"
Write-Host " Pasta de codigo: sec_scripts (nao Scripts)"
Write-Host " Cache-bust: $bust"
Write-Host "============================================================"

$map = @(
    @{ Remote = "numerar_contratos_indiretas.bat"; Local = "numerar_contratos_indiretas.bat" },
    @{ Remote = "numerar_contratos_indiretas.py"; Local = "numerar_contratos_indiretas.py" },
    @{ Remote = "scripts/__init__.py"; Local = "sec_scripts/__init__.py" },
    @{ Remote = "scripts/numerar_contratos_indiretas.py"; Local = "sec_scripts/numerar_contratos_indiretas.py" },
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

$pyMod = Join-Path $Root "sec_scripts\numerar_contratos_indiretas.py"
$txt = Get-Content -LiteralPath $pyMod -Raw
if ($txt -match "from scripts\.gerar_fluxos import") {
    throw "Download trouxe versao ANTIGA (import scripts.*). Rode o baixar de novo."
}
if ($txt -notmatch "_load_gerar_fluxos") {
    throw "Download incompleto: falta _load_gerar_fluxos em numerar_contratos_indiretas.py"
}
Write-Host "Versao numerar_contratos_indiretas.py: OK"

$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}
if (-not $py -or -not (Test-Path -LiteralPath $py)) {
    throw "python.exe nao encontrado em $Root"
}
Write-Host "Python: $py"

Write-Host "Instalando pandas/openpyxl/requests (se preciso)..."
& $py -m pip install "pandas>=2.0" "openpyxl>=3.1" "requests>=2.28"
if ($LASTEXITCODE -ne 0) { throw "pip install falhou" }

$need = @(
    "numerar_contratos_indiretas.bat",
    "numerar_contratos_indiretas.py",
    "sec_scripts\__init__.py",
    "sec_scripts\numerar_contratos_indiretas.py",
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
if (-not (Test-Path -LiteralPath (Join-Path $Root "dados"))) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root "dados") | Out-Null
}

Write-Host ""
Write-Host "OK. Proximo passo:"
Write-Host "  1) Coloque BNDES INDIRETAS AAAA.xlsx em .\dados\"
Write-Host "  2) Rode: .\numerar_contratos_indiretas.bat"
Write-Host "  Saida: .\saida\BNDES_INDIRETAS_NUMERADOS.xlsx"
