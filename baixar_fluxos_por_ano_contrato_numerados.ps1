# Baixa fluxos por ano a partir de BNDES_INDIRETAS_NUMERADOS.xlsx no ContAgil WinPython.
# Usa pasta sec_scripts (evita colisao com WinPython\Scripts).
#
# Cole no PowerShell (ja dentro da pasta winpython):
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/fluxos-ano-retomar-f342/baixar_fluxos_por_ano_contrato_numerados.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_fluxos_ano.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_fluxos_ano.ps1
#   .\fluxos_por_ano_contrato_numerados.bat

param(
    [string]$Ref = "cursor/fluxos-ano-retomar-f342",
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
Write-Host " Baixando fluxos por ano (NUMERADOS) -> $Root"
Write-Host " Pasta de codigo: sec_scripts (nao Scripts)"
Write-Host " Ref: $Ref | Cache-bust: $bust"
Write-Host "============================================================"

$map = @(
    @{ Remote = "fluxos_por_ano_contrato_numerados.bat"; Local = "fluxos_por_ano_contrato_numerados.bat" },
    @{ Remote = "fluxos_por_ano_contrato_numerados.py"; Local = "fluxos_por_ano_contrato_numerados.py" },
    @{ Remote = "scripts/__init__.py"; Local = "sec_scripts/__init__.py" },
    @{ Remote = "scripts/fluxos_por_ano_contrato_numerados.py"; Local = "sec_scripts/fluxos_por_ano_contrato_numerados.py" },
    @{ Remote = "scripts/gerar_fluxos.py"; Local = "sec_scripts/gerar_fluxos.py" },
    @{ Remote = "scripts/contagil_fluxos_seguro.py"; Local = "sec_scripts/contagil_fluxos_seguro.py" }
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

$pyMod = Join-Path $Root "sec_scripts\fluxos_por_ano_contrato_numerados.py"
$txt = Get-Content -LiteralPath $pyMod -Raw
if ($txt -notmatch "fluxos-por-ano-contrato-numerados-20260814b") {
    throw "Download incompleto/antigo: falta MARKER 20260814b (fix cp1252). Rode o baixar de novo."
}
if ($txt -match "retomar") {
    Write-Host "Versao fluxos_por_ano_contrato_numerados.py: OK (retomar/CSV-first)"
} else {
    throw "Download trouxe versao ANTIGA (sem retomar). Rode o baixar de novo."
}
$gf = Join-Path $Root "sec_scripts\gerar_fluxos.py"
$gfTxt = Get-Content -LiteralPath $gf -Raw
if ($gfTxt -match [char]0x1F680) {
    throw "gerar_fluxos.py ainda tem emoji foguete - download antigo. Rode o baixar de novo."
}
Write-Host "Versao gerar_fluxos.py: OK (sem emoji no console)"

$bat = Join-Path $Root "fluxos_por_ano_contrato_numerados.bat"
$batTxt = Get-Content -LiteralPath $bat -Raw
if ($batTxt -match "chcp\s+65001") {
    throw "fluxos_por_ano_contrato_numerados.bat ainda tem chcp 65001 (causa erro 'M' no cmd). Rode o baixar de novo."
}
Write-Host "Versao fluxos_por_ano_contrato_numerados.bat: OK (sem chcp 65001)"

$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}
if (-not $py -or -not (Test-Path -LiteralPath $py)) {
    throw "python.exe nao encontrado em $Root"
}
Write-Host "Python: $py"

Write-Host "Instalando pandas/openpyxl (se preciso)..."
& $py -m pip install "pandas>=2.0" "openpyxl>=3.1"
if ($LASTEXITCODE -ne 0) { throw "pip install falhou" }

$need = @(
    "fluxos_por_ano_contrato_numerados.bat",
    "fluxos_por_ano_contrato_numerados.py",
    "sec_scripts\__init__.py",
    "sec_scripts\fluxos_por_ano_contrato_numerados.py",
    "sec_scripts\gerar_fluxos.py",
    "sec_scripts\contagil_fluxos_seguro.py"
)
foreach ($rel in $need) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path -LiteralPath $p)) { throw "Arquivo faltando apos download: $rel" }
    if ((Get-Item -LiteralPath $p).Length -lt 50) { throw "Arquivo suspeito (muito pequeno): $rel" }
}
Write-Host "Checagem de arquivos: OK"

$saida = Join-Path $Root "saida"
if (-not (Test-Path -LiteralPath $saida)) {
    New-Item -ItemType Directory -Force -Path $saida | Out-Null
}

$numerados = Join-Path $saida "BNDES_INDIRETAS_NUMERADOS.xlsx"
Write-Host ""
Write-Host "OK. Proximo passo:"
if (Test-Path -LiteralPath $numerados) {
    Write-Host "  Achou: $numerados"
    Write-Host "  Rode: .\fluxos_por_ano_contrato_numerados.bat"
} else {
    Write-Host "  Falta saida\BNDES_INDIRETAS_NUMERADOS.xlsx"
    Write-Host "  Rode antes: .\numerar_contratos_indiretas.bat"
    Write-Host "  Depois: .\fluxos_por_ano_contrato_numerados.bat"
}
Write-Host "  Saidas: saida\fluxos_por_ano_contrato\YYYY.csv"
Write-Host "          saida\FLUXOS_BNDES_INDIRETAS_POR_ANO_CONTRATO.xlsx"
