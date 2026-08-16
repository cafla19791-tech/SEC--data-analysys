# Converte saida\fluxos_por_ano_contrato\YYYY.csv em YYYY.xlsx fatiado
# (varias abas de ate ~1M linhas - limite do Excel).
#
# Usa pasta sec_scripts (nao Scripts).
#
# Cole no PowerShell (ja dentro da pasta winpython):
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/csv-para-xlsx-fatiado-f342/csv_para_xlsx_fatiado_saida.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile csv_para_xlsx_fatiado_saida.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\csv_para_xlsx_fatiado_saida.ps1
#
# Opcoes:
#   .\csv_para_xlsx_fatiado_saida.ps1 -Ano 2011
#   .\csv_para_xlsx_fatiado_saida.ps1 -Retomar

param(
    [string]$Ref = "cursor/csv-para-xlsx-fatiado-f342",
    [string]$Python = "",
    [int]$Ano = 0,
    [switch]$Retomar
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
Write-Host " CSV -> XLSX fatiado (fluxos por ano de contrato)"
Write-Host " Pasta: $Root"
Write-Host " Ref: $Ref | Cache-bust: $bust"
Write-Host "============================================================"
Write-Host "AVISO: pode levar HORAS e ocupar dezenas de GB em disco."
Write-Host "       O CSV continua sendo a fonte completa."
Write-Host ""

$sec = Join-Path $Root "sec_scripts"
New-Item -ItemType Directory -Force -Path $sec | Out-Null

$uri = "$baseRaw/scripts/csv_fluxos_para_xlsx_fatiado.py?v=$bust"
$outPy = Join-Path $sec "csv_fluxos_para_xlsx_fatiado.py"
Write-Host "  scripts/csv_fluxos_para_xlsx_fatiado.py -> sec_scripts\csv_fluxos_para_xlsx_fatiado.py"
Invoke-WebRequest -Uri $uri -OutFile $outPy -Headers $headers -UseBasicParsing

$txt = Get-Content -LiteralPath $outPy -Raw
if ($txt -notmatch "csv-para-xlsx-fatiado-20260816a") {
    throw "Download incompleto/antigo: falta MARKER 20260816a. Rode o baixar de novo."
}
Write-Host "Versao: OK"

$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}
if (-not $py -or -not (Test-Path -LiteralPath $py)) {
    throw "python.exe nao encontrado em $Root"
}
Write-Host "Python: $py"

Write-Host "Instalando openpyxl (se preciso)..."
& $py -m pip install "openpyxl>=3.1"
if ($LASTEXITCODE -ne 0) { throw "pip install falhou" }

$pasta = Join-Path $Root "saida\fluxos_por_ano_contrato"
if (-not (Test-Path -LiteralPath $pasta)) {
    throw "Pasta nao encontrada: $pasta — rode fluxos_por_ano_contrato_numerados.bat primeiro."
}

$argsPy = @($outPy, "--pasta", $pasta)
if ($Ano -gt 0) { $argsPy += @("--ano", "$Ano") }
if ($Retomar) { $argsPy += "--retomar" }

Write-Host ""
Write-Host "Convertendo..."
& $py @argsPy
if ($LASTEXITCODE -ne 0) {
    throw "csv_fluxos_para_xlsx_fatiado.py falhou (exit $LASTEXITCODE)"
}

Write-Host ""
Write-Host "Concluido. XLSX por ano em:"
Write-Host "  $pasta\YYYY.xlsx"
Write-Host "Abas: YYYY_p01, YYYY_p02, ... (ate ~1M linhas de dados cada)"
