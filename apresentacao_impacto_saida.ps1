# Consolida Excel de apresentacao a partir dos resumos ja gerados na saida.
# Prefere impacto_fiscal_por_ano.csv + resumo_por_agente.csv do agregador.
#
# Usa pasta sec_scripts (nao Scripts) — evita colisao com WinPython\Scripts.
#
# Cole no PowerShell (ja dentro da pasta winpython):
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/apresentacao-impacto-f342/apresentacao_impacto_saida.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile apresentacao_impacto_saida.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\apresentacao_impacto_saida.ps1

param(
    [string]$Ref = "cursor/apresentacao-impacto-f342",
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
Write-Host " Excel de apresentacao — impacto BNDES Indiretas"
Write-Host " Pasta: $Root"
Write-Host " Ref: $Ref | Cache-bust: $bust"
Write-Host "============================================================"

$sec = Join-Path $Root "sec_scripts"
New-Item -ItemType Directory -Force -Path $sec | Out-Null

$uri = "$baseRaw/scripts/apresentacao_impacto_bndes.py?v=$bust"
$outPy = Join-Path $sec "apresentacao_impacto_bndes.py"
Write-Host "  scripts/apresentacao_impacto_bndes.py -> sec_scripts\apresentacao_impacto_bndes.py"
Invoke-WebRequest -Uri $uri -OutFile $outPy -Headers $headers -UseBasicParsing

$txt = Get-Content -LiteralPath $outPy -Raw
if ($txt -notmatch "apresentacao-impacto-bndes-20260816a") {
    throw "Download incompleto/antigo: falta MARKER 20260816a. Rode o baixar de novo."
}
Write-Host "Versao apresentacao_impacto_bndes.py: OK"

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

$saida = Join-Path $Root "saida"
if (-not (Test-Path -LiteralPath $saida)) {
    throw "Pasta saida nao encontrada: $saida"
}

$ano = Join-Path $saida "impacto_fiscal_por_ano.csv"
$ag = Join-Path $saida "resumo_por_agente.csv"
if (-not (Test-Path -LiteralPath $ano)) {
    throw "Falta impacto_fiscal_por_ano.csv — rode agregar_impacto_saida.ps1 primeiro."
}
if (-not (Test-Path -LiteralPath $ag)) {
    throw "Falta resumo_por_agente.csv — rode agregar_impacto_saida.ps1 primeiro."
}

Write-Host ""
Write-Host "Gerando workbook de apresentacao..."
& $py $outPy --pasta $saida
if ($LASTEXITCODE -ne 0) {
    throw "apresentacao_impacto_bndes.py falhou (exit $LASTEXITCODE)"
}

$xlsx = Join-Path $saida "APRESENTACAO_IMPACTO_BNDES_INDIRETAS.xlsx"
Write-Host ""
Write-Host "Concluido:"
Write-Host "  $xlsx"
Write-Host "Abas: Capa | Sumario | Por_Ano | Top_20_Agentes | Por_Agente | Notas"
