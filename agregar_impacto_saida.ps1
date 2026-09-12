# Baixa o agregador streaming e gera impacto fiscal / resumo por agente.
# Prefere saida\fluxos_por_ano_contrato\YYYY.csv (pipeline NUMERADOS);
# se nao houver, usa fluxos_*.csv (pipeline antigo).
#
# Usa pasta sec_scripts (nao Scripts) — evita colisao com WinPython\Scripts.
#
# Cole no PowerShell (ja dentro da pasta winpython):
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/agregar-ano-contrato-f342/agregar_impacto_saida.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile agregar_impacto_saida.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\agregar_impacto_saida.ps1

param(
    [string]$Ref = "cursor/agregar-ano-contrato-f342",
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
Write-Host " Agregar impacto fiscal (streaming)"
Write-Host " Pasta: $Root"
Write-Host " Ref: $Ref | Cache-bust: $bust"
Write-Host "============================================================"

$sec = Join-Path $Root "sec_scripts"
New-Item -ItemType Directory -Force -Path $sec | Out-Null

$files = @(
    @{ Remote = "scripts/__init__.py"; Local = "sec_scripts\__init__.py" },
    @{ Remote = "scripts/gerar_fluxos.py"; Local = "sec_scripts\gerar_fluxos.py" },
    @{ Remote = "scripts/impacto_fiscal_por_ano.py"; Local = "sec_scripts\impacto_fiscal_por_ano.py" },
    @{ Remote = "scripts/agregar_impacto_fluxos.py"; Local = "sec_scripts\agregar_impacto_fluxos.py" }
)

foreach ($item in $files) {
    $uri = "$baseRaw/$($item.Remote)?v=$bust"
    $out = Join-Path $Root $item.Local
    Write-Host "  $($item.Remote) -> $($item.Local)"
    Invoke-WebRequest -Uri $uri -OutFile $out -Headers $headers -UseBasicParsing
}

$pyMod = Join-Path $Root "sec_scripts\agregar_impacto_fluxos.py"
$txt = Get-Content -LiteralPath $pyMod -Raw
if ($txt -notmatch "agregar-impacto-streaming-20260816a-ano-contrato") {
    throw "Download incompleto/antigo: falta MARKER 20260816a-ano-contrato. Rode o baixar de novo."
}
if ($txt -notmatch "fluxos_por_ano_contrato") {
    throw "Download trouxe versao ANTIGA (sem fluxos_por_ano_contrato). Rode o baixar de novo."
}
Write-Host "Versao agregar_impacto_fluxos.py: OK (ano-contrato)"

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

$porAno = Join-Path $saida "fluxos_por_ano_contrato"
$temPorAno = $false
if (Test-Path -LiteralPath $porAno) {
    $anos = Get-ChildItem -LiteralPath $porAno -Filter "*.csv" | Where-Object {
        $_.BaseName -match '^\d{4}$'
    }
    if ($anos) {
        $temPorAno = $true
        Write-Host ""
        Write-Host "Fonte: fluxos_por_ano_contrato ($($anos.Count) anos)"
        $anos | Sort-Object Name | ForEach-Object {
            Write-Host ("  {0}  ({1:N1} MB)" -f $_.Name, ($_.Length / 1MB))
        }
    }
}

if (-not $temPorAno) {
    $csvs = Get-ChildItem -LiteralPath $saida -Filter "fluxos_*.csv" | Where-Object {
        $_.Name -notmatch "diario"
    }
    if (-not $csvs) {
        throw "Nenhum CSV: rode fluxos_por_ano_contrato_numerados.bat ou contagil_fluxos primeiro."
    }
    Write-Host ""
    Write-Host "Fonte: fluxos_*.csv (pipeline antigo)"
    $csvs | ForEach-Object {
        Write-Host ("  {0}  ({1:N1} MB)" -f $_.Name, ($_.Length / 1MB))
    }
}

Write-Host ""
Write-Host "Agregando impacto (streaming, modo=coluna)..."
Write-Host "Isso pode levar varios minutos com dezenas de milhoes de parcelas."
Write-Host ""

& $py (Join-Path $Root "sec_scripts\agregar_impacto_fluxos.py") --pasta $saida --modo coluna --output-dir $saida
if ($LASTEXITCODE -ne 0) {
    throw "agregar_impacto_fluxos.py falhou (exit $LASTEXITCODE)"
}

Write-Host ""
Write-Host "Concluido. Abra na pasta saida:"
Write-Host "  - resumo_impacto_bndes.xlsx"
Write-Host "  - impacto_fiscal_por_ano.xlsx"
Write-Host "  - resumo_por_agente.xlsx"
