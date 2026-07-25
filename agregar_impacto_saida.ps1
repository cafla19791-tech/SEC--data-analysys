# Baixa o agregador streaming e gera impacto fiscal / resumo por agente
# a partir dos fluxos_*.csv ja gerados na pasta saida.
#
# Uso (no WinPython ContAgil):
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\agregar_impacto_saida.ps1

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$b = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/agregar-impacto-streaming-f342"
New-Item -ItemType Directory -Force -Path (Join-Path $Root "scripts") | Out-Null

$files = @(
    "scripts/__init__.py",
    "scripts/gerar_fluxos.py",
    "scripts/impacto_fiscal_por_ano.py",
    "scripts/agregar_impacto_fluxos.py"
)

Write-Host "Baixando agregador streaming para $Root ..."
foreach ($rel in $files) {
    $url = "$b/$rel"
    $out = Join-Path $Root ($rel -replace "/", "\")
    Write-Host "  $rel"
    Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing
}

$pyCandidates = @(
    (Join-Path $Root "python.exe"),
    (Join-Path $Root "python\python.exe")
)
$python = $null
foreach ($c in $pyCandidates) {
    if (Test-Path -LiteralPath $c) { $python = $c; break }
}
if (-not $python) {
    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($cmd) { $python = $cmd.Source }
}
if (-not $python) {
    throw "python.exe nao encontrado em $Root"
}

$saida = Join-Path $Root "saida"
if (-not (Test-Path -LiteralPath $saida)) {
    throw "Pasta saida nao encontrada: $saida"
}

$csvs = Get-ChildItem -LiteralPath $saida -Filter "fluxos_*.csv" | Where-Object {
    $_.Name -notmatch "diario"
}
if (-not $csvs) {
    throw "Nenhum fluxos_*.csv em $saida. Rode contagil_fluxos.py primeiro."
}

Write-Host ""
Write-Host "CSV(s) encontrados:"
$csvs | ForEach-Object {
    Write-Host ("  {0}  ({1:N1} MB)" -f $_.Name, ($_.Length / 1MB))
}
Write-Host ""
Write-Host "Agregando impacto (streaming, modo=coluna)..."
Write-Host "Isso pode levar varios minutos com ~70M parcelas."
Write-Host ""

& $python (Join-Path $Root "scripts\agregar_impacto_fluxos.py") --pasta $saida --modo coluna
if ($LASTEXITCODE -ne 0) {
    throw "agregar_impacto_fluxos.py falhou (exit $LASTEXITCODE)"
}

Write-Host ""
Write-Host "Concluido. Abra na pasta saida:"
Write-Host "  - resumo_impacto_bndes.xlsx"
Write-Host "  - impacto_fiscal_por_ano.xlsx"
Write-Host "  - resumo_por_agente.xlsx"
