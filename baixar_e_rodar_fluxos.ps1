# Baixa os 4 .py obrigatorios e roda o calculo ContAgil.
# Uso: clique direito > Executar com PowerShell
#  ou: powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_e_rodar_fluxos.ps1

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
Set-Location -LiteralPath $Root

$b = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/normalizar-colunas-6f97"
New-Item -ItemType Directory -Force -Path (Join-Path $Root "scripts") | Out-Null

$files = @(
    "scripts/__init__.py",
    "scripts/contagil_fluxos.py",
    "scripts/contagil_fluxos_seguro.py",
    "scripts/gerar_fluxos.py"
)

Write-Host "Baixando para $Root ..."
foreach ($rel in $files) {
    $url = "$b/$rel"
    $out = Join-Path $Root ($rel -replace "/", "\")
    Write-Host "  $rel"
    Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing
}

Write-Host ""
Write-Host "Arquivos em scripts\:"
Get-ChildItem (Join-Path $Root "scripts\*.py") | ForEach-Object {
    Write-Host ("  {0,-32} {1,10} bytes" -f $_.Name, $_.Length)
}

$gf = Join-Path $Root "scripts\gerar_fluxos.py"
if (-not (Test-Path $gf) -or (Get-Item $gf).Length -lt 10000) {
    throw "gerar_fluxos.py ausente ou muito pequeno. Download falhou."
}

$head = Get-Content (Join-Path $Root "scripts\contagil_fluxos.py") -TotalCount 5
if (($head -join "`n") -notmatch "importlib") {
    Write-Host "AVISO: contagil_fluxos.py pode estar desatualizado (sem importlib)."
}

$dados = Join-Path $Root "dados"
$saida = Join-Path $Root "saida"
$fatores = Join-Path $Root "fator_acumulado_SELIC_TJLP_TLP.xlsx"

Write-Host ""
Write-Host "Executando calculo..."
& python (Join-Path $Root "scripts\contagil_fluxos.py") `
    --massa-dados $dados `
    --pasta-saida $saida `
    --arquivo-fatores $fatores

exit $LASTEXITCODE
