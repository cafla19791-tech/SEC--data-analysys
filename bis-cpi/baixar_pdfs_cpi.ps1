# Baixa PDFs CPI gerados no GitHub para winpython\pdf
param(
    [string]$Branch = "cursor/bis-long-cpi-reports-b311"
)

$ErrorActionPreference = "Stop"
$root = Get-Location
$pdfDir = Join-Path $root "pdf"
New-Item -ItemType Directory -Force -Path $pdfDir | Out-Null

$base = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Branch/output/cpi/pdf"
$files = @(
    "cpi_mensal_por_pais.pdf",
    "cpi_inflacao_acumulada_periodos.pdf"
)

foreach ($f in $files) {
    $url = "$base/$f"
    $out = Join-Path $pdfDir $f
    Write-Host "Baixando $f ..."
    try {
        Invoke-WebRequest $url -OutFile $out -Headers @{ "Cache-Control" = "no-cache" }
        Write-Host "  -> $out"
    } catch {
        Write-Warning "Falha em $f : $_"
    }
}
Write-Host "PDFs em $pdfDir"
