# Baixa os PDFs CBPOL pre-gerados para a pasta atual (ContAgil WinPython).
# ASCII-only. Cole no PowerShell DENTRO de:
#   C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython
#
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/bis-cbpol-mcp-41ca/bis-mcp/baixar_pdfs_cbpol.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_pdfs_cbpol.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_pdfs_cbpol.ps1

param(
    [string]$Ref = "cursor/bis-cbpol-mcp-41ca",
    [string]$OutDir = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$dest = if ($OutDir) { $OutDir } else { Join-Path $Root "pdf" }
if (-not (Test-Path -LiteralPath $dest)) {
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
}

$base = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref/output/pdf"
$headers = @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" }

$files = @(
    "cbpol_taxas_acumuladas_periodos.pdf",
    "cbpol_taxas_acumuladas_periodos_mensal.pdf",
    "cbpol_taxas_mensais_compostas.pdf",
    "cbpol_taxas_diarias_compostas_desde_2000.pdf",
    "cbpol_taxas_diarias_compostas.pdf",
    "cbpol_taxas_diarias_compostas_indice.pdf",
    "cbpol_taxas_diarias_compostas_desde_2000_indice.pdf"
)

Write-Host "============================================================"
Write-Host " Baixando PDFs CBPOL -> $dest"
Write-Host "============================================================"

foreach ($name in $files) {
    $url = "$base/$name`?v=1"
    $out = Join-Path $dest $name
    Write-Host "  $name"
    Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing -Headers $headers
}

Write-Host ""
Write-Host "OK. PDFs em:"
Write-Host "  $dest"
Get-ChildItem -LiteralPath $dest -Filter *.pdf | ForEach-Object {
    Write-Host ("  {0,10:N0}  {1}" -f $_.Length, $_.Name)
}
