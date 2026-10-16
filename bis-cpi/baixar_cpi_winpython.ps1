# Instala bis-cpi no ContAgil WinPython e opcionalmente baixa o CSV do BIS.
param(
    [switch]$DownloadCsv,
    [string]$Branch = "cursor/bis-long-cpi-reports-b311"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not (Test-Path (Join-Path $root "bis_cpi_cli.bat"))) {
    $root = Get-Location
}

$base = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Branch/bis-cpi"
$files = @(
    "pyproject.toml",
    "bis_cpi_cli.bat",
    "src/bis_cpi/__init__.py",
    "src/bis_cpi/area_names.py",
    "src/bis_cpi/loaders.py",
    "src/bis_cpi/excel_format.py",
    "src/bis_cpi/excel_mensal.py",
    "src/bis_cpi/excel_periodos.py",
    "src/bis_cpi/pdf_export.py",
    "src/bis_cpi/cli.py"
)

$dest = Join-Path $root "bis-cpi"
New-Item -ItemType Directory -Force -Path (Join-Path $dest "src\bis_cpi") | Out-Null

foreach ($rel in $files) {
    $url = "$base/$rel" + "?v=1"
    $out = Join-Path $dest ($rel -replace "/", "\")
    $parent = Split-Path -Parent $out
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
    Write-Host "Baixando $rel"
    Invoke-WebRequest $url -OutFile $out -Headers @{ "Cache-Control" = "no-cache" }
}

python -m pip install -q openpyxl
Write-Host "OK: $dest"

if ($DownloadCsv) {
    Push-Location $root
    $env:PYTHONPATH = (Join-Path $dest "src")
    python -m bis_cpi.cli download --dir .
    Pop-Location
}
