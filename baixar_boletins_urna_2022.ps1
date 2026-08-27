# Copia o script de Boletins de Urna 2022 para o ContAgil WinPython.
# Cole no PowerShell (ja dentro da pasta winpython):
#
#   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
#   $u="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/tse-boletins-urna-209b/baixar_boletins_urna_2022.ps1"
#   Invoke-WebRequest "$u`?v=1" -OutFile baixar_boletins_urna_2022.ps1 -Headers @{"Cache-Control"="no-cache"}
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_boletins_urna_2022.ps1
#   .\baixar_boletins_urna_2022.bat

param(
    [string]$Ref = "cursor/tse-boletins-urna-209b",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$bust = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$baseRaw = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Ref"
$headers = @{
    "Cache-Control" = "no-cache"
    "Pragma"        = "no-cache"
}

Write-Host "============================================================"
Write-Host " Baixando boletins de urna 2022 -> $Root"
Write-Host " Pasta de codigo: sec_scripts (nao Scripts)"
Write-Host "============================================================"

$map = @(
    @{ Remote = "baixar_boletins_urna_2022.bat"; Local = "baixar_boletins_urna_2022.bat" },
    @{ Remote = "baixar_boletins_urna_2022.py"; Local = "baixar_boletins_urna_2022.py" },
    @{ Remote = "baixar_zips_urna_curl.bat"; Local = "baixar_zips_urna_curl.bat" },
    @{ Remote = "baixar_zips_urna.ps1"; Local = "baixar_zips_urna.ps1" },
    @{ Remote = "baixar_resultado_urna_github.bat"; Local = "baixar_resultado_urna_github.bat" },
    @{ Remote = "scripts/baixar_boletins_urna_2022.py"; Local = "sec_scripts/baixar_boletins_urna_2022.py" }
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

$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}
if ($py -and (Test-Path -LiteralPath $py)) {
    Write-Host "Python: $py"
    Write-Host "Instalando pandas/requests (se preciso)..."
    & $py -m pip install "pandas>=2.0" "requests>=2.28"
    if ($LASTEXITCODE -ne 0) { throw "pip install falhou" }
}

Write-Host ""
Write-Host "OK. Na RFB, de duplo-clique em baixar_resultado_urna_github.bat"
Write-Host "  ZIPs  : $Root\dados\tse2022\raw"
Write-Host "  Saida : $Root\saida\tse2022\urnas_2t_presidente.csv"
Write-Host ""
