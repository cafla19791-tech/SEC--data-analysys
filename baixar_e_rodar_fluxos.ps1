# Baixa os 4 .py obrigatorios e roda o calculo ContAgil com o Python do WinPython.
# Uso:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_e_rodar_fluxos.ps1

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
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

$content = Get-Content (Join-Path $Root "scripts\contagil_fluxos.py") -Raw
if ($content -notmatch "importlib-20260725") {
    Write-Host "AVISO: contagil_fluxos.py sem marcador importlib-20260725."
}
if ($content -notmatch "progresso-lotes") {
    Write-Host "AVISO: baixe de novo - falta progresso/lotes para massas grandes."
}

function Find-WinPython {
    param([string]$RootDir)

    $candidates = New-Object System.Collections.Generic.List[string]

    # Locais tipicos do WinPython / ContAgil
    @(
        (Join-Path $RootDir "python.exe"),
        (Join-Path $RootDir "python\python.exe"),
        (Join-Path $RootDir "python-3.12.5.amd64\python.exe"),
        (Join-Path $RootDir "python-3.11.9.amd64\python.exe"),
        (Join-Path $RootDir "python-3.10.11.amd64\python.exe"),
        (Join-Path (Split-Path $RootDir -Parent) "python.exe"),
        (Join-Path (Split-Path $RootDir -Parent) "python\python.exe")
    ) | ForEach-Object { if ($_ -and (Test-Path -LiteralPath $_)) { [void]$candidates.Add($_) } }

    # Busca rasa (profundidade 2) por python.exe dentro do winpython
    Get-ChildItem -LiteralPath $RootDir -Filter python.exe -File -ErrorAction SilentlyContinue |
        ForEach-Object { [void]$candidates.Add($_.FullName) }
    Get-ChildItem -LiteralPath $RootDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $p = Join-Path $_.FullName "python.exe"
        if (Test-Path -LiteralPath $p) { [void]$candidates.Add($p) }
        Get-ChildItem -LiteralPath $_.FullName -Filter python.exe -File -Recurse -Depth 1 -ErrorAction SilentlyContinue |
            ForEach-Object { [void]$candidates.Add($_.FullName) }
    }

    # Evita o stub da Microsoft Store (WindowsApps)
    foreach ($c in $candidates | Select-Object -Unique) {
        if ($c -match "WindowsApps") { continue }
        if (Test-Path -LiteralPath $c) { return $c }
    }

    # py launcher
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { return "py" }

    return $null
}

$pyExe = Find-WinPython -RootDir $Root
if (-not $pyExe) {
    Write-Host ""
    Write-Host "ERRO: nao achei python.exe do WinPython/ContAgil."
    Write-Host "Procure manualmente, por exemplo:"
    Write-Host "  Get-ChildItem -Path `"$Root`" -Filter python.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -First 10 FullName"
    Write-Host ""
    Write-Host "Depois rode com o caminho completo, ex.:"
    Write-Host "  & `"$Root\python.exe`" scripts\contagil_fluxos.py --massa-dados `"$Root\dados`" --pasta-saida `"$Root\saida`" --arquivo-fatores `"$Root\fator_acumulado_SELIC_TJLP_TLP.xlsx`""
    exit 1
}

$dados = Join-Path $Root "dados"
$saida = Join-Path $Root "saida"
$fatores = Join-Path $Root "fator_acumulado_SELIC_TJLP_TLP.xlsx"
$scriptPy = Join-Path $Root "scripts\contagil_fluxos.py"

Write-Host ""
Write-Host "Python: $pyExe"
Write-Host "Executando calculo..."

if ($pyExe -eq "py") {
    & py -3 $scriptPy --massa-dados $dados --pasta-saida $saida --arquivo-fatores $fatores
} else {
    & $pyExe $scriptPy --massa-dados $dados --pasta-saida $saida --arquivo-fatores $fatores
}

exit $LASTEXITCODE
