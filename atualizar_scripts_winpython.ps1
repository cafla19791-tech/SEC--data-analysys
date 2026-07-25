# Baixa os .py corretos do GitHub para a pasta WinPython ContAgil.
# Uso (PowerShell), na pasta winpython:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\atualizar_scripts_winpython.ps1

$ErrorActionPreference = "Stop"

$Branch = "cursor/normalizar-colunas-6f97"
$Base = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Branch"

$Root = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
Set-Location -LiteralPath $Root

Write-Host "============================================================"
Write-Host "Atualizando scripts ContAgil a partir do GitHub"
Write-Host "Pasta: $Root"
Write-Host "Branch: $Branch"
Write-Host "============================================================"

$scriptsDir = Join-Path $Root "scripts"
if (-not (Test-Path -LiteralPath $scriptsDir)) {
    New-Item -ItemType Directory -Path $scriptsDir | Out-Null
}

$broken = Join-Path $scriptsDir "contagil_fluxos.py"
if (Test-Path -LiteralPath $broken) {
    $text = Get-Content -LiteralPath $broken -Raw -ErrorAction SilentlyContinue
    if ($text -match "REM ContAgil|@echo off") {
        $bak = Join-Path $scriptsDir "contagil_fluxos.py.bak_errado"
        Copy-Item -LiteralPath $broken -Destination $bak -Force
        Write-Host "[AVISO] scripts\contagil_fluxos.py misturado com .bat - backup: $bak"
    }
}

$files = @(
    @{ Rel = "scripts/__init__.py"; Out = (Join-Path $scriptsDir "__init__.py") },
    @{ Rel = "scripts/contagil_fluxos.py"; Out = (Join-Path $scriptsDir "contagil_fluxos.py") },
    @{ Rel = "scripts/contagil_fluxos_seguro.py"; Out = (Join-Path $scriptsDir "contagil_fluxos_seguro.py") },
    @{ Rel = "scripts/gerar_fluxos.py"; Out = (Join-Path $scriptsDir "gerar_fluxos.py") },
    @{ Rel = "contagil_fluxos.py"; Out = (Join-Path $Root "contagil_fluxos.py") },
    @{ Rel = "contagil_fluxos_bndes.bat"; Out = (Join-Path $Root "contagil_fluxos_bndes.bat") }
)

Write-Host ""
Write-Host "Baixando arquivos Python..."
foreach ($f in $files) {
    $url = "$Base/$($f.Rel)"
    Write-Host "  -> $($f.Rel)"
    Write-Host "     $url"
    Invoke-WebRequest -Uri $url -OutFile $f.Out -UseBasicParsing
}

$mainPy = Join-Path $scriptsDir "contagil_fluxos.py"
$head = Get-Content -LiteralPath $mainPy -TotalCount 3 -ErrorAction Stop
$joined = ($head -join "`n")
if ($joined -notmatch "#!/usr/bin/env python") {
    throw "Download invalido: scripts\contagil_fluxos.py nao comeca com shebang Python."
}
$all = Get-Content -LiteralPath $mainPy -Raw
if ($all -match "REM ContAgil") {
    throw "O .py ainda contem linhas REM. Abortando."
}

Write-Host ""
Write-Host "OK: scripts\contagil_fluxos.py e Python valido."
Write-Host ""
Write-Host "Proximo passo - cole NO CMD (uma linha):"
Write-Host ""
Write-Host ('python scripts\contagil_fluxos.py --massa-dados "{0}\dados" --pasta-saida "{0}\saida" --arquivo-fatores "{0}\fator_acumulado_SELIC_TJLP_TLP.xlsx"' -f $Root)
Write-Host ""
Write-Host "============================================================"
