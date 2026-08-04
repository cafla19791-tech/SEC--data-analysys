# Baixa o pacote nyse-mcp para a pasta atual (ex.: ContAgil WinPython)
# e roda o setup da CLI. Nao precisa clonar o repo inteiro.
# ASCII-only (PowerShell 5 / Invoke-WebRequest).
#
# Cole no PowerShell DENTRO de:
#   C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython
#
#   $b="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/nyse-mcp-winpython-cli-f342"
#   Invoke-WebRequest "$b/nyse-mcp/baixar_nyse_mcp_winpython.ps1" -OutFile baixar_nyse_mcp_winpython.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_nyse_mcp_winpython.ps1 -Symbol JPM

param(
    [string]$Symbol = "JPM",
    [string]$Branch = "cursor/nyse-mcp-winpython-cli-f342",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root

$base = "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/$Branch"
$dest = Join-Path $Root "nyse-mcp"

Write-Host "============================================================"
Write-Host " Baixando nyse-mcp -> $dest"
Write-Host " Branch: $Branch"
Write-Host "============================================================"

$files = @(
    "nyse-mcp/pyproject.toml",
    "nyse-mcp/README.md",
    "nyse-mcp/setup_e_rodar_cli.ps1",
    "nyse-mcp/nyse_mcp_cli.bat",
    "nyse-mcp/run_mcp.sh",
    "nyse-mcp/.env.example",
    "nyse-mcp/CADASTRO_MCP_CLOUD.md",
    "nyse-mcp/CLOUD_SETUP.md",
    "nyse-mcp/src/nyse_mcp/__init__.py",
    "nyse-mcp/src/nyse_mcp/cli.py",
    "nyse-mcp/src/nyse_mcp/providers.py",
    "nyse-mcp/src/nyse_mcp/server.py"
)

function Save-RawFile {
    param(
        [string]$Url,
        [string]$OutPath
    )
    $dir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    # .ps1/.bat: grava como texto UTF-8 com BOM (PowerShell 5 le melhor)
    if ($OutPath -match '\.(ps1|bat|cmd)$') {
        $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing
        $text = $resp.Content
        $utf8Bom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($OutPath, $text, $utf8Bom)
    } else {
        Invoke-WebRequest -Uri $Url -OutFile $OutPath -UseBasicParsing
    }
}

foreach ($rel in $files) {
    $url = "$base/$rel"
    $relLocal = $rel -replace "^nyse-mcp/", "nyse-mcp\"
    $out = Join-Path $Root ($relLocal -replace "/", "\")
    Write-Host "  $rel"
    Save-RawFile -Url $url -OutPath $out
}

$py = $Python
if (-not $py) {
    $localPy = Join-Path $Root "python.exe"
    if (Test-Path -LiteralPath $localPy) { $py = $localPy }
}

Write-Host ""
Write-Host "Rodando setup da CLI..."
$setup = Join-Path $dest "setup_e_rodar_cli.ps1"
if ($py) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $setup -Symbol $Symbol -Python $py
} else {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $setup -Symbol $Symbol
}
if ($LASTEXITCODE -ne 0) {
    throw "setup_e_rodar_cli.ps1 falhou (exit $LASTEXITCODE)"
}

Write-Host ""
Write-Host "Pronto. A partir de agora, nesta pasta WinPython:"
Write-Host "  cd nyse-mcp"
Write-Host "  .\nyse_mcp_cli.bat quote JPM"
Write-Host "  .\nyse_mcp_cli.bat history XOM --period 1y"
Write-Host "  .\nyse_mcp_cli.bat search `"Coca Cola`""
Write-Host "  .\nyse_mcp_cli.bat status"
