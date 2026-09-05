# Baixa os 28 ZIPs oficiais de BU 2022 usando WinHTTP + proxy da RFB + TLS 1.2.
# O curl.exe/Schannel costuma falhar no Archive.org com SEC_E_ILLEGAL_MESSAGE.
#
# Na pasta winpython:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\baixar_zips_urna.ps1

param(
    [string[]]$Ufs = @(),
    [switch]$SomenteProcessar
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = "SilentlyContinue"

$Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location -LiteralPath $Root
$Raw = Join-Path $Root "dados\tse2022\raw"
$Saida = Join-Path $Root "saida\tse2022"
New-Item -ItemType Directory -Force -Path $Raw | Out-Null
New-Item -ItemType Directory -Force -Path $Saida | Out-Null

if (-not $Ufs -or $Ufs.Count -eq 0) {
    $Ufs = @("AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO","ZZ")
}

function Get-WebClient([string]$UserAgent) {
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent", $UserAgent)
    try {
        $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        $wc.Proxy = $proxy
    } catch {}
    return $wc
}

function Save-Url([string]$Url, [string]$Dest, [string]$UserAgent) {
    $part = "$Dest.part"
    if (Test-Path -LiteralPath $part) { Remove-Item -LiteralPath $part -Force }
    $wc = Get-WebClient $UserAgent
    try {
        $wc.DownloadFile($Url, $part)
        if ((Get-Item -LiteralPath $part).Length -lt 1000) {
            throw "arquivo vazio"
        }
        Move-Item -LiteralPath $part -Destination $Dest -Force
        return $true
    } catch {
        if (Test-Path -LiteralPath $part) { Remove-Item -LiteralPath $part -Force }
        return $false
    } finally {
        $wc.Dispose()
    }
}

$Tse = "https://cdn.tse.jus.br/estatistica/sead/eleicoes/eleicoes2022/buweb"
$Ia = "https://web.archive.org/web/20221108000702id_/https://cdn.tse.jus.br/estatistica/sead/eleicoes/eleicoes2022/buweb"
$UaTse = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
$UaIa = "ContAgil-TSE-BU/1.0"

Write-Host "ZIPs : $Raw"
if (-not $SomenteProcessar) {
    foreach ($uf in $Ufs) {
        $nome = "bweb_2t_${uf}_311020221535.zip"
        $dest = Join-Path $Raw $nome
        if ((Test-Path -LiteralPath $dest) -and (Get-Item -LiteralPath $dest).Length -gt 1000) {
            Write-Host "[ok] $uf ja existe"
            continue
        }
        Write-Host "[TSE] $uf"
        if (Save-Url "$Tse/$nome" $dest $UaTse) {
            Write-Host "[ok] $uf via TSE"
            continue
        }
        Write-Host "[IA] $uf"
        if (Save-Url "$Ia/$nome" $dest $UaIa) {
            Write-Host "[ok] $uf via Archive.org"
            continue
        }
        Write-Host "[ERRO] $uf"
    }
}

$py = Join-Path $Root "python.exe"
if (-not (Test-Path -LiteralPath $py)) { $py = "python" }
$runner = Join-Path $Root "baixar_boletins_urna_2022.py"
if (Test-Path -LiteralPath $runner) {
    Write-Host "Processando..."
    & $py $runner --somente-processar --massa-dados (Join-Path $Root "dados") --pasta-saida (Join-Path $Root "saida")
}
