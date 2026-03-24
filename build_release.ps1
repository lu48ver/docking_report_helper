param(
    [string]$Version
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

if (-not $Version) {
    $Version = python -c "from app_meta import APP_VERSION; print(APP_VERSION)"
}

$releaseRoot = Join-Path $repoRoot "artifacts\releases"
$releaseName = "ReportHelper-v$Version-win64"
$zipPath = Join-Path $releaseRoot "$releaseName.zip"
$hashPath = Join-Path $releaseRoot "$releaseName.sha256.txt"

if (Test-Path "build") {
    Remove-Item "build" -Recurse -Force
}
if (Test-Path "dist\ReportHelper") {
    Remove-Item "dist\ReportHelper" -Recurse -Force
}

New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null

python -m PyInstaller "ReportHelper.spec" --noconfirm

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Compress-Archive -Path "dist\ReportHelper\*" -DestinationPath $zipPath -CompressionLevel Optimal

$hash = (Get-FileHash $zipPath -Algorithm SHA256).Hash
"SHA256  $releaseName.zip  $hash" | Set-Content -Path $hashPath -Encoding UTF8

Write-Output "ZIP=$zipPath"
Write-Output "SHA256=$hashPath"
