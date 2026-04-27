# Deployment script for Top'ICO v1.0.2
param(
    [string]$Version = "1.0.2"
)

$ReleaseDir = "TopICO_v$Version"
$OutputPath = "$ReleaseDir.zip"

Write-Host "Creating TopICO v$Version release package..." -ForegroundColor Cyan

if (Test-Path $ReleaseDir) {
    Remove-Item $ReleaseDir -Recurse -Force
}

New-Item -ItemType Directory -Path $ReleaseDir | Out-Null

Write-Host "Copying files..." -ForegroundColor Gray

$exePath = "TopICO\bin\Release\net8.0-windows\win-x64\publish\TopICO.exe"
if (Test-Path $exePath) {
    Copy-Item $exePath -Destination "$ReleaseDir\TopICO.exe"
    Write-Host "OK TopICO.exe copied"
}

$filesToCopy = @(
    "fix_icons.py",
    "fix_icons_simple.ps1",
    "fix_icons.ps1",
    "README.md"
)

foreach ($file in $filesToCopy) {
    if (Test-Path $file) {
        Copy-Item $file -Destination $ReleaseDir
        Write-Host "OK $file copied"
    }
}

Write-Host "Creating archive..." -ForegroundColor Gray
if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force
}

Compress-Archive -Path "$ReleaseDir\*" -DestinationPath $OutputPath -Force
Write-Host "OK Archive created: $OutputPath" -ForegroundColor Green

Write-Host ""
Write-Host "Release contents:" -ForegroundColor Green
Get-ChildItem $ReleaseDir | ForEach-Object {
    Write-Host ("  - " + $_.Name)
}

Write-Host ""
Write-Host "Package ready for GitHub release." -ForegroundColor Green