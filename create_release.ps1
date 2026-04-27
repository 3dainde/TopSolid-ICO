# Deployment script for Top'ICO v1.0.2
param([string]$Version = "1.0.2")

$ReleaseDir = "TopICO_v$Version"
$OutputPath = "$ReleaseDir.zip"

Write-Host "Creating Top'ICO v$Version release package..." -ForegroundColor Cyan

# Create release directory
if (Test-Path $ReleaseDir) { Remove-Item $ReleaseDir -Recurse -Force }
New-Item -ItemType Directory -Path $ReleaseDir | Out-Null

# Copy files
Write-Host "Copying files..." -ForegroundColor Gray

$exePath = "TopICO\bin\Release\net8.0-windows\win-x64\publish\TopICO.exe"
if (Test-Path $exePath) {
    Copy-Item $exePath -Destination "$ReleaseDir\TopICO.exe"
    Write-Host "✓ TopICO.exe copied"
}

$scripts = @("fix_icons.py", "fix_icons_simple.ps1", "fix_icons.ps1", "README.md")
foreach ($script in $scripts) {
    if (Test-Path $script) {
        Copy-Item $script -Destination $ReleaseDir
        Write-Host "✓ $script copied"
    }
}

# Create archive
Write-Host "Creating archive..." -ForegroundColor Gray
Compress-Archive -Path "$ReleaseDir\*" -DestinationPath $OutputPath -Force
Write-Host "✓ Archive created: $OutputPath"

# Show contents
Write-Host "`nRelease contents:" -ForegroundColor Green
Get-ChildItem $ReleaseDir | ForEach-Object { Write-Host "  - $($_.Name)" }

Write-Host "`nPackage ready for GitHub release!" -ForegroundColor Green