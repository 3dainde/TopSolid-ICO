# TopSolid Icon Fixer - Windows 10/11 Compatible
# Run: powershell -ExecutionPolicy Bypass -File fix_icons_simple.ps1

function Get-TopSolidPath {
    $paths = @(
        "C:\Missler\V627",
        "C:\Missler\V626",
        "C:\Missler\V625",
        "C:\Program Files\Missler\TopSolid"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

function Get-TopSolidExe {
    param($BasePath)
    $exes = @("top627.exe", "top626.exe", "top625.exe", "TopSolid.exe")
    foreach ($exe in $exes) {
        $full = Join-Path -Path $BasePath -ChildPath "bin" | Join-Path -ChildPath $exe
        if (Test-Path $full) { return $full }
    }
    return $null
}

function Fix-Icon {
    param($ext, $exePath, $index)
    $iconVal = "`"$exePath`",$index"
    $regPath = "HKCU:\Software\Classes\$ext\DefaultIcon"
    
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
    }
    
    try {
        Set-ItemProperty -Path $regPath -Name "(Default)" -Value $iconVal
        Write-Host "[OK] $ext fixed"
        return 1
    }
    catch {
        Write-Host "[FAIL] $ext : $_"
        return 0
    }
}

Write-Host "TopSolid Icon Fixer v2.1"
Write-Host "=========================="
Write-Host ""

$topPath = Get-TopSolidPath
if (-not $topPath) {
    Write-Host "[ERROR] TopSolid not found"
    exit 1
}

Write-Host "Found: $topPath"

$exePath = Get-TopSolidExe $topPath
if (-not $exePath) {
    Write-Host "[ERROR] Executable not found"
    exit 1
}

Write-Host "Executable: $(Split-Path -Path $exePath -Leaf)"
Write-Host ""

$count = 0
$count += Fix-Icon ".top" $exePath 0
$count += Fix-Icon ".dft" $exePath 1
$count += Fix-Icon ".vdx" $exePath 2
$count += Fix-Icon ".prj" $exePath 3

Write-Host ""
Write-Host "Refreshing Explorer..."
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500
Start-Process explorer
Write-Host ""
Write-Host "[DONE] $count files fixed"
