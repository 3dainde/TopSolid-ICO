# TopSolid Icon Fixer - PowerShell Version (Windows 10/11)
# Run with: powershell -ExecutionPolicy Bypass -File fix_icons_v2.ps1

function Test-AdminRights {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Find-TopSolidPath {
    Write-Host "Searching for TopSolid..." -ForegroundColor Gray
    
    $commonPaths = @(
        "C:\Missler\V627",
        "C:\Missler\V626",
        "C:\Missler\V625",
        "C:\Missler\V624",
        "C:\Program Files\Missler\TopSolid",
        "C:\Program Files (x86)\Missler\TopSolid"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    
    return $null
}

function Find-TopSolidExe {
    param([string]$BasePath)
    
    $exeNames = @("top627.exe", "top626.exe", "top625.exe", "top624.exe", "TopSolid.exe")
    
    foreach ($exe in $exeNames) {
        $candidate = Join-Path $BasePath "bin" $exe
        if (Test-Path $candidate) { return $candidate }
        
        $candidate = Join-Path $BasePath $exe
        if (Test-Path $candidate) { return $candidate }
    }
    
    return $null
}

function Set-IconReg {
    param([string]$Ext, [string]$ExePath, [int]$Index)
    
    $iconVal = "`"$ExePath`",$Index"
    $fixed = $false
    
    try {
        $path = "HKCU:\Software\Classes\$Ext\DefaultIcon"
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force | Out-Null
        }
        Set-ItemProperty -Path $path -Name "(Default)" -Value $iconVal
        Write-Host "  ✓ $Ext (HKCU)" -ForegroundColor Green
        $fixed = $true
    }
    catch {
        Write-Host "  ✗ $Ext (HKCU): $($_)" -ForegroundColor Red
    }
    
    if (Test-AdminRights) {
        try {
            $path = "HKLM:\SOFTWARE\Classes\$Ext\DefaultIcon"
            if (-not (Test-Path $path)) {
                New-Item -Path $path -Force | Out-Null
            }
            Set-ItemProperty -Path $path -Name "(Default)" -Value $iconVal
            Write-Host "  ✓ $Ext (HKLM)" -ForegroundColor Green
        }
        catch {
            Write-Host "  ⚠ $Ext (HKLM): $($_)" -ForegroundColor Yellow
        }
    }
    
    return $fixed
}

Write-Host "`nTopSolid Icon Fixer v2.0 (Windows 10/11)" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$isAdmin = Test-AdminRights
Write-Host "Admin: $(if ($isAdmin) { 'YES' } else { 'NO' })" -ForegroundColor $(if ($isAdmin) { 'Green' } else { 'Yellow' })

$topPath = Find-TopSolidPath
if (-not $topPath) {
    Write-Host "`nERROR: TopSolid not found!" -ForegroundColor Red
    exit 1
}

Write-Host "Found: $topPath" -ForegroundColor Green

$exePath = Find-TopSolidExe $topPath
if (-not $exePath) {
    Write-Host "ERROR: Executable not found!" -ForegroundColor Red
    exit 1
}

Write-Host "Using: $(Split-Path $exePath -Leaf)`n" -ForegroundColor Green

Write-Host "Fixing icons..." -ForegroundColor Cyan
$fixed = 0
$fixed += (Set-IconReg ".top" $exePath 0)
$fixed += (Set-IconReg ".dft" $exePath 1)
$fixed += (Set-IconReg ".vdx" $exePath 2)
$fixed += (Set-IconReg ".prj" $exePath 3)

Write-Host "`nRefreshing Explorer..." -ForegroundColor Gray
try {
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    Start-Process explorer
    Write-Host "✓ Explorer refreshed`n" -ForegroundColor Green
}
catch {
    Write-Host "⚠ Could not refresh (try F5)`n" -ForegroundColor Yellow
}

Write-Host "✓ Done! ($fixed fixes applied)" -ForegroundColor Green
