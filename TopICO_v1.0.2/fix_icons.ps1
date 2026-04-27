# TopSolid Icon Fixer - PowerShell Version (Windows 10/11)
# Run this script with Administrator privileges for best results

param([switch]$Force)

function Test-AdminRights {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-OSVersion {
    $osVersion = [System.Environment]::OSVersion
    return "$($osVersion.Platform) $($osVersion.Version)"
}

function Find-TopSolidPath {
    Write-Host "Searching for TopSolid installation..." -ForegroundColor Gray
    
    # Try registry first
    $regPath = @(
        "HKLM:\SOFTWARE\Missler",
        "HKLM:\SOFTWARE\WOW6432Node\Missler"
    )
    
    foreach ($path in $regPath) {
        if (Test-Path $path) {
            try {
                $topPath = Get-ItemProperty $path -Name "TopSolidPath" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty TopSolidPath
                if ($topPath -and (Test-Path $topPath)) {
                    return $topPath
                }
            } catch {}
        }
    }
    
    # Try common installation paths
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

function Find-TopSolidExecutable {
    param([string]$BasePath)
    
    $exeNames = @("top627.exe", "top626.exe", "top625.exe", "top624.exe", "TopSolid.exe")
    
    # Try bin subdirectory
    foreach ($exe in $exeNames) {
        $candidate = Join-Path $BasePath "bin" $exe
        if (Test-Path $candidate) {
            return $candidate
        }
    }
    
    # Try root directory
    foreach ($exe in $exeNames) {
        $candidate = Join-Path $BasePath $exe
        if (Test-Path $candidate) {
            return $candidate
        }
    }
    
    return $null
}

function Set-IconAssociation {
    param(
        [string]$Extension,
        [string]$IconPath,
        [int]$IconIndex
    )
    
    $regValue = "`"$IconPath`",$IconIndex"
    $success = $false
    
    # Try HKEY_CURRENT_USER (always works)
    try {
        $path = "HKCU:\Software\Classes\$Extension\DefaultIcon"
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $path -Name "(Default)" -Value $regValue -ErrorAction Stop
        Write-Host "  ✓ Fixed $Extension (HKEY_CURRENT_USER)"
        $success = $true
    } catch {
        Write-Host "  ✗ Failed to set $Extension in HKEY_CURRENT_USER: $_" -ForegroundColor Red
    }
    
    # Try HKEY_LOCAL_MACHINE (may need admin)
    try {
        $path = "HKLM:\SOFTWARE\Classes\$Extension\DefaultIcon"
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $path -Name "(Default)" -Value $regValue -ErrorAction Stop
        Write-Host "  ✓ Fixed $Extension (HKEY_LOCAL_MACHINE - system-wide)"
    } catch {
        # Silent fail for HKLM if not admin
        if ((Test-AdminRights)) {
            Write-Host "  ⚠ Could not set $Extension in HKEY_LOCAL_MACHINE: $_" -ForegroundColor Yellow
        }
    }
    
    return $success
}

function Refresh-Explorer {
    Write-Host "Refreshing Windows Explorer..." -ForegroundColor Gray
    
    try {
        # Method 1: Restart Explorer
        $explorer = Get-Process explorer -ErrorAction SilentlyContinue
        if ($explorer) {
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 1000
            Start-Process explorer
            Write-Host "✓ Explorer restarted"
            return $true
        }
    }
    catch {
        # Continue to next method
    }
    
    # Method 2: Notify shell of change
    try {
        [Runtime.InteropServices.Marshal]::ReleaseComObject([Runtime.InteropServices.Marshal]::GetActiveObject('Shell.Application')) | Out-Null
        Write-Host "✓ Shell refresh completed"
        return $true
    }
    catch {
        # Continue to error message
    }
    
    Write-Host "⚠ Could not refresh Explorer (try F5 on Desktop)" -ForegroundColor Yellow
    return $false
}

function Verify-Fixes {
    Write-Host "`n=== Verification ===" -ForegroundColor Cyan
    
    $desktopPath = Join-Path ([System.Environment]::GetFolderPath("Desktop")) ""
    $topSolidFiles = @(
        Get-ChildItem $desktopPath -Filter "*.top" -ErrorAction SilentlyContinue
        Get-ChildItem $desktopPath -Filter "*.dft" -ErrorAction SilentlyContinue
        Get-ChildItem $desktopPath -Filter "*.vdx" -ErrorAction SilentlyContinue
        Get-ChildItem $desktopPath -Filter "*.prj" -ErrorAction SilentlyContinue
    )
    
    if ($topSolidFiles.Count -gt 0) {
        Write-Host "Found $($topSolidFiles.Count) TopSolid files on Desktop:" -ForegroundColor Green
        foreach ($file in $topSolidFiles) {
            Write-Host "  - $($file.Name)"
        }
    } else {
        Write-Host "No TopSolid files found on Desktop to verify" -ForegroundColor Gray
    }
}

# Main Script
Clear-Host
$osVersion = Get-OSVersion
$isAdmin = Test-AdminRights

Write-Host "TopSolid Icon Fixer v2.0 (Windows 10/11)" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "OS: $osVersion"
Write-Host "Admin Rights: $(if ($isAdmin) { '✓ YES' } else { '✗ NO (limited features)' })" -ForegroundColor $(if ($isAdmin) { 'Green' } else { 'Yellow' })
Write-Host "============================================" -ForegroundColor Cyan

if (-not $isAdmin) {
    Write-Host "`n⚠ WARNING: Not running as Administrator" -ForegroundColor Yellow
    Write-Host "  User-level fixes will be applied."
    Write-Host "  For system-wide fixes, run this script as Administrator.`n" -ForegroundColor Yellow
    
    if (-not $Force) {
        $response = Read-Host "Continue? (Y/n)"
        if ($response -eq "n") { exit }
    }
}

# Find TopSolid
$topSolidPath = Find-TopSolidPath
if (-not $topSolidPath) {
    Write-Host "`n✗ ERROR: TopSolid installation not found!" -ForegroundColor Red
    Write-Host "Please install TopSolid or check registry." -ForegroundColor Red
    exit 1
}

Write-Host "`nFound TopSolid at: $topSolidPath" -ForegroundColor Green

# Find executable
$exePath = Find-TopSolidExecutable $topSolidPath
if (-not $exePath) {
    Write-Host "`n✗ ERROR: TopSolid executable not found!" -ForegroundColor Red
    exit 1
}

Write-Host "Using executable: $exePath`n" -ForegroundColor Green

# Fix associations
Write-Host "Fixing file associations..." -ForegroundColor Cyan
$extensions = @(
    @{ Ext = ".top"; Index = 0 },
    @{ Ext = ".dft"; Index = 1 },
    @{ Ext = ".vdx"; Index = 2 },
    @{ Ext = ".prj"; Index = 3 }
)

$fixedCount = 0
foreach ($ext in $extensions) {
    if (Set-IconAssociation $ext.Ext $exePath $ext.Index) {
        $fixedCount++
    }
}

Write-Host "`n✓ Successfully fixed $fixedCount file associations" -ForegroundColor Green

# Refresh shell
Refresh-Explorer

# Verify
Verify-Fixes

Write-Host "`n✓ Icon repair complete!" -ForegroundColor Green
Write-Host "`nTip: If icons don't appear immediately, try:" -ForegroundColor Gray
Write-Host "  1. Press F5 on Desktop" -ForegroundColor Gray
Write-Host "  2. Delete icon cache: del %LocalAppData%\IconCache.db" -ForegroundColor Gray
Write-Host "  3. Restart Windows Explorer" -ForegroundColor Gray
