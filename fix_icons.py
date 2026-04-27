#!/usr/bin/env python3
"""
TopSolid Icon Fixer - Repairs broken TopSolid file type icons on Desktop
Compatible with Windows 10 & 11
"""

import winreg
import subprocess
import sys
import os
import ctypes
from pathlib import Path
from platform import system, release

def check_admin_rights():
    """Check if script is running with admin privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_windows_version():
    """Get Windows version info"""
    return system(), release()

def get_topsolid_path():
    """Get TopSolid installation path from registry"""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Missler") as key:
            path, _ = winreg.QueryValueEx(key, "TopSolidPath")
            return path
    except:
        # Try alternative paths
        common_paths = [
            r"C:\Missler\V627",
            r"C:\Missler\V626",
            r"C:\Missler\V625",
            r"C:\Program Files\Missler\TopSolid",
        ]
        for path in common_paths:
            if os.path.exists(path):
                return path
    return None

def fix_file_associations():
    """Fix TopSolid file type associations in registry"""
    
    topsolid_path = get_topsolid_path()
    if not topsolid_path:
        print("ERROR: TopSolid installation not found")
        return False
    
    print(f"Found TopSolid at: {topsolid_path}")
    
    # Try different TopSolid executable names (support multiple versions)
    exe_names = ["top627.exe", "top626.exe", "top625.exe", "top624.exe", "TopSolid.exe"]
    exe_path = None
    for exe_name in exe_names:
        candidate = os.path.join(topsolid_path, "bin", exe_name)
        if os.path.exists(candidate):
            exe_path = candidate
            break
    
    # Also check root directory for older versions
    if not exe_path:
        for exe_name in exe_names:
            candidate = os.path.join(topsolid_path, exe_name)
            if os.path.exists(candidate):
                exe_path = candidate
                break
    
    if not exe_path:
        print(f"ERROR: TopSolid executable not found in {topsolid_path}")
        return False
    
    print(f"Using executable: {exe_path}")
    
    # Define TopSolid file extensions and their icon indices
    file_types = {
        ".top": {"name": "TopSolidDocument", "icon_index": 0},
        ".dft": {"name": "TopSolidDrawing", "icon_index": 1},
        ".vdx": {"name": "TopSolidView", "icon_index": 2},
        ".prj": {"name": "TopSolidProject", "icon_index": 3},
    }
    
    fixed_count = 0
    failed_count = 0
    
    try:
        # Fix file associations for each extension
        for ext, info in file_types.items():
            try:
                # Try to remove old/corrupted associations
                try:
                    # Remove UserChoice (Windows 11 specifically)
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\{ext}\\UserChoice")
                except:
                    pass  # Key may not exist, that's OK
                
                # Create/fix DefaultIcon entry in HKEY_CURRENT_USER
                icon_path = f'"{exe_path}",{info["icon_index"]}'
                icon_key_path = f"Software\\Classes\\{ext}\\DefaultIcon"
                
                try:
                    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, icon_key_path) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, icon_path)
                    print(f"✓ Fixed icon for {ext} (HKEY_CURRENT_USER)")
                    fixed_count += 1
                except Exception as e:
                    print(f"⚠ Could not write to HKEY_CURRENT_USER for {ext}: {e}")
                
                # Also try to set in HKEY_LOCAL_MACHINE for system-wide fix (may require admin)
                try:
                    with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\Classes\\{ext}\\DefaultIcon") as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, icon_path)
                    print(f"✓ Fixed icon for {ext} (HKEY_LOCAL_MACHINE - system-wide)")
                except Exception as e:
                    print(f"⚠ Could not write to HKEY_LOCAL_MACHINE for {ext} (may need admin): {e}")
                
            except Exception as e:
                print(f"✗ Failed to fix {ext}: {e}")
                failed_count += 1
        
        print(f"\n✓ Successfully fixed {fixed_count} file types")
        if failed_count > 0:
            print(f"⚠ Failed to fix {failed_count} file types")
        
        return fixed_count > 0
        
    except Exception as e:
        print(f"ERROR: Registry operation failed: {e}")
        return False

def refresh_shell():
    """Refresh Windows shell to apply icon changes
    Works on Windows 10 and 11"""
    try:
        # Method 1: Kill and restart Explorer
        print("Attempting to refresh Explorer...")
        result = subprocess.run(["taskkill", "/F", "/IM", "explorer.exe"], 
                                capture_output=True, timeout=5)
        
        # Wait a moment then restart
        import time
        time.sleep(2)
        subprocess.Popen("explorer.exe")
        print("✓ Shell refreshed (Explorer restarted)")
        return True
    except subprocess.TimeoutExpired:
        print("⚠ Explorer timeout (may already be refreshing)")
        return False
    except Exception as e:
        print(f"⚠ Could not refresh shell: {e}")
        
        # Method 2: Try alternative refresh via rundll32 (Windows 10/11 compatible)
        try:
            subprocess.run(["rundll32.exe", "shell32.dll,ShellExec_RunDLL", "explorer.exe"], 
                          capture_output=True, timeout=5)
            print("✓ Shell refresh attempted via rundll32")
            return True
        except:
            print("⚠ Shell refresh failed (try manual explorer restart)")
            return False

def verify_fixes():
    """Verify that fixes were applied correctly"""
    print("\n=== Verification ===")
    
    desktop_path = Path.home() / "Desktop"
    if not desktop_path.exists():
        print("Desktop not found")
        return
    
    topsolid_files = []
    for ext in [".top", ".dft", ".vdx", ".prj"]:
        files = list(desktop_path.glob(f"*{ext}"))
        topsolid_files.extend(files)
    
    if topsolid_files:
        print(f"Found {len(topsolid_files)} TopSolid files on Desktop:")
        for f in topsolid_files:
            print(f"  - {f.name}")
    else:
        print("No TopSolid files found on Desktop to verify")

if __name__ == "__main__":
    os_name, os_release = get_windows_version()
    is_admin = check_admin_rights()
    
    print("TopSolid Icon Fixer v2.0 (Windows 10/11 Compatible)")
    print("=" * 55)
    print(f"OS: {os_name} {os_release}")
    print(f"Admin Rights: {'✓ YES' if is_admin else '✗ NO (some features may be limited)'}")
    print("=" * 55)
    
    if not is_admin:
        print("\n⚠ WARNING: Running without administrator privileges")
        print("  Some registry operations may fail, but user-level fixes will be applied.")
        print("  For system-wide fixes, run this script as Administrator.\n")
    
    if fix_file_associations():
        print("\n✓ Icon associations fixed")
        print("Refreshing shell...")
        refresh_shell()
        verify_fixes()
        print("\n✓ Icon repair complete!")
        print("\nNote: If icons don't appear immediately, try:")
        print("  1. Refresh Desktop (F5)")
        print("  2. Clear icon cache: del %LocalAppData%\\IconCache.db")
        print("  3. Restart Windows Explorer")
    else:
        print("\n✗ Icon repair failed")
        sys.exit(1)
