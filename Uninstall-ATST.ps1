#Requires -Version 5.1
<#
===============================================================================
 Uninstall-ATST.ps1
 Intune Win32 uninstall script for AT-ST

 Removes the install folder, Start menu shortcut, detection registry key and
 Apps & features entry.

 By default it KEEPS C:\ProgramData\IntuneAssignmentViewer (audit/launcher
 logs, group cache, theme preference). Pass -RemoveAppData to delete it.

 Exit codes: 0 = success, 1 = failure
 Log: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AT-ST-Install.log
===============================================================================
#>
param (
    [switch]$RemoveAppData
)

# --- Relaunch in 64-bit PowerShell if needed ------------------------------
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $ps64 = Join-Path $env:WINDIR "sysnative\WindowsPowerShell\v1.0\powershell.exe"
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    if ($RemoveAppData) { $argList += " -RemoveAppData" }
    $proc = Start-Process -FilePath $ps64 -ArgumentList $argList -Wait -PassThru -WindowStyle Hidden
    exit $proc.ExitCode
}

$ErrorActionPreference = "Stop"

$InstallDir   = Join-Path $env:ProgramFiles "AT-ST"
$DataDir      = Join-Path $env:ProgramData "IntuneAssignmentViewer"
$ShortcutPath = Join-Path $env:ProgramData "Microsoft\Windows\Start Menu\Programs\AT-ST.lnk"
$RegKey       = "HKLM:\SOFTWARE\AT-ST"
$UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AT-ST"
$LogDir       = Join-Path $env:ProgramData "Microsoft\IntuneManagementExtension\Logs"
$LogFile      = Join-Path $LogDir "AT-ST-Install.log"

function Write-InstallLog {
    param ([string]$Message)
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    Add-Content -Path $LogFile -Value ("{0} | {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
}

try {
    Write-InstallLog "===== Uninstall started ====="

    # Move out of the install folder in case this script is running from it
    Set-Location -Path $env:TEMP

    if (Test-Path $ShortcutPath) {
        Remove-Item -Path $ShortcutPath -Force
        Write-InstallLog "Removed Start menu shortcut"
    }

    if (Test-Path $InstallDir) {
        Remove-Item -Path $InstallDir -Recurse -Force
        Write-InstallLog "Removed $InstallDir"
    }

    foreach ($key in @($RegKey, $UninstallKey)) {
        if (Test-Path $key) {
            Remove-Item -Path $key -Recurse -Force
            Write-InstallLog "Removed $key"
        }
    }

    if ($RemoveAppData -and (Test-Path $DataDir)) {
        Remove-Item -Path $DataDir -Recurse -Force
        Write-InstallLog "Removed app data at $DataDir"
    }

    Write-InstallLog "===== Uninstall completed successfully ====="
    exit 0
}
catch {
    Write-InstallLog "ERROR: $($_.Exception.Message)"
    Write-InstallLog "===== Uninstall FAILED ====="
    exit 1
}
