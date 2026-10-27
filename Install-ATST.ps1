#Requires -Version 5.1
<#
===============================================================================
 Install-ATST.ps1  (rev 2)
 Intune Win32 install script for AT-ST (Assignment and Targeted Scope Tracker)

 rev 2 changes:
   - Registry values are written with the .NET registry API instead of
     Set-ItemProperty (avoids PowerShell registry-provider errors such as
     "The property string ... does not exist or was not found")
   - version.txt is sanitized: BOM, quotes, whitespace, and prefixes like
     "Version=" are stripped; only the version number is kept
   - Every step is logged, and errors include the failing line number

 Runs as SYSTEM. Expects these files next to it (the packaged dist folder):
   AT-ST.exe, AT-ST.ps1, config.json, version.txt
   AT-ST.ico          (optional - Start menu / window icon)
   Modules\           (optional - bundled Microsoft.Graph.Authentication)
   Uninstall-ATST.ps1

 Exit codes: 0 = success, 1 = failure
 Log: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AT-ST-Install.log
===============================================================================
#>

# --- Relaunch in 64-bit PowerShell if needed ------------------------------
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $ps64 = Join-Path $env:WINDIR "sysnative\WindowsPowerShell\v1.0\powershell.exe"
    $proc = Start-Process -FilePath $ps64 -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Wait -PassThru -WindowStyle Hidden
    exit $proc.ExitCode
}

$ErrorActionPreference = "Stop"

$DisplayName  = "AT-ST Assignment and Targeted Scope Tracker"
$Publisher    = ""
$InstallDir   = Join-Path $env:ProgramFiles "AT-ST"
$DataDir      = Join-Path $env:ProgramData "IntuneAssignmentViewer"
$ShortcutPath = Join-Path $env:ProgramData "Microsoft\Windows\Start Menu\Programs\AT-ST.lnk"
$RegSubKey    = "SOFTWARE\AT-ST"
$UninstSubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AT-ST"
$LogDir       = Join-Path $env:ProgramData "Microsoft\IntuneManagementExtension\Logs"
$LogFile      = Join-Path $LogDir "AT-ST-Install.log"
$Source       = $PSScriptRoot

function Write-InstallLog {
    param ([string]$Message)
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    Add-Content -Path $LogFile -Value ("{0} | {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
}

# Write registry values via .NET (64-bit HKLM view), bypassing the PS provider
function Set-RegValues {
    param (
        [string]$SubKey,
        [hashtable]$StringValues,
        [hashtable]$DwordValues = @{}
    )
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                [Microsoft.Win32.RegistryView]::Registry64)
    $key = $base.CreateSubKey($SubKey)
    try {
        foreach ($name in $StringValues.Keys) {
            $key.SetValue($name, [string]$StringValues[$name], [Microsoft.Win32.RegistryValueKind]::String)
        }
        foreach ($name in $DwordValues.Keys) {
            $key.SetValue($name, [int]$DwordValues[$name], [Microsoft.Win32.RegistryValueKind]::DWord)
        }
    } finally {
        $key.Close()
        $base.Close()
    }
}

$step = "start"
try {
    Write-InstallLog "===== Install started from $Source ====="

    # --- Validate package contents ---------------------------------------
    $step = "validate package"
    foreach ($f in @("AT-ST.exe", "AT-ST.ps1", "config.json", "version.txt")) {
        if (-not (Test-Path (Join-Path $Source $f))) { throw "Package is missing required file: $f" }
    }

    $step = "read version.txt"
    $rawVersion = [System.IO.File]::ReadAllText((Join-Path $Source "version.txt"))
    Write-InstallLog ("version.txt raw content: [{0}]" -f ($rawVersion -replace "`r", '\r' -replace "`n", '\n'))
    $m = [regex]::Match($rawVersion, '\d+(\.\d+){1,3}')
    if (-not $m.Success) { throw "version.txt does not contain a version number (expected e.g. 5.6.0)" }
    $Version = $m.Value
    Write-InstallLog "Package version: $Version"

    # --- Clean previous install, copy files -------------------------------
    $step = "copy files"
    if (Test-Path $InstallDir) {
        Write-InstallLog "Removing previous install at $InstallDir"
        Remove-Item -Path $InstallDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    Get-ChildItem -Path $Source | Where-Object { $_.Name -ne "Install-ATST.ps1" } | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination $InstallDir -Recurse -Force
    }
    Write-InstallLog "Copied files to $InstallDir"

    $step = "check signatures"
    foreach ($f in @("AT-ST.exe", "AT-ST.ps1")) {
        $sig = Get-AuthenticodeSignature -FilePath (Join-Path $InstallDir $f)
        $subj = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { "no signer" }
        Write-InstallLog ("Signature {0}: {1} ({2})" -f $f, $sig.Status, $subj)
    }

    if (Test-Path (Join-Path $InstallDir "Modules\Microsoft.Graph.Authentication")) {
        Write-InstallLog "Bundled Microsoft.Graph.Authentication module installed"
    } else {
        Write-InstallLog "WARNING: No bundled Graph module. Users will need Microsoft.Graph.Authentication installed."
    }

    # --- App data folder (writable by standard users) ---------------------
    $step = "app data folder"
    if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir -Force | Out-Null }
    $acl   = Get-Acl -Path $DataDir
    $users = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-545")   # BUILTIN\Users
    $rule  = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $users, "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.AddAccessRule($rule)
    Set-Acl -Path $DataDir -AclObject $acl
    Write-InstallLog "Granted Users Modify on $DataDir"

    # --- Start menu shortcut ---------------------------------------------
    $step = "shortcut"
    $exePath = Join-Path $InstallDir "AT-ST.exe"
    $icoPath = Join-Path $InstallDir "AT-ST.ico"
    $iconRef = if (Test-Path $icoPath) { $icoPath } else { $exePath }
    $shell = New-Object -ComObject WScript.Shell
    $lnk   = $shell.CreateShortcut($ShortcutPath)
    $lnk.TargetPath       = $exePath
    $lnk.WorkingDirectory = $InstallDir
    $lnk.Description      = $DisplayName
    $lnk.IconLocation     = "$iconRef,0"
    $lnk.Save()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    Write-InstallLog "Created Start menu shortcut: $ShortcutPath"

    # --- Detection registry value ----------------------------------------
    $step = "detection registry"
    Set-RegValues -SubKey $RegSubKey -StringValues @{
        Version     = $Version
        InstallPath = $InstallDir
    }
    Write-InstallLog "Set HKLM\$RegSubKey\Version = $Version"

    # --- Apps & features entry -------------------------------------------
    $step = "apps and features"
    Set-RegValues -SubKey $UninstSubKey -StringValues @{
        DisplayName     = $DisplayName
        DisplayVersion  = $Version
        Publisher       = $Publisher
        InstallLocation = $InstallDir
        DisplayIcon     = $iconRef
        UninstallString = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$InstallDir\Uninstall-ATST.ps1`""
    } -DwordValues @{
        NoModify = 1
        NoRepair = 1
    }
    Write-InstallLog "Registered in Apps & features"

    Write-InstallLog "===== Install completed successfully ====="
    exit 0
}
catch {
    Write-InstallLog "ERROR during step '$step' (line $($_.InvocationInfo.ScriptLineNumber)): $($_.Exception.Message)"
    Write-InstallLog "===== Install FAILED ====="
    exit 1
}
