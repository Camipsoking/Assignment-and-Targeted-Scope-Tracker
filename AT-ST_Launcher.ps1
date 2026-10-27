#Requires -Version 5.1
<#
===============================================================================
 AT-ST_Launcher.ps1
 Thin PS 5.1 launcher that AT-ST-builder.ps1 compiles into AT-ST.exe.

 Job:
   1. Verify AT-ST.ps1 is signed by the SAME certificate that signed this
      launcher (refuses to run a modified or substituted script).
   2. Locate pwsh.exe (PowerShell 7).
   3. Launch AT-ST.ps1 under PowerShell 7, hidden, in STA mode.

 Signature check rules:
   - Launcher signed + Valid      -> AT-ST.ps1 must be Valid AND signed with
                                     the same certificate thumbprint.
   - Launcher NOT signed          -> development build; check is skipped.
   - Launcher signature broken    -> refuse to run (tampered or untrusted).

 Because the expected certificate is read from the launcher's own signature,
 nothing is hardcoded. Script updates only need the script re-signed with the
 same cert. A renewed signing cert means re-signing the EXE as well.

 Install folder layout:
   AT-ST.exe     (this launcher, signed)
   AT-ST.ps1     (the app, signed with the same cert)
   config.json   (tenant/app IDs)
   version.txt   (app version, written by the builder)

 Launcher events are logged to:
   %ProgramData%\IntuneAssignmentViewer\launcher.log
===============================================================================
#>

Add-Type -AssemblyName PresentationFramework

$LogDir  = "$env:ProgramData\IntuneAssignmentViewer"
$LogPath = Join-Path $LogDir "launcher.log"

function Write-LauncherLog {
    param ([string]$Message)
    try {
        if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
        Add-Content -Path $LogPath -Value ("{0} | {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
    } catch { }
}

function Show-LauncherError {
    param ([string]$Text, [string]$Title)
    Write-LauncherLog "ERROR: $Title - $($Text -replace "`r?`n", ' ')"
    [System.Windows.MessageBox]::Show($Text, $Title, "OK", "Error") | Out-Null
}

# --- Resolve own path (.ps1 when testing, .exe when compiled) ------------
$selfPath = $null
if ($PSCommandPath) {
    $selfPath = $PSCommandPath
} else {
    try { $selfPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName } catch { }
}

if ($selfPath) {
    $here = Split-Path -Parent $selfPath
} else {
    $here = (Get-Location).Path
}

$appScript = Join-Path $here "AT-ST.ps1"
Write-LauncherLog "Launcher started from $here"

if (-not (Test-Path $appScript)) {
    Show-LauncherError -Title "AT-ST - Missing Application File" -Text (
        "AT-ST.ps1 was not found next to the launcher.`n`nExpected at:`n$appScript`n`n" +
        "The application files may be corrupt or incompletely deployed.")
    return
}

# --- Signature verification ----------------------------------------------
Import-Module Microsoft.PowerShell.Security -ErrorAction SilentlyContinue

$selfSig = $null
if ($selfPath) {
    try { $selfSig = Get-AuthenticodeSignature -FilePath $selfPath -ErrorAction Stop } catch { }
}

if (-not $selfSig -or $selfSig.Status -eq 'NotSigned') {
    Write-LauncherLog "Launcher is not signed (development build). Skipping AT-ST.ps1 signature check."
}
elseif ($selfSig.Status -ne 'Valid') {
    Show-LauncherError -Title "AT-ST - Launcher Signature Invalid" -Text (
        "The AT-ST launcher's own digital signature is not valid ($($selfSig.Status)).`n`n" +
        "The application may have been modified or this machine does not trust the signing certificate. " +
        "Contact IT to reinstall AT-ST.")
    return
}
else {
    $expectedThumb = $selfSig.SignerCertificate.Thumbprint

    $appSig = $null
    try { $appSig = Get-AuthenticodeSignature -FilePath $appScript -ErrorAction Stop } catch { }

    if (-not $appSig -or $appSig.Status -ne 'Valid') {
        $status = if ($appSig) { $appSig.Status } else { "Unreadable" }
        Show-LauncherError -Title "AT-ST - Application Signature Invalid" -Text (
            "AT-ST.ps1 does not have a valid digital signature (status: $status).`n`n" +
            "The application file may have been modified. For security, AT-ST will not start. " +
            "Contact IT to reinstall AT-ST.")
        return
    }

    if ($appSig.SignerCertificate.Thumbprint -ne $expectedThumb) {
        Show-LauncherError -Title "AT-ST - Unexpected Publisher" -Text (
            "AT-ST.ps1 is signed, but not by the same certificate as the launcher.`n`n" +
            "Expected: $expectedThumb`nFound:    $($appSig.SignerCertificate.Thumbprint)`n`n" +
            "For security, AT-ST will not start. Contact IT to reinstall AT-ST.")
        return
    }

    Write-LauncherLog "AT-ST.ps1 signature verified (signer $expectedThumb)."
}

# --- Locate pwsh.exe (PowerShell 7) --------------------------------------
function Find-Pwsh {
    $cmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $candidates = @(
        "$env:ProgramFiles\PowerShell\7\pwsh.exe",
        "${env:ProgramFiles(x86)}\PowerShell\7\pwsh.exe",
        "$env:LOCALAPPDATA\Microsoft\PowerShell\7\pwsh.exe"
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path $c)) { return $c }
    }

    $root = "$env:ProgramFiles\PowerShell"
    if (Test-Path $root) {
        $found = Get-ChildItem $root -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^7' } |
            ForEach-Object { Join-Path $_.FullName "pwsh.exe" } |
            Where-Object { Test-Path $_ } |
            Select-Object -First 1
        if ($found) { return $found }
    }
    return $null
}

$pwsh = Find-Pwsh
if (-not $pwsh) {
    Show-LauncherError -Title "AT-ST - PowerShell 7 Required" -Text (
        "PowerShell 7 (pwsh.exe) was not found on this machine.`n`n" +
        "AT-ST requires PowerShell 7. It is normally deployed automatically. " +
        "If this message persists, contact IT.")
    return
}

# --- Launch the app under PS7 --------------------------------------------
try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName         = $pwsh
    $psi.Arguments        = "-NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File `"$appScript`""
    $psi.WorkingDirectory = $here
    $psi.UseShellExecute  = $false
    $psi.CreateNoWindow   = $true
    [void][System.Diagnostics.Process]::Start($psi)
    Write-LauncherLog "Started AT-ST.ps1 with $pwsh"
} catch {
    Show-LauncherError -Title "AT-ST - Launch Failed" -Text (
        "Failed to launch AT-ST under PowerShell 7.`n`n$($_.Exception.Message)")
}
