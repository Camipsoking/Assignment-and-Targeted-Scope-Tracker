#Requires -Version 5.1
<#
===============================================================================
 AT-ST-builder.ps1
 Build + sign pipeline for AT-ST (Assignment and Targeted Scope Tracker)

 OUTPUT (.\dist):
   AT-ST.exe     compiled launcher (signed, with icon)
   AT-ST.ps1     the app script (signed with the same cert)
   config.json   tenant/app IDs (kept between builds unless -ForceConfig)
   version.txt   app version, read from $AppVersion in the app script

 MODES
   Full build (default)
     Generates the icon, compiles the launcher, stages and signs both files.
     Needed the first time, when the launcher changes, or when the signing
     certificate is renewed.

   -ScriptOnly
     Stages and signs ONLY AT-ST.ps1 and updates version.txt. Reuses the
     existing signed AT-ST.exe. Use this for normal app updates.
     The script must be signed with the SAME cert as the EXE, or the launcher
     will refuse to run it - the builder checks this for you.

 VERSION
   Read automatically from $AppVersion in the app script. If you pass
   -Version and it doesn't match, the build stops.

 ICON
   If -IconFile isn't given, AT-ST.ico is generated from the logo Base64
   embedded in the app script and saved next to this builder.

 EXAMPLES
   .\AT-ST-builder.ps1 -CertThumbprint "ABC123..."              # full build
   .\AT-ST-builder.ps1 -CertThumbprint "ABC123..." -ScriptOnly  # app update
   .\AT-ST-builder.ps1 -NoSign                                  # unsigned test
===============================================================================
#>

[CmdletBinding()]
param(
    [string]$LauncherScript = ".\AT-ST_Launcher.ps1",
    [string]$AppScript      = ".\AT-ST_IntuneAssignmentViewer_v5.6.ps1",
    [string]$ConfigFile     = ".\config.json",
    [string]$DistDir        = ".\dist",

    # Optional; defaults to $AppVersion read from the app script
    [string]$Version,

    [string]$Company     = "Company Name",
    [string]$Product     = "Assignment and Targeted Scope Tracker",
    [string]$Description = "Intune group assignment auditing tool",

    # Optional; if omitted, generated from the embedded logo
    [string]$IconFile,

    [string]$CertThumbprint,
    [string]$TimestampServer = "http://timestamp.digicert.com",

    [switch]$ScriptOnly,
    [switch]$ForceConfig,
    [switch]$NoSign
)

$ErrorActionPreference = "Stop"

function Write-Step  { param([string]$m) Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Write-Ok    { param([string]$m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn2 { param([string]$m) Write-Host "  [!]  $m" -ForegroundColor Yellow }

$BuilderDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# ---------------------------------------------------------------------------
# ICON GENERATION (multi-size .ico from a PNG Base64 string)
#   Sizes < 256 are stored as 32-bit DIBs (most compatible);
#   256 is stored as PNG (standard for Vista+ icons).
# ---------------------------------------------------------------------------
function Get-DibBytes {
    param ([System.Drawing.Bitmap]$Bitmap)

    $s    = $Bitmap.Width
    $rect = New-Object System.Drawing.Rectangle 0, 0, $s, $s
    $data = $Bitmap.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
                             [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $stride = $data.Stride
    $raw    = New-Object byte[] ($stride * $s)
    [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $raw, 0, $raw.Length)
    $Bitmap.UnlockBits($data)

    $maskRow = [int]([Math]::Floor(($s + 31) / 32) * 4)
    $ms = New-Object System.IO.MemoryStream
    $w  = New-Object System.IO.BinaryWriter($ms)

    # BITMAPINFOHEADER (height doubled to include the AND mask)
    $w.Write([UInt32]40)
    $w.Write([Int32]$s)
    $w.Write([Int32]($s * 2))
    $w.Write([UInt16]1)
    $w.Write([UInt16]32)
    $w.Write([UInt32]0)
    $w.Write([UInt32]($s * $s * 4 + $maskRow * $s))
    $w.Write([Int32]0)
    $w.Write([Int32]0)
    $w.Write([UInt32]0)
    $w.Write([UInt32]0)

    # Pixel rows, bottom-up, BGRA
    for ($row = $s - 1; $row -ge 0; $row--) {
        $w.Write($raw, $row * $stride, $s * 4)
    }

    # AND mask (all zero - transparency comes from the alpha channel)
    $w.Write((New-Object byte[] ($maskRow * $s)))

    $w.Flush()
    $bytes = $ms.ToArray()
    $w.Close()
    return ,$bytes
}

function New-AppIcon {
    param ([string]$Base64, [string]$OutPath)

    Add-Type -AssemblyName System.Drawing

    $srcBytes  = [Convert]::FromBase64String($Base64)
    $srcStream = New-Object System.IO.MemoryStream(,$srcBytes)
    $src       = [System.Drawing.Image]::FromStream($srcStream)

    $entries = New-Object System.Collections.ArrayList
    try {
        foreach ($s in @(16, 24, 32, 48, 64, 128, 256)) {
            $bmp = New-Object System.Drawing.Bitmap $s, $s, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.Clear([System.Drawing.Color]::Transparent)
            $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

            # Fit the logo inside the square, preserving aspect ratio
            $scale = [Math]::Min($s / $src.Width, $s / $src.Height)
            $w = [int][Math]::Round($src.Width * $scale)
            $h = [int][Math]::Round($src.Height * $scale)
            $x = [int](($s - $w) / 2)
            $y = [int](($s - $h) / 2)
            $g.DrawImage($src, $x, $y, $w, $h)
            $g.Dispose()

            if ($s -ge 256) {
                $pms = New-Object System.IO.MemoryStream
                $bmp.Save($pms, [System.Drawing.Imaging.ImageFormat]::Png)
                $bytes = $pms.ToArray()
                $pms.Dispose()
            } else {
                $bytes = Get-DibBytes -Bitmap $bmp
            }
            $bmp.Dispose()

            [void]$entries.Add(@{ Size = $s; Bytes = $bytes })
        }
    } finally {
        $src.Dispose()
        $srcStream.Dispose()
    }

    $fs = [System.IO.File]::Create($OutPath)
    $bw = New-Object System.IO.BinaryWriter($fs)
    try {
        # ICONDIR
        $bw.Write([UInt16]0)
        $bw.Write([UInt16]1)
        $bw.Write([UInt16]$entries.Count)

        # ICONDIRENTRY x N
        $offset = 6 + (16 * $entries.Count)
        foreach ($e in $entries) {
            $dim = if ($e.Size -ge 256) { 0 } else { $e.Size }
            $bw.Write([byte]$dim)
            $bw.Write([byte]$dim)
            $bw.Write([byte]0)
            $bw.Write([byte]0)
            $bw.Write([UInt16]1)
            $bw.Write([UInt16]32)
            $bw.Write([UInt32]$e.Bytes.Length)
            $bw.Write([UInt32]$offset)
            $offset += $e.Bytes.Length
        }

        # Image data
        foreach ($e in $entries) { $bw.Write([byte[]]$e.Bytes) }
    } finally {
        $bw.Close()
    }
}

# ---------------------------------------------------------------------------
# 1. VALIDATE SOURCES + VERSION
# ---------------------------------------------------------------------------
Write-Step "Validating sources"

if (-not (Test-Path $AppScript)) { throw "App script not found: $AppScript" }
Write-Ok "App script: $AppScript"

$appContent = Get-Content $AppScript -Raw

$scriptVer = $null
if ($appContent -match '\$AppVersion\s*=\s*"([^"]+)"') { $scriptVer = $Matches[1] }

if (-not $Version) {
    if (-not $scriptVer) { throw "Could not read `$AppVersion from $AppScript. Add it or pass -Version." }
    $Version = $scriptVer
    Write-Ok "Version read from script: $Version"
} elseif ($scriptVer -and $scriptVer -ne $Version) {
    throw "Version mismatch: -Version '$Version' but the script's `$AppVersion is '$scriptVer'. Update one so they match."
} else {
    Write-Ok "Version: $Version"
}

$parsedVer = $null
if (-not [version]::TryParse($Version, [ref]$parsedVer)) {
    throw "Version '$Version' is not a valid version number (expected e.g. 5.2.0)."
}
$FileVersion = "{0}.{1}.{2}.{3}" -f $parsedVer.Major, $parsedVer.Minor, [Math]::Max($parsedVer.Build, 0), [Math]::Max($parsedVer.Revision, 0)

$logoB64 = $null
if ($appContent -match '(?s)\$LogoBase64\s*=\s*@''\s*(.*?)\s*''@') {
    $logoB64 = $Matches[1] -replace '\s', ''
}
$hasLogo = $logoB64 -and $logoB64 -ne 'PASTE_YOUR_BASE64_HERE'
if (-not $hasLogo) {
    Write-Warn2 "Logo Base64 placeholder is still in the app script. The app will run without a logo."
}

if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir -Force | Out-Null }
$outExe = Join-Path $DistDir "AT-ST.exe"
$outApp = Join-Path $DistDir "AT-ST.ps1"

if ($ScriptOnly) {
    Write-Ok "Mode: SCRIPT ONLY (reusing existing AT-ST.exe)"
    if (-not (Test-Path $outExe)) {
        throw "No existing $outExe found. Run a full build (without -ScriptOnly) first."
    }
} else {
    Write-Ok "Mode: FULL BUILD"
    if (-not (Test-Path $LauncherScript)) { throw "Launcher script not found: $LauncherScript" }
    Write-Ok "Launcher: $LauncherScript"
}

# ---------------------------------------------------------------------------
# 2. RESOLVE SIGNING CERTIFICATE
# ---------------------------------------------------------------------------
$signingCert = $null
if (-not $NoSign) {
    Write-Step "Resolving code-signing certificate"

    if ($CertThumbprint) {
        $tp = $CertThumbprint -replace '\s', ''
        $signingCert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -CodeSigningCert -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $tp } | Select-Object -First 1
        if (-not $signingCert) { throw "No code-signing cert with thumbprint '$CertThumbprint' found." }
    } else {
        $signingCert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -CodeSigningCert -ErrorAction SilentlyContinue |
            Where-Object { $_.NotAfter -gt (Get-Date) } |
            Sort-Object NotAfter -Descending | Select-Object -First 1
        if (-not $signingCert) { throw "No valid code-signing certificate found. Enroll from your enterprise CA or pass -CertThumbprint." }
        Write-Warn2 "No -CertThumbprint provided; auto-selected:"
    }

    Write-Ok ("Subject   : {0}" -f $signingCert.Subject)
    Write-Ok ("Thumbprint: {0}" -f $signingCert.Thumbprint)
    Write-Ok ("Expires   : {0}" -f $signingCert.NotAfter)
}

# ---------------------------------------------------------------------------
# 3. SCRIPT-ONLY COMPATIBILITY CHECK
#    The launcher only runs a script signed by its own certificate.
# ---------------------------------------------------------------------------
if ($ScriptOnly) {
    Write-Step "Checking existing AT-ST.exe signature"
    $exeSig = Get-AuthenticodeSignature -FilePath $outExe

    if ($exeSig.Status -eq 'NotSigned') {
        Write-Warn2 "Existing AT-ST.exe is unsigned (test build). Its launcher skips the signature check."
    } elseif ($exeSig.Status -ne 'Valid') {
        throw "Existing AT-ST.exe signature is $($exeSig.Status). Run a full build."
    } elseif ($NoSign) {
        throw "Existing AT-ST.exe is signed, so its launcher will refuse an unsigned AT-ST.ps1. Remove -NoSign."
    } elseif ($exeSig.SignerCertificate.Thumbprint -ne $signingCert.Thumbprint) {
        throw ("AT-ST.exe was signed with a different certificate ({0}).`n" -f $exeSig.SignerCertificate.Thumbprint) +
              "The launcher will refuse a script signed with $($signingCert.Thumbprint). " +
              "Run a full build to re-sign both files with the current certificate."
    } else {
        Write-Ok "AT-ST.exe signer matches the signing certificate"
    }
}

# ---------------------------------------------------------------------------
# 4. FULL BUILD: ICON + COMPILE LAUNCHER
# ---------------------------------------------------------------------------
if (-not $ScriptOnly) {
    Write-Step "Preparing icon"

    $iconPath = $null
    if ($IconFile) {
        if (-not (Test-Path $IconFile)) { throw "Icon file not found: $IconFile" }
        $iconPath = (Resolve-Path $IconFile).Path
        Write-Ok "Using provided icon: $iconPath"
    } elseif ($hasLogo) {
        $iconPath = Join-Path $BuilderDir "AT-ST.ico"
        try {
            New-AppIcon -Base64 $logoB64 -OutPath $iconPath
            Write-Ok "Generated icon from embedded logo: $iconPath"
        } catch {
            Write-Warn2 "Icon generation failed: $($_.Exception.Message). Building without an icon."
            $iconPath = $null
        }
    } else {
        Write-Warn2 "No logo and no -IconFile. Building without an icon."
    }

    Write-Step "Compiling launcher to AT-ST.exe"

    if (-not (Get-Module -ListAvailable -Name ps2exe)) {
        Write-Warn2 "ps2exe not found. Installing for current user..."
        Install-Module ps2exe -Scope CurrentUser -Force
    }
    Import-Module ps2exe -ErrorAction Stop

    $ps2exeArgs = @{
        InputFile   = (Resolve-Path $LauncherScript).Path
        OutputFile  = $outExe
        Title       = $Product
        Product     = $Product
        Company     = $Company
        Version     = $FileVersion
        Description = $Description
        NoConsole   = $true
        STA         = $true
    }
    if ($iconPath) { $ps2exeArgs.IconFile = $iconPath }

    Invoke-ps2exe @ps2exeArgs
    if (-not (Test-Path $outExe)) { throw "ps2exe did not produce $outExe" }
    Write-Ok "Compiled: $outExe (file version $FileVersion)"
}

# ---------------------------------------------------------------------------
# 5. STAGE APP SCRIPT
# ---------------------------------------------------------------------------
Write-Step "Staging app script"
Copy-Item $AppScript -Destination $outApp -Force
Write-Ok "Staged: $outApp"

# ---------------------------------------------------------------------------
# 6. SIGN + VERIFY
# ---------------------------------------------------------------------------
if ($ScriptOnly) { $toSign = @($outApp) } else { $toSign = @($outExe, $outApp) }

if (-not $NoSign) {
    Write-Step "Signing"
    foreach ($target in $toSign) {
        $signParams = @{
            FilePath      = $target
            Certificate   = $signingCert
            HashAlgorithm = "SHA256"
        }
        if ($TimestampServer) { $signParams.TimestampServer = $TimestampServer }

        $res = Set-AuthenticodeSignature @signParams
        if ($res.Status -ne "Valid") {
            throw "Signing failed for $target : $($res.Status) - $($res.StatusMessage)"
        }
        Write-Ok "Signed: $target"
    }

    Write-Step "Verifying signatures"
    $exeV = Get-AuthenticodeSignature -FilePath $outExe
    $appV = Get-AuthenticodeSignature -FilePath $outApp
    foreach ($v in @($exeV, $appV)) {
        $ts = if ($v.TimeStamperCertificate) { "timestamped" } else { "NOT timestamped" }
        Write-Ok ("{0} : {1} ({2})" -f (Split-Path $v.Path -Leaf), $v.Status, $ts)
    }
    if ($appV.Status -ne 'Valid') { throw "AT-ST.ps1 verification failed: $($appV.Status)" }
    if ($exeV.Status -eq 'Valid' -and $exeV.SignerCertificate.Thumbprint -ne $appV.SignerCertificate.Thumbprint) {
        throw "AT-ST.exe and AT-ST.ps1 are signed by different certificates. The launcher will refuse to start."
    }
    if ($exeV.Status -eq 'Valid') { Write-Ok "Launcher and script signers match - launcher check will pass" }
} else {
    Write-Warn2 "-NoSign: outputs are UNSIGNED. Do not deploy unsigned to production."
}

# ---------------------------------------------------------------------------
# 7. VERSION FILE (used by Intune install/detection in Phase 4)
# ---------------------------------------------------------------------------
Write-Step "Writing version.txt"
Set-Content -Path (Join-Path $DistDir "version.txt") -Value $Version -Encoding ASCII
Write-Ok "version.txt = $Version"

# ---------------------------------------------------------------------------
# 8. CONFIG (preserve existing values unless -ForceConfig)
# ---------------------------------------------------------------------------
Write-Step "Config"
$distConfig = Join-Path $DistDir "config.json"
if ((Test-Path $distConfig) -and -not $ForceConfig) {
    Write-Ok "Kept existing $distConfig (use -ForceConfig to overwrite)"
} elseif (Test-Path $ConfigFile) {
    Copy-Item $ConfigFile -Destination $distConfig -Force
    Write-Ok "Copied $ConfigFile to $distConfig"
} else {
    Write-Warn2 "No config at $ConfigFile and none in $DistDir. The app will create a template on first launch."
}

# ---------------------------------------------------------------------------
# DONE
# ---------------------------------------------------------------------------
Write-Step "Build complete"
Write-Host ""
Write-Host "  $DistDir contains:" -ForegroundColor Green
Write-Host "    AT-ST.exe    $(if ($ScriptOnly) { '(unchanged)' } else { '(rebuilt)' })"
Write-Host "    AT-ST.ps1    (v$Version)"
Write-Host "    config.json"
Write-Host "    version.txt"
Write-Host ""
