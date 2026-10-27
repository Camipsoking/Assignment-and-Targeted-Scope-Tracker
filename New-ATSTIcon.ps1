# New-ATSTIcon.ps1 - converts AT-ST-icon.png into a multi-size AT-ST.ico
param(
    [string]$PngPath = ".\AT-ST-icon.png",
    [string]$IcoPath = ".\AT-ST.ico"
)

Add-Type -AssemblyName System.Drawing

$src = [System.Drawing.Image]::FromFile((Resolve-Path $PngPath).Path)
$sizes = 16, 24, 32, 48, 64, 128, 256
$images = @()

foreach ($s in $sizes) {
    $bmp = New-Object System.Drawing.Bitmap $s, $s
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.DrawImage($src, 0, 0, $s, $s)
    $g.Dispose()

    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $images += ,@($s, $ms.ToArray())
}
$src.Dispose()

$outFull = Join-Path (Get-Location).Path (Split-Path $IcoPath -Leaf)
$fs = [System.IO.File]::Create($outFull)
$bw = New-Object System.IO.BinaryWriter $fs

# ICONDIR header
$bw.Write([UInt16]0)
$bw.Write([UInt16]1)
$bw.Write([UInt16]$images.Count)

# Directory entries
$offset = 6 + (16 * $images.Count)
foreach ($img in $images) {
    $dim = if ($img[0] -ge 256) { 0 } else { $img[0] }
    $bw.Write([byte]$dim)
    $bw.Write([byte]$dim)
    $bw.Write([byte]0)
    $bw.Write([byte]0)
    $bw.Write([UInt16]1)
    $bw.Write([UInt16]32)
    $bw.Write([UInt32]$img[1].Length)
    $bw.Write([UInt32]$offset)
    $offset += $img[1].Length
}

# Image data
foreach ($img in $images) { $bw.Write([byte[]]$img[1]) }
$bw.Close()

Write-Host "Created $outFull" -ForegroundColor Green