param(
    [string]$SourcePng = "",
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

function New-ScaledBitmap {
    param(
        [System.Drawing.Image]$Source,
        [int]$Size
    )

    $bmp = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $g.Clear([System.Drawing.Color]::Transparent)
        $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.DrawImage($Source, 0, 0, $Size, $Size)
    }
    finally {
        $g.Dispose()
    }

    return $bmp
}

function New-RoundBitmap {
    param(
        [System.Drawing.Image]$Source,
        [int]$Size,
        [double]$InsetRatio = 0.0125,
        [int]$Supersample = 8
    )

    $ss = [Math]::Max(2, [Math]::Min(16, $Supersample))
    $ssSize = $Size * $ss
    $high = New-Object System.Drawing.Bitmap $ssSize, $ssSize, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $gHigh = [System.Drawing.Graphics]::FromImage($high)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    try {
        $gHigh.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $gHigh.Clear([System.Drawing.Color]::Transparent)
        $gHigh.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
        $gHigh.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $gHigh.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $gHigh.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $gHigh.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::Half

        # Keep circle edges inside bounds and reserve space for anti-aliased pixels.
        $ratioInsetPx = [Math]::Max(0.0, [Math]::Min(0.45, $InsetRatio)) * $Size
        $insetPx = [Math]::Max(1.0, $ratioInsetPx)
        $insetHi = [float]($insetPx * $ss)
        $diameterHi = [float]$ssSize - (2.0 * $insetHi)
        if ($diameterHi -lt 1.0) { $diameterHi = 1.0 }
        $path.AddEllipse($insetHi, $insetHi, $diameterHi, $diameterHi)
        $gHigh.SetClip($path)
        $gHigh.DrawImage($Source, 0, 0, $ssSize, $ssSize)
        $gHigh.ResetClip()

        $out = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $gOut = [System.Drawing.Graphics]::FromImage($out)
        $imgAttr = New-Object System.Drawing.Imaging.ImageAttributes
        try {
            $gOut.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
            $gOut.Clear([System.Drawing.Color]::Transparent)
            $gOut.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
            $gOut.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $gOut.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $gOut.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $gOut.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $imgAttr.SetWrapMode([System.Drawing.Drawing2D.WrapMode]::TileFlipXY)
            $destRect = New-Object System.Drawing.Rectangle 0, 0, $Size, $Size
            $gOut.DrawImage($high, $destRect, 0, 0, $ssSize, $ssSize, [System.Drawing.GraphicsUnit]::Pixel, $imgAttr)
        }
        finally {
            $imgAttr.Dispose()
            $gOut.Dispose()
        }

        return $out
    }
    finally {
        $path.Dispose()
        $gHigh.Dispose()
        $high.Dispose()
    }
}

function Get-PngBytes {
    param([System.Drawing.Bitmap]$Bitmap)

    $ms = New-Object System.IO.MemoryStream
    try {
        $Bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        return $ms.ToArray()
    }
    finally {
        $ms.Dispose()
    }
}

function Write-Ico {
    param(
        [System.Drawing.Image]$Source,
        [int[]]$Sizes,
        [string]$OutputPath,
        [switch]$Round
    )

    $images = @()
    $roundMaster = $null
    try {
        if ($Round) {
            $maxSize = [int](($Sizes | Measure-Object -Maximum).Maximum)
            $masterSize = [Math]::Max(1024, $maxSize * 4)
            $roundMaster = New-RoundBitmap -Source $Source -Size $masterSize -InsetRatio 0.0125 -Supersample 8
        }

        foreach ($size in $Sizes) {
            $bmp = if ($Round) { New-ScaledBitmap -Source $roundMaster -Size $size } else { New-ScaledBitmap -Source $Source -Size $size }
            try {
                $bytes = Get-PngBytes -Bitmap $bmp
                $images += [pscustomobject]@{
                    Size  = $size
                    Bytes = $bytes
                }
            }
            finally {
                $bmp.Dispose()
            }
        }
    }
    finally {
        if ($null -ne $roundMaster) {
            $roundMaster.Dispose()
        }
    }

    $dir = Split-Path -Path $OutputPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $fs = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $bw = New-Object System.IO.BinaryWriter $fs
    try {
        $count = $images.Count
        $bw.Write([UInt16]0) # reserved
        $bw.Write([UInt16]1) # type: icon
        $bw.Write([UInt16]$count)

        $offset = 6 + (16 * $count)
        foreach ($img in $images) {
            $size = [int]$img.Size
            $bytes = [byte[]]$img.Bytes
            $dim = if ($size -ge 256) { [byte]0 } else { [byte]$size }

            $bw.Write($dim)                # width
            $bw.Write($dim)                # height
            $bw.Write([byte]0)             # color count
            $bw.Write([byte]0)             # reserved
            $bw.Write([UInt16]1)           # color planes
            $bw.Write([UInt16]32)          # bpp
            $bw.Write([UInt32]$bytes.Length)
            $bw.Write([UInt32]$offset)
            $offset += $bytes.Length
        }

        foreach ($img in $images) {
            $bw.Write([byte[]]$img.Bytes)
        }
    }
    finally {
        $bw.Dispose()
        $fs.Dispose()
    }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($SourcePng)) {
    $SourcePng = Join-Path $repoRoot "docs\logo2.png"
}

$sourcePath = (Resolve-Path $SourcePng).Path
if (-not (Test-Path $sourcePath)) {
    throw "Source image not found: $SourcePng"
}

$toolIco = Join-Path $repoRoot "src\PbiMetadataTool\app.ico"
$installerIco = Join-Path $repoRoot "src\PbiMetadataInstaller\app.ico"
$roundPng = Join-Path $repoRoot "external-tools\icon-round.png"
$roundIco = Join-Path $repoRoot "external-tools\PBIClaw.round.ico"
$pbiToolJson = Join-Path $repoRoot "external-tools\PBIClaw.pbitool.json"

$iconSizes = @(16, 20, 24, 32, 40, 48, 64, 96, 128, 256)

$source = [System.Drawing.Image]::FromFile($sourcePath)
try {
    Write-Ico -Source $source -Sizes $iconSizes -OutputPath $toolIco
    Write-Ico -Source $source -Sizes $iconSizes -OutputPath $installerIco
    Write-Ico -Source $source -Sizes $iconSizes -OutputPath $roundIco -Round

    $roundMaster = New-RoundBitmap -Source $source -Size 1024 -InsetRatio 0.0125 -Supersample 8
    try {
        $round256 = New-ScaledBitmap -Source $roundMaster -Size 256
        try {
            $round256.Save($roundPng, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $round256.Dispose()
        }
    }
    finally {
        $roundMaster.Dispose()
    }
}
finally {
    $source.Dispose()
}

if (Test-Path $pbiToolJson) {
    $json = Get-Content $pbiToolJson -Raw | ConvertFrom-Json
    $json.iconData = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($roundPng))
    $json | ConvertTo-Json -Depth 20 | Set-Content $pbiToolJson -Encoding UTF8
}

if (-not $Quiet) {
    Write-Host "Generated:"
    Write-Host "  - $toolIco"
    Write-Host "  - $installerIco"
    Write-Host "  - $roundIco"
    Write-Host "  - $roundPng"
    Write-Host "  - $pbiToolJson (iconData)"
}
