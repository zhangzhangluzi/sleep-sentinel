param(
    [string]$IconPath = (Join-Path $PSScriptRoot "..\\Assets\\AppIcon.ico"),
    [string]$PreviewPath = (Join-Path $PSScriptRoot "..\\Assets\\AppIcon-preview.png")
)

Add-Type -AssemblyName System.Drawing

function New-RoundedRectanglePath {
    param(
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [float]$Radius
    )

    $diameter = $Radius * 2
    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
    $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-ShieldPath {
    param([int]$Size)

    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $points = [System.Drawing.PointF[]]@(
        [System.Drawing.PointF]::new($Size * 0.50, $Size * 0.18),
        [System.Drawing.PointF]::new($Size * 0.74, $Size * 0.28),
        [System.Drawing.PointF]::new($Size * 0.70, $Size * 0.60),
        [System.Drawing.PointF]::new($Size * 0.50, $Size * 0.82),
        [System.Drawing.PointF]::new($Size * 0.30, $Size * 0.60),
        [System.Drawing.PointF]::new($Size * 0.26, $Size * 0.28)
    )
    $path.AddPolygon($points)
    return $path
}

function New-StarPath {
    param(
        [float]$CenterX,
        [float]$CenterY,
        [float]$OuterRadius,
        [float]$InnerRadius
    )

    $points = [System.Collections.Generic.List[System.Drawing.PointF]]::new()
    for ($i = 0; $i -lt 10; $i++) {
        $angle = (-90 + ($i * 36)) * [Math]::PI / 180
        $radius = if ($i % 2 -eq 0) { $OuterRadius } else { $InnerRadius }
        $points.Add([System.Drawing.PointF]::new(
                $CenterX + ([Math]::Cos($angle) * $radius),
                $CenterY + ([Math]::Sin($angle) * $radius)))
    }

    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $path.AddPolygon($points.ToArray())
    return $path
}

function New-IconBitmap {
    param([int]$Size)

    $bitmap = [System.Drawing.Bitmap]::new($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $graphics.Clear([System.Drawing.Color]::Transparent)

    try {
        $outerMargin = [float]($Size * 0.08)
        $outerSize = [float]($Size - ($outerMargin * 2))
        $outerPath = New-RoundedRectanglePath -X $outerMargin -Y $outerMargin -Width $outerSize -Height $outerSize -Radius ([float]($Size * 0.22))

        $backgroundBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
            [System.Drawing.PointF]::new(0, 0),
            [System.Drawing.PointF]::new($Size, $Size),
            [System.Drawing.Color]::FromArgb(255, 13, 34, 58),
            [System.Drawing.Color]::FromArgb(255, 9, 104, 118))
        $graphics.FillPath($backgroundBrush, $outerPath)

        $topGlowRect = [System.Drawing.RectangleF]::new($outerMargin, $outerMargin, $outerSize, $outerSize * 0.62)
        $topGlowBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
            [System.Drawing.PointF]::new(0, $outerMargin),
            [System.Drawing.PointF]::new(0, $outerMargin + ($outerSize * 0.62)),
            [System.Drawing.Color]::FromArgb(90, 255, 255, 255),
            [System.Drawing.Color]::FromArgb(0, 255, 255, 255))
        $graphics.FillEllipse($topGlowBrush, $topGlowRect)

        $borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(120, 227, 245, 255), [Math]::Max(1.2, $Size * 0.03))
        $graphics.DrawPath($borderPen, $outerPath)

        $shieldPath = New-ShieldPath -Size $Size
        $shadowMatrix = [System.Drawing.Drawing2D.Matrix]::new()
        $shadowMatrix.Translate($Size * 0.018, $Size * 0.03)
        $shieldShadow = $shieldPath.Clone()
        $shieldShadow.Transform($shadowMatrix)
        $graphics.FillPath([System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(56, 2, 13, 24)), $shieldShadow)

        $shieldBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
            [System.Drawing.PointF]::new($Size * 0.26, $Size * 0.18),
            [System.Drawing.PointF]::new($Size * 0.72, $Size * 0.82),
            [System.Drawing.Color]::FromArgb(255, 232, 243, 250),
            [System.Drawing.Color]::FromArgb(255, 182, 216, 230))
        $graphics.FillPath($shieldBrush, $shieldPath)

        $shieldBorder = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(180, 7, 62, 78), [Math]::Max(1.0, $Size * 0.022))
        $graphics.DrawPath($shieldBorder, $shieldPath)

        $moonBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 252, 205, 92))
        $graphics.FillEllipse($moonBrush, $Size * 0.34, $Size * 0.28, $Size * 0.24, $Size * 0.24)

        $moonCutBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 207, 231, 241))
        $graphics.FillEllipse($moonCutBrush, $Size * 0.43, $Size * 0.24, $Size * 0.22, $Size * 0.22)

        $sleepLinePen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(220, 21, 103, 119), [Math]::Max(1.6, $Size * 0.034))
        $sleepLinePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
        $sleepLinePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        $graphics.DrawLine($sleepLinePen, $Size * 0.38, $Size * 0.59, $Size * 0.62, $Size * 0.59)

        $starPath = New-StarPath -CenterX ($Size * 0.68) -CenterY ($Size * 0.34) -OuterRadius ($Size * 0.045) -InnerRadius ($Size * 0.02)
        $graphics.FillPath([System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 255, 247, 222)), $starPath)

        $graphics.FillEllipse([System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(170, 255, 245, 220)), $Size * 0.23, $Size * 0.24, $Size * 0.035, $Size * 0.035)

        $shieldShadow.Dispose()
        $shadowMatrix.Dispose()
        $starPath.Dispose()
        $sleepLinePen.Dispose()
        $moonCutBrush.Dispose()
        $moonBrush.Dispose()
        $shieldBorder.Dispose()
        $shieldBrush.Dispose()
        $borderPen.Dispose()
        $topGlowBrush.Dispose()
        $backgroundBrush.Dispose()
        $shieldPath.Dispose()
        $outerPath.Dispose()
    }
    finally {
        $graphics.Dispose()
    }

    return $bitmap
}

function Save-IconFile {
    param(
        [System.Collections.Generic.List[System.Drawing.Bitmap]]$Bitmaps,
        [string]$Path
    )

    $directory = Split-Path -Parent $Path
    if ($directory) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    $imageEntries = [System.Collections.Generic.List[object]]::new()
    foreach ($bitmap in $Bitmaps) {
        $stream = [System.IO.MemoryStream]::new()
        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        $imageEntries.Add([pscustomobject]@{
                Width  = $bitmap.Width
                Height = $bitmap.Height
                Bytes  = $stream.ToArray()
            })
        $stream.Dispose()
    }

    $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $writer = [System.IO.BinaryWriter]::new($fileStream)

    try {
        $writer.Write([UInt16]0)
        $writer.Write([UInt16]1)
        $writer.Write([UInt16]$imageEntries.Count)

        $offset = 6 + (16 * $imageEntries.Count)
        foreach ($entry in $imageEntries) {
            $widthValue = if ($entry.Width -ge 256) { 0 } else { $entry.Width }
            $heightValue = if ($entry.Height -ge 256) { 0 } else { $entry.Height }
            $writer.Write([byte]$widthValue)
            $writer.Write([byte]$heightValue)
            $writer.Write([byte]0)
            $writer.Write([byte]0)
            $writer.Write([UInt16]1)
            $writer.Write([UInt16]32)
            $writer.Write([UInt32]$entry.Bytes.Length)
            $writer.Write([UInt32]$offset)
            $offset += $entry.Bytes.Length
        }

        foreach ($entry in $imageEntries) {
            $writer.Write($entry.Bytes)
        }
    }
    finally {
        $writer.Dispose()
        $fileStream.Dispose()
    }
}

$sizes = 16, 20, 24, 32, 40, 48, 64, 128, 256
$bitmaps = [System.Collections.Generic.List[System.Drawing.Bitmap]]::new()

try {
    foreach ($size in $sizes) {
        $bitmaps.Add((New-IconBitmap -Size $size))
    }

    Save-IconFile -Bitmaps $bitmaps -Path $IconPath

    $previewBitmap = New-IconBitmap -Size 512
    try {
        $previewDirectory = Split-Path -Parent $PreviewPath
        if ($previewDirectory) {
            New-Item -ItemType Directory -Force -Path $previewDirectory | Out-Null
        }

        $previewBitmap.Save($PreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $previewBitmap.Dispose()
    }
}
finally {
    foreach ($bitmap in $bitmaps) {
        $bitmap.Dispose()
    }
}
