param(
    [Parameter(Mandatory = $true)][string]$ContentPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$TemplatePath
)

$ErrorActionPreference = 'Stop'

# --- Gottfried & Somberg (GSWM) branded chart deck -------------------------
# Same GSWM_template.pptx branding as build-deck.ps1 (Wall St bull title slide,
# navy header band, GS logo, standing disclosure footer). This variant differs
# in one way: instead of leaving a "Paste chart here" drop-zone, it embeds an
# actual chart image (JSON "image" field) on the right, with a short takeaway
# bullet column on the left. Used for chart-driven decks like the market
# outlook, where the figures themselves are the point.
# ---------------------------------------------------------------------------

if (-not $TemplatePath) {
    $TemplatePath = Join-Path $PSScriptRoot 'GSWM_template.pptx'
}
if (-not (Test-Path $TemplatePath)) {
    throw "Template not found: $TemplatePath"
}

function ConvertTo-OleColor {
    param([Parameter(Mandatory = $true)][string]$Hex)
    $r = [Convert]::ToInt32($Hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($Hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($Hex.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

$BlackColor = ConvertTo-OleColor '000000'
$WhiteColor = ConvertTo-OleColor 'FFFFFF'

$msoTextOrientationHorizontal = 1
$msoTrue = -1
$msoFalse = 0
$ppSaveAsOpenXMLPresentation = 24
$msoShapeAutoSizeNone = 0
$msoAutoSizeTextToFitShape = 2

# Slide is 960 x 540 pt (16:9). Header band ~1.72in; footer holds the logo.
$BodyTopPt = 168

# Left takeaway column
$TextLeftPt = 48
$TextTopPt = $BodyTopPt
$TextWidthPt = 252
$TextHeightPt = 330

# Right chart area (image is fit to width, aspect preserved)
$ChartLeftPt = 312
$ChartMaxWidthPt = 600
$ChartMaxHeightPt = 322

$content = Get-Content -Raw -Path $ContentPath -Encoding UTF8 | ConvertFrom-Json
$contentDir = Split-Path -Parent ([System.IO.Path]::GetFullPath($ContentPath))

$outputFull = [System.IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Force -Path (Split-Path $outputFull) | Out-Null

# Build on a LOCAL temp file, never directly in the OneDrive-synced output
# folder (PowerPoint COM is unreliable against a syncing folder).
$workFile = Join-Path ([System.IO.Path]::GetTempPath()) ("deck_build_" + [System.Guid]::NewGuid().ToString('N') + ".pptx")
Copy-Item -Path $TemplatePath -Destination $workFile -Force

function Get-ShapeByName {
    param($Slide, [string]$Pattern)
    foreach ($sh in $Slide.Shapes) {
        if ($sh.Name -like $Pattern) { return $sh }
    }
    return $null
}

function Set-TitleText {
    param($Shape, [string]$Text, [int]$Size)
    $tr = $Shape.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = 'Garamond'
    $tr.Font.Size = $Size
    $tr.Font.Bold = $msoFalse
    $tr.Font.Color.RGB = $WhiteColor
    $Shape.TextFrame2.AutoSize = $msoAutoSizeTextToFitShape
}

Add-Type -AssemblyName System.Drawing

$wasAlreadyRunning = $null -ne (Get-Process POWERPNT -ErrorAction SilentlyContinue)

$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $msoTrue

    $pres = $ppt.Presentations.Open($workFile, $false, $false, $msoTrue)

    # --- Title slide -------------------------------------------------------
    $titleSlide = $pres.Slides.Item(1)
    $subtitle = Get-ShapeByName -Slide $titleSlide -Pattern 'Subtitle*'
    if ($null -eq $subtitle) { $subtitle = Get-ShapeByName -Slide $titleSlide -Pattern 'Title*' }
    if ($null -ne $subtitle) {
        Set-TitleText -Shape $subtitle -Text $content.title -Size 32
    }

    # --- Content slides ----------------------------------------------------
    $templateContent = $pres.Slides.Item(2)

    foreach ($s in $content.slides) {
        $dup = $templateContent.Duplicate()
        $dup.MoveTo($pres.Slides.Count)
        $slide = $pres.Slides.Item($pres.Slides.Count)

        $titleShape = Get-ShapeByName -Slide $slide -Pattern 'Title*'
        if ($null -ne $titleShape) {
            Set-TitleText -Shape $titleShape -Text $s.title -Size 36
        }

        # Left takeaway bullets
        $bulletText = ($s.bullets | ForEach-Object { "$([char]0x2022)  $_" }) -join "`r"
        $box = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $TextLeftPt, $TextTopPt, $TextWidthPt, $TextHeightPt)
        $box.TextFrame.WordWrap = $msoTrue
        $box.TextFrame2.AutoSize = $msoShapeAutoSizeNone
        $range = $box.TextFrame.TextRange
        $range.Text = $bulletText
        $range.Font.Name = 'Times New Roman'
        $range.Font.Size = 16
        $range.Font.Bold = $msoFalse
        $range.Font.Color.RGB = $BlackColor
        $range.ParagraphFormat.LineRuleWithin = $true
        $range.ParagraphFormat.SpaceWithin = 1.2
        $range.ParagraphFormat.SpaceAfter = 14

        # Right chart image, fit to width with aspect preserved
        if ($s.PSObject.Properties['image'] -and $s.image) {
            $imgPath = $s.image
            if (-not [System.IO.Path]::IsPathRooted($imgPath)) {
                $imgPath = Join-Path $contentDir $imgPath
            }
            if (-not (Test-Path $imgPath)) { throw "Image not found: $imgPath" }

            $img = [System.Drawing.Image]::FromFile($imgPath)
            $aspect = $img.Height / $img.Width
            $img.Dispose()

            $w = $ChartMaxWidthPt
            $h = $w * $aspect
            if ($h -gt $ChartMaxHeightPt) {
                $h = $ChartMaxHeightPt
                $w = $h / $aspect
            }
            $left = $ChartLeftPt + (($ChartMaxWidthPt - $w) / 2)
            $top = $BodyTopPt + (($ChartMaxHeightPt - $h) / 2)

            $pic = $slide.Shapes.AddPicture($imgPath, $msoFalse, $msoTrue, $left, $top, $w, $h)
            $pic.Name = 'ChartImage'
        }
    }

    # remove the leftover template content slide (now at index 2)
    $pres.Slides.Item(2).Delete()

    $pres.SaveAs($workFile, $ppSaveAsOpenXMLPresentation)
    $pres.Close()
    if (-not $wasAlreadyRunning) {
        $ppt.Quit()
    }

    Copy-Item -Path $workFile -Destination $outputFull -Force
    Write-Host "Created deck: $outputFull"
}
finally {
    if ($pres) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null }
    if ($ppt) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if (Test-Path $workFile) { Remove-Item $workFile -Force -ErrorAction SilentlyContinue }
}
