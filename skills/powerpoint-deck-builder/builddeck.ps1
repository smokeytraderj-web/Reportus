param(
    [Parameter(Mandatory = $true)][string]$ContentPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$TemplatePath
)

$ErrorActionPreference = 'Stop'

# --- Gottfried & Somberg (GSWM) branded template ---------------------------
# Every deck is built on GSWM_template.pptx, which carries the firm branding:
#   Slide 1  = title slide  (Wall St bull image, GS circular logo, navy band)
#   Slide 2  = content slide (navy gradient header band, white serif title,
#              GS logo watermark bottom-left, standing disclosure footer)
# This script fills the title slide, then duplicates the branded content slide
# once per entry in the JSON and drops the bullet text into the body area.
# Do not draw the branding by hand here -- it lives in the template so it stays
# pixel-identical to the firm's real weekly decks.
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

# Body layout in points (slide is 960 x 540 pt / 13.333in x 7.5in, 16:9).
# Header band is ~1.72in tall; footer + logo occupy the bottom strip.
$msoTextOrientationHorizontal = 1
$msoTrue = -1
$msoFalse = 0
$ppSaveAsOpenXMLPresentation = 24
$msoShapeAutoSizeNone = 0
$msoAutoSizeTextToFitShape = 2
$msoShapeRectangle = 1
$ppAlignCenter = 2
$ppAnchorTop = 1
$ppAnchorMiddle = 3

$GrayFill = ConvertTo-OleColor 'F7F7F7'   # subtle drop-zone fill
$GrayLine = ConvertTo-OleColor 'BFBFBF'   # placeholder border
$GrayText = ConvertTo-OleColor '909090'   # placeholder caption

$BodyLeftPt = 54      # 0.75in
$BodyTopPt = 150      # just below the navy header band
$BodyWidthPt = 852    # 960 - 2*54
$BodyHeightPt = 322   # stops above the footer strip

# Chart-slide layout: bullets in a left column, chart image on the right.
# Text column narrowed and the image area widened/enlarged per firm request.
$ChartTextWidthPt = 330   # left bullet column (54 -> 384)
$ChartTextHeightPt = 285  # shorter than full body so the column clears the logo
$ChartAreaLeftPt = 404    # gap, then chart image
$ChartAreaTopPt = 145
$ChartAreaWidthPt = 502   # 404 -> 906
$ChartAreaHeightPt = 320

$content = Get-Content -Raw -Path $ContentPath -Encoding UTF8 | ConvertFrom-Json

$outputFull = [System.IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Force -Path (Split-Path $outputFull) | Out-Null

# Build on a LOCAL temp file, never directly in the (OneDrive-synced) output
# folder. PowerPoint COM automation is unreliable against a syncing folder --
# multi-step edits intermittently fail with HRESULT 0x80CA1007 -- so we do all
# the work locally and copy the finished deck to the output path at the end.
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
    # shrink text if a long title would spill past the band
    $Shape.TextFrame2.AutoSize = $msoAutoSizeTextToFitShape
}

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
        $titleLinesProp = $content.PSObject.Properties['title_lines']
        if ($null -ne $titleLinesProp -and $titleLinesProp.Value) {
            # Multi-line title block: name, review label, small muted caption --
            # each line its own paragraph with its own size/color.
            $lines = @($titleLinesProp.Value)
            $joined = ($lines | ForEach-Object { [string]$_.text }) -join "`r"
            $tr = $subtitle.TextFrame.TextRange
            $tr.Text = $joined
            $subtitle.TextFrame2.AutoSize = $msoShapeAutoSizeNone
            $MutedColor = ConvertTo-OleColor 'A9B4CC'
            for ($li = 0; $li -lt $lines.Count; $li++) {
                $p = $tr.Paragraphs($li + 1, 1)
                $p.Font.Name = 'Garamond'
                $p.Font.Bold = $msoFalse
                $sizeProp = $lines[$li].PSObject.Properties['size']
                $p.Font.Size = if ($sizeProp -and $sizeProp.Value) { [int]$sizeProp.Value } else { 24 }
                $mutedProp = $lines[$li].PSObject.Properties['muted']
                if ($null -ne $mutedProp -and $mutedProp.Value) {
                    $p.Font.Color.RGB = $MutedColor
                } else {
                    $p.Font.Color.RGB = $WhiteColor
                }
                # Title block is centered on the full-width slide (matches the
                # reference deck's title slide), not left-aligned in a narrow box.
                $p.ParagraphFormat.Alignment = $ppAlignCenter
            }
        } else {
            Set-TitleText -Shape $subtitle -Text $content.title -Size 32
        }
    }

    # --- Content slides ----------------------------------------------------
    # Slide 2 is the branded blank; duplicate it once per entry, fill, then
    # remove the original template slide at the end.
    $templateContent = $pres.Slides.Item(2)

    foreach ($s in $content.slides) {
        $dup = $templateContent.Duplicate()
        $slide = $pres.Slides.Item($pres.Slides.Count) # duplicate lands last-ish; grab by move
        $dup.MoveTo($pres.Slides.Count)
        $slide = $pres.Slides.Item($pres.Slides.Count)

        $titleShape = Get-ShapeByName -Slide $slide -Pattern 'Title*'
        if ($null -ne $titleShape) {
            Set-TitleText -Shape $titleShape -Text $s.title -Size 40
        }

        # Optional small right-aligned metadata in the header band (e.g. the
        # return-measurement window and the benchmark return), matching the
        # firm's "Review" deck format. Up to two short lines.
        $headerRightProp = $s.PSObject.Properties['header_right']
        if ($null -ne $headerRightProp -and $headerRightProp.Value) {
            $hrLines = @($headerRightProp.Value)
            $hrBox = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, 560, 28, 346, 60)
            $hrBox.TextFrame.WordWrap = $msoTrue
            $hrBox.TextFrame2.AutoSize = $msoShapeAutoSizeNone
            $hrTr = $hrBox.TextFrame.TextRange
            $hrTr.Text = ($hrLines -join "`r")
            $MutedHeaderColor = ConvertTo-OleColor 'C7CEDD'
            for ($li = 0; $li -lt $hrLines.Count; $li++) {
                $p = $hrTr.Paragraphs($li + 1, 1)
                $p.Font.Name = 'Times New Roman'
                $p.Font.Size = 12
                $p.Font.Color.RGB = $MutedHeaderColor
                $p.ParagraphFormat.Alignment = 3 # ppAlignRight
            }
        }

        # Optional footer line pair (bottom-left brand line, bottom-right
        # source/basis note), matching the firm's "Review" deck format.
        $footerLeftProp = $s.PSObject.Properties['footer_left']
        $footerRightProp = $s.PSObject.Properties['footer_right']
        if (($null -ne $footerLeftProp -and $footerLeftProp.Value) -or ($null -ne $footerRightProp -and $footerRightProp.Value)) {
            $FooterColor = ConvertTo-OleColor '8C8C8C'
            if ($null -ne $footerLeftProp -and $footerLeftProp.Value) {
                $flBox = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, 74, 505, 400, 20)
                $flBox.TextFrame.WordWrap = $msoFalse
                $flTr = $flBox.TextFrame.TextRange
                $flTr.Text = [string]$footerLeftProp.Value
                $flTr.Font.Name = 'Times New Roman'
                $flTr.Font.Size = 9
                $flTr.Font.Color.RGB = $FooterColor
            }
            if ($null -ne $footerRightProp -and $footerRightProp.Value) {
                $frBox = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, 560, 505, 346, 20)
                $frBox.TextFrame.WordWrap = $msoFalse
                $frTr = $frBox.TextFrame.TextRange
                $frTr.Text = [string]$footerRightProp.Value
                $frTr.Font.Name = 'Times New Roman'
                $frTr.Font.Size = 9
                $frTr.Font.Color.RGB = $FooterColor
                $frTr.ParagraphFormat.Alignment = 3 # ppAlignRight
            }
        }

        # A "chart_full" slide (JSON entry has a "footnote") shows the chart
        # large and centered, with a single footnote line beneath it and no
        # bullet column. This is the clean, executive one-chart-per-slide look.
        $footnote = $null
        $fnProp = $s.PSObject.Properties['footnote']
        if ($null -ne $fnProp -and $fnProp.Value) { $footnote = [string]$fnProp.Value }

        if ($footnote) {
            $FootColor = ConvertTo-OleColor '595959'
            $chartImg = $null
            $ciProp = $s.PSObject.Properties['chart_image']
            if ($null -ne $ciProp -and $ciProp.Value) { $chartImg = [string]$ciProp.Value }

            $FullLeft = 80; $FullTop = 138; $FullWidth = 800; $FullHeight = 286
            if ($chartImg -and (Test-Path $chartImg)) {
                $pic = $slide.Shapes.AddPicture($chartImg, $msoFalse, $msoTrue, $FullLeft, $FullTop, -1, -1)
                $pic.LockAspectRatio = $msoTrue
                $scale = [Math]::Min($FullWidth / $pic.Width, $FullHeight / $pic.Height)
                $pic.Width = [single]($pic.Width * $scale)
                $pic.Left = [single]($FullLeft + ($FullWidth - $pic.Width) / 2)
                $pic.Top = [single]($FullTop + ($FullHeight - $pic.Height) / 2)
                $pic.Name = 'ChartImage'
            }

            # footnote line beneath the chart (indented to clear the corner logo)
            $fnBox = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, 138, 432, 768, 70)
            $fnBox.TextFrame.WordWrap = $msoTrue
            $fnBox.TextFrame2.AutoSize = $msoShapeAutoSizeNone
            $fnr = $fnBox.TextFrame.TextRange
            $fnr.Text = $footnote
            $fnr.Font.Name = 'Times New Roman'
            $fnr.Font.Size = 11
            $fnr.Font.Italic = $msoTrue
            $fnr.Font.Color.RGB = $FootColor
            $fnr.ParagraphFormat.SpaceWithin = 1.1
            continue
        }

        # A "table" slide (JSON entry has a "table" object) renders a native
        # PowerPoint table instead of bullets -- used for data recap decks
        # (e.g. stock recommendation performance) rather than talking points.
        $tableProp = $s.PSObject.Properties['table']
        if ($null -ne $tableProp -and $tableProp.Value) {
            $tbl = $tableProp.Value
            $headers = @($tbl.headers)
            $dataRows = @($tbl.rows)
            $colorCol = -1
            $ccProp = $tbl.PSObject.Properties['color_column']
            if ($null -ne $ccProp -and $null -ne $ccProp.Value) { $colorCol = [int]$ccProp.Value }

            $numCols = $headers.Count
            $numRows = $dataRows.Count + 1

            # Font size scales down as row count rises so the table clears the
            # header band (150pt) and stays above the logo watermark (~452pt).
            if ($numRows -le 8) { $fs = 15 }
            elseif ($numRows -le 12) { $fs = 13 }
            elseif ($numRows -le 16) { $fs = 11 }
            else { $fs = 10 }

            $TableTop = 145
            $TableHeight = 300
            $graphicFrame = $slide.Shapes.AddTable($numRows, $numCols, $BodyLeftPt, $TableTop, $BodyWidthPt, $TableHeight)
            $ppTable = $graphicFrame.Table
            # AddTable's default table style bands rows and draws grid lines
            # that our own Fill/Border settings can't fully override -- swap
            # to the built-in "No Style, No Grid" style first so only our
            # explicit per-cell formatting shows.
            try { $ppTable.ApplyStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307C}'); $ppTable.HorizBanding = $msoFalse; $ppTable.VertBanding = $msoFalse } catch {}

            # Optional per-column width weights (fractions summing to ~1.0);
            # falls back to equal widths if not supplied.
            $colWidthsProp = $tbl.PSObject.Properties['col_widths']
            if ($null -ne $colWidthsProp -and $colWidthsProp.Value) {
                $weights = @($colWidthsProp.Value)
                for ($c = 1; $c -le $numCols; $c++) {
                    $ppTable.Columns.Item($c).Width = [single]($BodyWidthPt * [double]$weights[$c - 1])
                }
            }

            $NavyColor = ConvertTo-OleColor '1B2A4A'
            $GreenColor = ConvertTo-OleColor '0B7A32'
            $RedColor = ConvertTo-OleColor 'B00020'
            $LightRowColor = ConvertTo-OleColor 'F7F7F7'

            for ($c = 1; $c -le $numCols; $c++) {
                $cell = $ppTable.Cell(1, $c)
                $cell.Shape.TextFrame.VerticalAnchor = $ppAnchorMiddle
                $cell.Shape.TextFrame.MarginTop = 2
                $cell.Shape.TextFrame.MarginBottom = 2
                $cell.Shape.Fill.ForeColor.RGB = $NavyColor
                $tr = $cell.Shape.TextFrame.TextRange
                $tr.Text = [string]$headers[$c - 1]
                $tr.Font.Name = 'Times New Roman'
                $tr.Font.Bold = $msoTrue
                $tr.Font.Size = $fs
                $tr.Font.Color.RGB = $WhiteColor
            }

            for ($r = 0; $r -lt $dataRows.Count; $r++) {
                $rowVals = @($dataRows[$r])
                for ($c = 1; $c -le $numCols; $c++) {
                    $cell = $ppTable.Cell($r + 2, $c)
                    $cell.Shape.TextFrame.VerticalAnchor = $ppAnchorMiddle
                    $cell.Shape.TextFrame.MarginTop = 2
                    $cell.Shape.TextFrame.MarginBottom = 2
                    $cell.Shape.Fill.ForeColor.RGB = if ($r % 2 -eq 0) { $WhiteColor } else { $LightRowColor }
                    $tr = $cell.Shape.TextFrame.TextRange
                    $cellText = [string]$rowVals[$c - 1]
                    $tr.Text = $cellText
                    $tr.Font.Name = 'Times New Roman'
                    $tr.Font.Bold = $msoFalse
                    $tr.Font.Size = $fs
                    if (($c - 1) -eq $colorCol -and $cellText.StartsWith('-')) {
                        $tr.Font.Color.RGB = $RedColor
                    } elseif (($c - 1) -eq $colorCol) {
                        $tr.Font.Color.RGB = $GreenColor
                    } else {
                        $tr.Font.Color.RGB = $BlackColor
                    }
                }
            }
            continue
        }

        # A "picks_table" slide renders the firm's "Review" deck format: a
        # category banner row (merged, light fill) followed by stock rows
        # with a two-line Name cell (ticker bold + company gray beneath) and
        # a colored Rating pill, plus Return/vs.Benchmark columns. No Price
        # or Notes column -- dropped per firm request (doesn't apply to a
        # fund holdings review the way it does to a stock pick).
        $ptProp = $s.PSObject.Properties['picks_table']
        if ($null -ne $ptProp -and $ptProp.Value) {
            $prows = @($ptProp.Value.rows)
            $numCols = 4
            $numRows = $prows.Count + 1
            # Name only needs enough room for the longest company line (~210pt
            # at 10pt Times New Roman); the old 0.50 weight left a huge dead
            # gap before the Rating column. Tightened so the row reads as one
            # organized line instead of text stranded on the left.
            $colWeights = @(0.36, 0.16, 0.20, 0.28)
            $headersP = @('Name', 'Rating', 'Return', 'vs. Benchmark')

            # Table left/width measured off the reference PDF (960x540pt slide):
            # table spans x=45.1pt to x=914.6pt, i.e. ~45pt margins each side,
            # narrower than the shared BodyLeftPt/BodyWidthPt used by bullet
            # slides -- kept local to this block so it doesn't affect them.
            $PicksLeftPt = 45
            $PicksWidthPt = 870
            $TableTop = 145
            # Row heights are set per-row-type after the table is built (banner
            # rows are one line, stock rows are two) rather than forcing a fixed
            # total height divided evenly across every row -- that uniform
            # division was stretching short banner rows to match tall stock
            # rows, making the table look loosely/awkwardly spread out.
            $HeaderRowH = 30
            $BannerRowH = 26
            $StockRowH = 36
            $TableHeight = $HeaderRowH + ($BannerRowH * (@($prows | Where-Object { [string]$_.type -eq 'banner' })).Count) + ($StockRowH * (@($prows | Where-Object { [string]$_.type -ne 'banner' })).Count)
            $graphicFrame = $slide.Shapes.AddTable($numRows, $numCols, $PicksLeftPt, $TableTop, $PicksWidthPt, $TableHeight)
            $ppTable = $graphicFrame.Table
            try { $ppTable.ApplyStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307C}'); $ppTable.HorizBanding = $msoFalse; $ppTable.VertBanding = $msoFalse } catch {}
            for ($c = 1; $c -le $numCols; $c++) {
                $ppTable.Columns.Item($c).Width = [single]($PicksWidthPt * $colWeights[$c - 1])
            }

            # Exact hex values sampled from the reference PDF (not eyeballed):
            # header navy 16253F, banner fill E7ECF2, alt-row F5F7FA.
            $NavyColor = ConvertTo-OleColor '16253F'
            $BannerFill = ConvertTo-OleColor 'E7ECF2'
            $BannerText = ConvertTo-OleColor '16253F'
            $LightRowColor = ConvertTo-OleColor 'F5F7FA'
            $GreenColor = ConvertTo-OleColor '0B7A32'
            $RedColor = ConvertTo-OleColor 'B00020'
            $ppAlignRight = 3

            for ($c = 1; $c -le $numCols; $c++) {
                $cell = $ppTable.Cell(1, $c)
                $cell.Shape.TextFrame.VerticalAnchor = $ppAnchorMiddle
                $cell.Shape.TextFrame.MarginTop = 2
                $cell.Shape.TextFrame.MarginBottom = 2
                $cell.Shape.Fill.ForeColor.RGB = $NavyColor
                $tr = $cell.Shape.TextFrame.TextRange
                $tr.Text = $headersP[$c - 1]
                $tr.Font.Name = 'Times New Roman'
                $tr.Font.Bold = $msoTrue
                $tr.Font.Size = 12
                $tr.Font.Color.RGB = $WhiteColor
                if ($c -ge 3) { $tr.ParagraphFormat.Alignment = $ppAlignRight }
            }
            $ppTable.Rows.Item(1).Height = $HeaderRowH

            $rowIdx = 2
            $ratingBadges = @()
            foreach ($row in $prows) {
                if ([string]$row.type -eq 'banner') {
                    $ppTable.Cell($rowIdx, 1).Merge($ppTable.Cell($rowIdx, $numCols))
                    $mcell = $ppTable.Cell($rowIdx, 1)
                    $mcell.Shape.TextFrame.VerticalAnchor = $ppAnchorMiddle
                    $mcell.Shape.TextFrame.MarginTop = 2
                    $mcell.Shape.TextFrame.MarginBottom = 2
                    $mcell.Shape.TextFrame.MarginLeft = 10
                    $mcell.Shape.Fill.ForeColor.RGB = $BannerFill
                    $mtr = $mcell.Shape.TextFrame.TextRange
                    $mtr.Text = [string]$row.text
                    $mtr.Font.Name = 'Times New Roman'
                    $mtr.Font.Bold = $msoTrue
                    $mtr.Font.Size = 10.5
                    $mtr.Font.Color.RGB = $BannerText
                    $ppTable.Rows.Item($rowIdx).Height = $BannerRowH
                    $rowIdx++
                    continue
                }

                $bg = if (($rowIdx % 2) -eq 0) { $WhiteColor } else { $LightRowColor }

                # Name: ticker bold on top, company name smaller/gray below.
                # Top-anchored (not middle) so the ticker line always starts at
                # a fixed offset from the row top -- the rating badge below is
                # positioned relative to that same offset, so this keeps the
                # badge correctly aligned regardless of row height.
                $nameCell = $ppTable.Cell($rowIdx, 1)
                $nameCell.Shape.TextFrame.VerticalAnchor = $ppAnchorTop
                $nameCell.Shape.TextFrame.MarginTop = 4
                $nameCell.Shape.TextFrame.MarginBottom = 4
                $nameCell.Shape.Fill.ForeColor.RGB = $bg
                $ntr = $nameCell.Shape.TextFrame.TextRange
                $ntr.Text = "$([string]$row.ticker)`r$([string]$row.company)"
                $p1 = $ntr.Paragraphs(1, 1)
                $p1.Font.Name = 'Times New Roman'; $p1.Font.Bold = $msoTrue; $p1.Font.Size = 13; $p1.Font.Color.RGB = $BlackColor
                $p2 = $ntr.Paragraphs(2, 1)
                $p2.Font.Name = 'Times New Roman'; $p2.Font.Bold = $msoFalse; $p2.Font.Size = 10; $p2.Font.Color.RGB = $GrayText

                # Rating: the real template renders this as a small pill badge
                # floating in the middle of the cell (padding visible on all
                # sides), NOT a color fill spanning the whole cell. Leave the
                # cell itself as plain row background and queue the badge to
                # be drawn as a separate AutoShape once row heights are final
                # (row heights are still provisional mid-loop).
                switch ([string]$row.rating_color) {
                    'green' { $fillHex = 'E3F2E6'; $textHex = '1B7A3D' }
                    'amber' { $fillHex = 'FCEFD8'; $textHex = 'B36B00' }
                    'red' { $fillHex = 'FBE1E1'; $textHex = 'A81C1C' }
                    default { $fillHex = 'ECECEC'; $textHex = '555555' }
                }
                $ratCell = $ppTable.Cell($rowIdx, 2)
                $ratCell.Shape.TextFrame.VerticalAnchor = $ppAnchorMiddle
                $ratCell.Shape.Fill.ForeColor.RGB = $bg
                $ratingBadges += [PSCustomObject]@{
                    Row = $rowIdx
                    Text = [string]$row.rating
                    Fill = $fillHex
                    TextColor = $textHex
                }

                # Return / vs. Benchmark -- right aligned, bold, colored.
                $plainCols = @(3, 4)
                $plainVals = @([string]$row.return, [string]$row.vs_sp)
                for ($i = 0; $i -lt 2; $i++) {
                    $c = $plainCols[$i]
                    $cell = $ppTable.Cell($rowIdx, $c)
                    $cell.Shape.TextFrame.VerticalAnchor = $ppAnchorMiddle
                    $cell.Shape.TextFrame.MarginTop = 4
                    $cell.Shape.TextFrame.MarginBottom = 4
                    $cell.Shape.TextFrame.MarginRight = 10
                    $cell.Shape.Fill.ForeColor.RGB = $bg
                    $tr2 = $cell.Shape.TextFrame.TextRange
                    $val = $plainVals[$i]
                    $tr2.Text = $val
                    $tr2.Font.Name = 'Times New Roman'
                    $tr2.Font.Size = 12
                    $tr2.Font.Bold = $msoTrue
                    $tr2.ParagraphFormat.Alignment = $ppAlignRight
                    if ($val -eq 'n/a') { $tr2.Font.Color.RGB = $GrayText }
                    elseif ($val.StartsWith('-')) { $tr2.Font.Color.RGB = $RedColor }
                    else { $tr2.Font.Color.RGB = $GreenColor }
                }

                $ppTable.Rows.Item($rowIdx).Height = $StockRowH
                $rowIdx++
            }

            # AddTable applies a default PowerPoint table style with visible
            # white gridlines between cells; the reference has none. Strip
            # every cell border explicitly rather than fight the style.
            for ($r = 1; $r -le $numRows; $r++) {
                for ($c = 1; $c -le $numCols; $c++) {
                    foreach ($bIdx in 1, 2, 3, 4) {
                        try { $ppTable.Cell($r, $c).Borders.Item($bIdx).Visible = $msoFalse } catch {}
                    }
                }
            }

            # Second pass: draw the rating pill badges now that every row's
            # final height is known. Auto-size each badge to hug its text
            # (tight padding, like the real template) then center it in the
            # cell -- a fixed-size box would either clip long tags ("Mkt
            # Neutral") or float tiny ones ("Buy") off-center.
            $msoShapeRoundedRectangle = 5
            $msoAutoSizeShapeToFitText = 1
            $col1Width = $ppTable.Columns.Item(1).Width
            $col2Width = $ppTable.Columns.Item(2).Width
            foreach ($badge in $ratingBadges) {
                $rowTop = $graphicFrame.Top
                for ($r = 1; $r -lt $badge.Row; $r++) {
                    $rowTop += $ppTable.Rows.Item($r).Height
                }
                $rowHeight = $ppTable.Rows.Item($badge.Row).Height
                $cellLeft = $graphicFrame.Left + $col1Width

                $shp = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $cellLeft, $rowTop, 70, 22)
                $shp.Adjustments.Item(1) = 0.3
                $shp.Fill.ForeColor.RGB = ConvertTo-OleColor $badge.Fill
                $shp.Line.ForeColor.RGB = ConvertTo-OleColor $badge.TextColor
                $shp.Line.Weight = 0.75
                $shp.Shadow.Visible = $msoFalse
                $shp.TextFrame.MarginLeft = 6
                $shp.TextFrame.MarginRight = 6
                $shp.TextFrame.MarginTop = 1
                $shp.TextFrame.MarginBottom = 1
                # WordWrap must be off before AutoSize is applied, otherwise a
                # two-word tag like "Mkt Neutral" wraps to two lines inside the
                # still-70pt-wide box, doubling the badge's height so it floats
                # up out of its row into the row above. With wrap off, AutoSize
                # grows the box wider instead, keeping every badge one line tall.
                $shp.TextFrame.WordWrap = $msoFalse
                $shp.TextFrame.AutoSize = $msoAutoSizeShapeToFitText
                $btr = $shp.TextFrame.TextRange
                $btr.Text = $badge.Text
                $btr.Font.Name = 'Times New Roman'
                $btr.Font.Bold = $msoTrue
                $btr.Font.Size = 11
                $btr.Font.Color.RGB = ConvertTo-OleColor $badge.TextColor
                $btr.ParagraphFormat.Alignment = $ppAlignCenter

                # Re-center horizontally now that AutoSize has resolved the
                # actual width. Vertically, the reference aligns the badge
                # with the ticker line (top of the 2-line Name cell), not the
                # midpoint of the whole row -- center it on the cell's top
                # margin (4pt) plus half the ticker line's height (~15.5pt
                # for 13pt bold Times New Roman at single spacing).
                $shp.Left = [single]($cellLeft + ($col2Width - $shp.Width) / 2)
                $tickerLineCenter = $rowTop + 4 + 7.75
                $shp.Top = [single]($tickerLineCenter - $shp.Height / 2)
            }

            continue
        }

        # A slide is a "chart slide" when the JSON entry sets "chart": true.
        # When the notes reference an on-screen chart/figure, mark that slide so
        # bullets move to a left column and a blank drop-zone is left on the
        # right for the user to paste the chart into.
        $isChart = $false
        $chartProp = $s.PSObject.Properties['chart']
        if ($null -ne $chartProp -and $chartProp.Value) { $isChart = $true }

        if ($isChart) {
            $boxWidth = $ChartTextWidthPt
            $boxHeight = $ChartTextHeightPt
        } else {
            $boxWidth = $BodyWidthPt
            $boxHeight = $BodyHeightPt
        }

        $bulletText = ($s.bullets | ForEach-Object { "$([char]0x2022)  $_" }) -join "`r"
        $box = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $BodyLeftPt, $BodyTopPt, $boxWidth, $boxHeight)
        $box.TextFrame.WordWrap = $msoTrue
        $box.Height = $boxHeight
        $range = $box.TextFrame.TextRange
        $range.Text = $bulletText
        $range.Font.Name = 'Times New Roman'
        $range.Font.Bold = $msoFalse
        $range.Font.Color.RGB = $BlackColor
        $range.ParagraphFormat.LineRuleWithin = $true
        $range.ParagraphFormat.SpaceWithin = 1.5

        if ($isChart) {
            # Narrow column beside the chart: keep the compact 16pt body and
            # shrink to fit so it never overflows into the chart drop-zone.
            $box.TextFrame2.AutoSize = $msoAutoSizeTextToFitShape
            $range.Font.Size = 16
            $range.ParagraphFormat.SpaceAfter = 18
        } else {
            # Full-width text slide (no chart): enlarge the type and spread the
            # bullets so they fill the body, but keep the whole block inside a
            # region that STOPS above the GS logo watermark in the bottom-left
            # corner (the logo spans ~452 to 532pt). We middle-anchor an ~286pt
            # region (150 -> 436) so the block sits below the header and, even if
            # a bullet wraps, it grows both ways and never collides with the logo.
            # Font and gap scale down as bullet count rises so a busy slide fits.
            $box.Top = 150
            $box.Height = 286
            $box.TextFrame.VerticalAnchor = $ppAnchorMiddle
            $box.TextFrame2.AutoSize = $msoShapeAutoSizeNone
            $n = @($s.bullets).Count
            switch ($n) {
                { $_ -le 3 } { $fs = 28; $sa = 60; break }
                4 { $fs = 24; $sa = 36; break }
                5 { $fs = 22; $sa = 22; break }
                6 { $fs = 20; $sa = 14; break }
                default { $fs = 18; $sa = 10 }
            }
            $range.Font.Size = $fs
            $range.ParagraphFormat.SpaceAfter = $sa
        }

        if ($isChart) {
            # If the JSON entry supplies "chart_image", drop that image straight
            # into the chart zone (scaled to fit, centered). Otherwise leave the
            # blank "Paste chart here" drop-zone as before.
            $chartImg = $null
            $ciProp = $s.PSObject.Properties['chart_image']
            if ($null -ne $ciProp -and $ciProp.Value) { $chartImg = [string]$ciProp.Value }

            if ($chartImg -and (Test-Path $chartImg)) {
                # AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
                # Width/Height = -1 keeps native pixel size; we then scale to fit.
                $pic = $slide.Shapes.AddPicture($chartImg, $msoFalse, $msoTrue, $ChartAreaLeftPt, $ChartAreaTopPt, -1, -1)
                $pic.LockAspectRatio = $msoTrue
                $scale = [Math]::Min($ChartAreaWidthPt / $pic.Width, $ChartAreaHeightPt / $pic.Height)
                # Setting Width with aspect ratio locked scales Height to match.
                $pic.Width = [single]($pic.Width * $scale)
                # center the image within the chart drop-zone
                $pic.Left = [single]($ChartAreaLeftPt + ($ChartAreaWidthPt - $pic.Width) / 2)
                $pic.Top = [single]($ChartAreaTopPt + ($ChartAreaHeightPt - $pic.Height) / 2)
                $pic.Name = 'ChartImage'
                # Thin flat frame, no shadow -- PowerPoint's "Simple Frame,
                # Black" picture style (the firm's preferred 2nd option in the
                # Picture Styles gallery), replicated manually since Quick
                # Styles aren't reliably settable via COM.
                $pic.Line.Visible = $msoTrue
                $pic.Line.ForeColor.RGB = ConvertTo-OleColor '404040'
                $pic.Line.Weight = 1
                $pic.Shadow.Visible = $msoFalse
            } else {
                # Blank chart drop-zone: a light bordered rectangle with a faint
                # caption. Paste the chart over it (or delete it after pasting).
                $ph = $slide.Shapes.AddShape($msoShapeRectangle, $ChartAreaLeftPt, $ChartAreaTopPt, $ChartAreaWidthPt, $ChartAreaHeightPt)
                $ph.Name = 'ChartPlaceholder'
                $ph.Fill.Solid()
                $ph.Fill.ForeColor.RGB = $GrayFill
                $ph.Line.ForeColor.RGB = $GrayLine
                $ph.Line.Weight = 1
                $pr = $ph.TextFrame.TextRange
                $pr.Text = 'Paste chart here'
                $pr.Font.Name = 'Times New Roman'
                $pr.Font.Size = 16
                $pr.Font.Italic = $msoTrue
                $pr.Font.Color.RGB = $GrayText
                $pr.ParagraphFormat.Alignment = $ppAlignCenter
                $ph.TextFrame.VerticalAnchor = $ppAnchorMiddle
            }
        }
    }

    # remove the leftover template content slide (now at index 2)
    $pres.Slides.Item(2).Delete()

    # save locally, then copy the finished deck to the output path
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
