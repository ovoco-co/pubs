<#
.SYNOPSIS
    Builds one deck per GFX chart: styles each Excel chart to charts/chart-styling-guide.md
    (Arial, 8 pt everywhere, guide palette and lines), sizes it to its placed size so the 8 pt
    survives, pastes it into a 6.5 in wide slide whose height follows the original figure's
    proportions, adds the source number, saves <gfx>a-NWIRP.pptx and a 300 dpi PNG.

    Layouts: single, side (two across), stack (two down), grid (grid = [columns, rows]).

    Mirrors StyleCharts_Excel.bas / StyleCharts_PowerPoint.bas, driven over COM so no VBA
    project access is needed, and adds the scatter and area arms those modules lack.
#>
param(
    [Parameter(Mandatory = $true)][string]$Spec,
    [string[]]$Only
)
$ErrorActionPreference = 'Stop'

$W = 468.0           # 6.5 in
$LABEL_STRIP = 9.0   # room under the art for the source number
$FONT = 'Arial'; $PT = 8

function RGB([string]$hex) {
    [int]('0x' + $hex.Substring(0, 2)) + 256 * [int]('0x' + $hex.Substring(2, 2)) + 65536 * [int]('0x' + $hex.Substring(4, 2))
}
$INK = RGB '000000'; $AXIS = RGB '000000'; $GRID = RGB 'EEEEF2'; $BASE = RGB 'D8D8DE'; $WHITE = RGB 'FFFFFF'   # all text black

function Set-Font($fnt, [int]$color, [bool]$bold = $false) {
    $fnt.Name = $FONT; $fnt.Size = $PT; $fnt.Color = $color; $fnt.Bold = $bold
}

# ---- PowerPoint legend ---------------------------------------------------------------------------
# House legend style (nw-2.3-Timeline-v2-ARTB.pptx, slides 4-6, without its rule): a letter-spaced
# bold LEGEND heading, then items in left-aligned columns - an 8 pt rounded
# square (or a short stroke for a line series) and an Arial 8 pt #333333 label 12.7 pt
# from the swatch's left edge, rows 12 pt apart. As many columns as fit across the slide; two-panel
# figures share one legend (items de-duplicated by name, first-seen order).
$script:legendItems = New-Object System.Collections.ArrayList
$script:legendLeft = 0.0          # x where the legend starts: the first chart's plot-area left edge

function Add-LegendItem([string]$name, [string]$hex, [string]$kind, $dash, $weight, $marker) {
    foreach ($it in $script:legendItems) { if ($it.name -eq $name) { return } }
    [void]$script:legendItems.Add([pscustomobject]@{ name = $name; hex = $hex; kind = $kind; dash = $dash
                                                     weight = $weight; marker = $marker })
}

$LEG_FONT = 'Arial'; $LEG_HEAD_FONT = 'Arial'; $LEG_INK = RGB '000000'; $LEG_RULE = RGB '333333'   # Arial, like the chart text
$SWATCH = 8.0; $SWATCH_RADIUS = 0.32383          # adjustment value used in the timeline legends
$LINE_SWATCH = 22.0; $LINE_SWATCH_MAXW = 1.75       # long and thin enough to show two dash cycles
$COL_GAP = 13.0; $ROW_PITCH = 12.0
$RULE_GAP = 4.0; $HEAD_H = 10.0; $HEAD_GAP = 3.0

function Measure-Label($slide, [string]$text, [string]$font = $LEG_FONT) {
    $tb = $slide.Shapes.AddTextbox(1, 0, 0, 400, 10)
    $tf = $tb.TextFrame
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.WordWrap = 0
    $tf.TextRange.Text = $text
    $tf.TextRange.Font.Name = $font; $tf.TextRange.Font.Size = $PT
    $bw = $tf.TextRange.BoundWidth; $tb.Delete(); return $bw
}

function Add-Label($slide, [string]$text, [single]$left, [single]$cy, [string]$font, [bool]$bold, [single]$spacing) {
    $tb = $slide.Shapes.AddTextbox(1, $left, $cy - 5, 300, 10)
    $tf = $tb.TextFrame
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.WordWrap = 0; $tf.AutoSize = 1; $tf.VerticalAnchor = 1
    $tf.TextRange.Text = $text
    $tf.TextRange.Font.Name = $font; $tf.TextRange.Font.Size = $PT; $tf.TextRange.Font.Bold = [int]$bold
    $tf.TextRange.Font.Color.RGB = $LEG_INK
    if ($spacing) { $tb.TextFrame2.TextRange.Font.Spacing = $spacing }
    # place the rendered text, not the box: left edge at $left, vertical centre on $cy
    $tr = $tf.TextRange
    $tb.Left = $tb.Left + ($left - $tr.BoundLeft)
    $tb.Top = $tb.Top + ($cy - ($tr.BoundTop + $tr.BoundHeight / 2))
    return $tb
}

function Add-Marker($slide, $style, [single]$cx, [single]$cy, [int]$rgb) {
    # Excel marker style -> PowerPoint autoshape: circle, square, diamond, triangle, star, plus
    $shape = 9
    switch ($style) { 8 { $shape = 9 } 1 { $shape = 1 } 2 { $shape = 4 } 3 { $shape = 7 } 5 { $shape = 92 } 9 { $shape = 11 } }
    $sz = 5.0
    $m = $slide.Shapes.AddShape($shape, $cx - $sz / 2, $cy - $sz / 2, $sz, $sz)
    $m.Fill.ForeColor.RGB = $rgb; $m.Line.Visible = 0
    return $m
}

function Get-LegendPlan($slide) {
    # grid of items: the largest column count whose summed column widths fit the slide width
    if ($script:legendItems.Count -eq 0) { return @{ cols = 0; height = 0 } }
    $items = @()
    foreach ($it in $script:legendItems) {
        $sw = $SWATCH
        if ($it.kind -eq 'line') { $sw = $LINE_SWATCH }
        $items += [pscustomobject]@{ it = $it; sw = $sw; tw = (Measure-Label $slide $it.name) }
    }
    $n = $items.Count
    for ($cols = $n; $cols -ge 1; $cols--) {
        $widths = @(0.0) * $cols
        for ($k = 0; $k -lt $n; $k++) {
            $c = $k % $cols
            $iw = [math]::Max($items[$k].sw, $SWATCH) + 4.7 + $items[$k].tw    # label starts 12.7 pt after an 8 pt swatch
            if ($iw -gt $widths[$c]) { $widths[$c] = $iw }
        }
        $total = ($widths | Measure-Object -Sum).Sum + $COL_GAP * ($cols - 1)
        if ($total -le ($W - $script:legendLeft)) { break }
    }
    $rows = [math]::Ceiling($n / $cols)
    $h = $RULE_GAP + $HEAD_H + $HEAD_GAP + $rows * $ROW_PITCH
    return @{ items = $items; cols = $cols; widths = $widths; rows = $rows; height = $h }
}

function Add-Legend($slide, $plan, [single]$top) {
    if ($plan.cols -eq 0) { return }
    $names = New-Object System.Collections.ArrayList
    # no rule above the legend: it ran edge to edge across the figure
    $hy = $top + $RULE_GAP + $HEAD_H / 2
    [void]$names.Add((Add-Label $slide 'LEGEND' $script:legendLeft $hy $LEG_HEAD_FONT $true 1).Name)

    $x = @([double]$script:legendLeft)
    for ($c = 1; $c -lt $plan.cols; $c++) { $x += $x[$c - 1] + $plan.widths[$c - 1] + $COL_GAP }
    for ($k = 0; $k -lt $plan.items.Count; $k++) {
        $item = $plan.items[$k]; $it = $item.it; $rgb = RGB $it.hex
        $c = $k % $plan.cols; $r = [math]::Floor($k / $plan.cols)
        $x0 = $x[$c]
        $cy = $top + $RULE_GAP + $HEAD_H + $HEAD_GAP + $r * $ROW_PITCH + $ROW_PITCH / 2
        if ($it.kind -eq 'line') {
            $ln = $slide.Shapes.AddLine($x0, $cy, $x0 + $LINE_SWATCH, $cy)
            $ln.Line.ForeColor.RGB = $rgb; $ln.Line.Weight = [single][math]::Min([double]$it.weight, $LINE_SWATCH_MAXW); $ln.Line.DashStyle = $it.dash
            [void]$names.Add($ln.Name)
            if ($it.marker -ne -4142) { [void]$names.Add((Add-Marker $slide $it.marker ($x0 + $LINE_SWATCH / 2) $cy $rgb).Name) }
            $labelX = $x0 + $LINE_SWATCH + 4.7
        } elseif ($it.kind -eq 'marker') {
            [void]$names.Add((Add-Marker $slide 8 ($x0 + $SWATCH / 2) $cy $rgb).Name)
            $labelX = $x0 + 12.7
        } else {
            $b = $slide.Shapes.AddShape(5, $x0, $cy - $SWATCH / 2, $SWATCH, $SWATCH)
            $b.Adjustments.Item(1) = $SWATCH_RADIUS
            $b.Fill.ForeColor.RGB = $rgb; $b.Line.Visible = 0
            [void]$names.Add($b.Name)
            $labelX = $x0 + 12.7
        }
        [void]$names.Add((Add-Label $slide $it.name $labelX $cy $LEG_FONT $false 0).Name)
    }
    $g = $slide.Shapes.Range([string[]]$names.ToArray()).Group(); $g.Name = 'Legend'
}

function Style-Chart($c, $e, [int]$index) {
    $c.ChartArea.Format.Fill.Visible = -1; $c.ChartArea.Format.Fill.ForeColor.RGB = $WHITE
    $c.ChartArea.Format.Line.Visible = 0
    $c.ChartArea.Font.Name = $FONT; $c.ChartArea.Font.Size = $PT; $c.ChartArea.Font.Color = $AXIS
    $c.PlotArea.Format.Fill.Visible = 0; $c.PlotArea.Format.Line.Visible = 0

    if ($e.panel_titles) {
        $c.HasTitle = $true
        $c.ChartTitle.Text = $e.panel_titles[$index]
        Set-Font $c.ChartTitle.Font $INK $true
        $c.ChartTitle.IncludeInLayout = $true
    } else {
        $c.HasTitle = $false     # the Word caption carries the title
    }

    $horizontal = $c.ChartType -in @(57, 58, 59)     # xlBarClustered / Stacked / Stacked100
    foreach ($t in 1, 2) {
        try { $ax = $c.Axes($t) } catch { continue }
        $ax.MajorTickMark = -4142; $ax.MinorTickMark = -4142
        Set-Font $ax.TickLabels.Font $AXIS
        if ($ax.HasTitle) {
            if ($t -eq 2 -and $e.value_title) { $ax.AxisTitle.Text = $e.value_title }   # e.g. with a line break to fit a narrow panel
            Set-Font $ax.AxisTitle.Font $AXIS; $ax.AxisTitle.IncludeInLayout = $true
        }
        if ($t -eq 2) {
            # the step came from the original full-size chart; at the placed size, labels closer than
            # 11 pt collide (16 pt side by side on a horizontal bar chart), so widen the step until they clear
            try {
                $span = $ax.MaximumScale - $ax.MinimumScale
                $inner = $c.PlotArea.InsideHeight; $minGap = 11
                if ($horizontal) { $inner = $c.PlotArea.InsideWidth; $minGap = 16 }
                $guard = 0
                while ($ax.MajorUnit -gt 0 -and ($inner / ($span / $ax.MajorUnit)) -lt $minGap -and $guard -lt 6) {
                    $ax.MajorUnit = $ax.MajorUnit * 2; $guard++
                }
            } catch {}
            $ax.Format.Line.Visible = 0
            $ax.HasMajorGridlines = $true
            $ax.MajorGridlines.Format.Line.ForeColor.RGB = $GRID
            $ax.MajorGridlines.Format.Line.Weight = 0.75
        } else {
            $ax.Format.Line.Visible = -1
            $ax.Format.Line.ForeColor.RGB = $BASE
            $ax.Format.Line.Weight = 0.5
            $ax.HasMajorGridlines = $false
            $ax.TickLabelPosition = -4134          # low: labels stay under the plot even when 0 is mid-axis
            if ($e.x_format) { $ax.TickLabels.NumberFormat = $e.x_format }
        }
    }

    $n = $c.SeriesCollection().Count
    $isBar = $e.kind -in @('bar', 'hist')
    $c.HasLegend = $false          # legends are drawn in PowerPoint (Add-Legend), not by Excel
    $wantLegend = ($n -gt 1)

    for ($k = 1; $k -le $n; $k++) {
        $s = $c.SeriesCollection($k)
        $name = $s.Name
        $hex = $null
        if ($e.colors -and $e.colors.PSObject.Properties[$name]) { $hex = $e.colors.$name }
        elseif ($e.chart_colors) { $hex = $e.chart_colors[$index] }
        if (-not $hex) { Write-Warning "  no colour for series '$name' in $($e.gfx)"; continue }
        $col = RGB $hex

        if ($e.kind -eq 'line') {
            if ($e.marker_only -and ($e.marker_only -contains $name)) {
                $s.Format.Line.Visible = 0
                $s.MarkerStyle = 8; $s.MarkerSize = 6
                $s.MarkerForegroundColor = $col; $s.MarkerBackgroundColor = $col
                if ($wantLegend) { Add-LegendItem $name $hex 'marker' 1 0 8 }
                continue
            }
            $s.Format.Line.Visible = -1
            $s.Format.Line.ForeColor.RGB = $col
            $s.Format.Line.Weight = $(if ($name -match 'bound|threshold') { 1.25 } else { 2.25 })
            if ($e.markers) {
                $s.MarkerSize = 5
                $s.MarkerForegroundColor = $col; $s.MarkerBackgroundColor = $col
            } else {
                $s.MarkerStyle = -4142
            }
            if ($wantLegend) {
                $mk = -4142
                if ($e.markers) { $mk = $s.MarkerStyle }
                Add-LegendItem $name $hex 'line' $s.Format.Line.DashStyle $s.Format.Line.Weight $mk
            }
        } elseif ($e.kind -eq 'area') {
            $s.Format.Fill.Visible = -1; $s.Format.Fill.Solid(); $s.Format.Fill.ForeColor.RGB = $col
            $s.Format.Line.Visible = -1; $s.Format.Line.ForeColor.RGB = $WHITE; $s.Format.Line.Weight = 0.5
            if ($wantLegend) { Add-LegendItem $name $hex 'box' 1 0 -4142 }
        } else {
            $s.Format.Fill.Visible = -1; $s.Format.Fill.Solid(); $s.Format.Fill.ForeColor.RGB = $col
            if ($wantLegend) { Add-LegendItem $name $hex 'box' 1 0 -4142 }
            if ($e.kind -eq 'hist') {
                $s.Format.Line.Visible = -1; $s.Format.Line.ForeColor.RGB = $WHITE; $s.Format.Line.Weight = 0.75
            } else {
                $s.Format.Line.Visible = 0
                $sh = $s.Format.Shadow
                $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Transparency = 0.85
                $sh.Blur = 4; $sh.OffsetX = 0; $sh.OffsetY = 2
            }
            if ($e.point_colors) {
                for ($p = 1; $p -le $s.Points().Count; $p++) {
                    $pt = $s.Points($p)
                    $pt.Format.Fill.Solid(); $pt.Format.Fill.ForeColor.RGB = (RGB $e.point_colors[$p - 1])
                }
            }
            if ($s.HasErrorBars) {
                $eb = $s.ErrorBars
                $eb.EndStyle = 2                           # xlNoCap
                $eb.Format.Line.ForeColor.RGB = $INK; $eb.Format.Line.Weight = 0.75
            }
            if ($e.labels) {
                $s.HasDataLabels = $true
                $dl = $s.DataLabels()
                $dl.Position = 2
                Set-Font $dl.Font $INK
            }
        }
    }

    if ($isBar) {
        $g = $c.ChartGroups(1)
        if ($e.kind -eq 'hist') { $g.GapWidth = 0; $g.Overlap = 0 }
        else { $g.GapWidth = 60; $g.Overlap = $(if ($n -gt 1) { 10 } else { 0 }) }
    }
}

$entries = Get-Content -Raw -Encoding UTF8 $Spec | ConvertFrom-Json
if ($Only) { $entries = $entries | Where-Object { $Only -contains $_.gfx } }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$ppt = New-Object -ComObject PowerPoint.Application
try {
    foreach ($e in $entries) {
        $name = "$($e.gfx)a-NWIRP"
        $wbPath = Join-Path $e.dir $e.wb
        $wb = $xl.Workbooks.Open($wbPath, 0, $true)
        try {
            $ws = $wb.Worksheets.Item('Data')
            $cos = @(); foreach ($co in $ws.ChartObjects()) { $cos += $co }
            $cos = @($cos | Sort-Object -Property @{Expression = { $_.Top }}, @{Expression = { $_.Left }})

            $W = [double]$e.width_in * 72                 # script-scope: the legend layout reads it too
            $H = [math]::Round([double]$e.height_in * 72, 2)
            switch ($e.layout) {
                'single' { $cw = $W;     $chh = $H;     $pos = @(,@(0, 0)) }
                'side'   { $cw = $W / 2; $chh = $H;     $pos = @(@(0, 0), @(($W / 2), 0)) }
                'stack'  { $cw = $W;     $chh = $H / 2; $pos = @(@(0, 0), @(0, ($H / 2))) }
                'grid'   {                                # grid = [columns, rows], filled row by row
                    $gc = [int]$e.grid[0]; $gr = [int]$e.grid[1]
                    $cw = $W / $gc; $chh = $H / $gr; $pos = @()
                    for ($r = 0; $r -lt $gr; $r++) { for ($q = 0; $q -lt $gc; $q++) { $pos += ,@(($q * $cw), ($r * $chh)) } }
                }
            }
            if ($cos.Count -ne $pos.Count) { throw "$name expects $($pos.Count) chart(s), workbook has $($cos.Count)" }

            $script:legendItems.Clear()
            $pres = $ppt.Presentations.Add(0)
            $pres.PageSetup.SlideWidth = $W
            $pres.PageSetup.SlideHeight = $H + $LABEL_STRIP
            $slide = $pres.Slides.Add(1, 12)
            $slide.Name = "Fig $($e.fig)"

            for ($i = 0; $i -lt $cos.Count; $i++) {
                $cos[$i].Width = [double]$cw; $cos[$i].Height = [double]$chh
                Style-Chart $cos[$i].Chart $e $i
            }
            # grid panels share one plot rectangle, so axes line up across rows and columns
            if ($e.layout -eq 'grid') {
                $L = 0.0; $T = 0.0; $R = 1e9; $B = 1e9
                foreach ($co in $cos) {
                    $pa = $co.Chart.PlotArea
                    $L = [math]::Max($L, $pa.InsideLeft); $T = [math]::Max($T, $pa.InsideTop)
                    $R = [math]::Min($R, $pa.InsideLeft + $pa.InsideWidth); $B = [math]::Min($B, $pa.InsideTop + $pa.InsideHeight)
                }
                if ($e.plot_left) { $L = [double]$e.plot_left }     # room for long category labels Excel would truncate
                foreach ($co in $cos) {
                    $pa = $co.Chart.PlotArea
                    $pa.InsideLeft = $L; $pa.InsideTop = $T; $pa.InsideWidth = $R - $L; $pa.InsideHeight = $B - $T
                }
            }
            # a single series coloured by point (5-9) names its colours explicitly
            foreach ($li in @($e.legend_items)) { if ($li) { Add-LegendItem $li[0] $li[1] 'box' 1 0 -4142 } }
            # align the legend with the plot area, so nothing sits on the image edge
            $script:legendLeft = [double]$pos[0][0] + [double]$cos[0].Chart.PlotArea.InsideLeft
            $plan = @(Get-LegendPlan $slide)[-1]       # last output only: PowerShell can wrap the return
            $legendH = [double]$plan.height
            $pres.PageSetup.SlideHeight = $H + $legendH + $LABEL_STRIP

            for ($i = 0; $i -lt $cos.Count; $i++) {
                $co = $cos[$i]
                $shape = $null
                for ($a = 1; $a -le 5 -and -not $shape; $a++) {
                    $co.Copy() | Out-Null
                    Start-Sleep -Milliseconds (250 * $a)
                    try { $shape = $slide.Shapes.Paste().Item(1) } catch { $shape = $null }
                    # the clipboard is shared with the desktop; a paste can arrive as a picture
                    if ($shape -and $shape.HasChart -eq 0) { $shape.Delete(); $shape = $null }
                }
                if (-not $shape) { throw "Could not paste chart $i of $name" }
                $shape.LockAspectRatio = 0
                $shape.Left = [single]$pos[$i][0]; $shape.Top = [single]$pos[$i][1]
                $shape.Width = [single]$cw; $shape.Height = [single]$chh
                if ($shape.HasChart -eq 0) { Write-Warning "  $name chart $i pasted as a picture" }
            }

            Add-Legend $slide $plan ($H + 2)
            $tb = $slide.Shapes.AddTextbox(1, $W - 60, $H + $legendH + 1, 60, 8)
            $tf = $tb.TextFrame
            $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
            $tf.WordWrap = -1; $tf.AutoSize = 0
            $tf.TextRange.Text = $name
            $tf.TextRange.ParagraphFormat.Alignment = 3
            $tf.TextRange.Font.Name = $FONT; $tf.TextRange.Font.Size = 6; $tf.TextRange.Font.Color.RGB = (RGB '000000')
            $tb.Name = 'Source number'

            $pptDir = Join-Path $e.dir 'ppt'; $pngDir = Join-Path $e.dir 'png'
            New-Item -ItemType Directory -Force $pptDir, $pngDir | Out-Null
            $out = Join-Path $pptDir "$name.pptx"
            if (Test-Path $out) { Remove-Item $out -Force }
            $pres.SaveAs($out)
            $slideH = $pres.PageSetup.SlideHeight
            $slide.Export((Join-Path $pngDir "$name.png"), 'PNG', [int]($W / 72 * 300), [int]($slideH / 72 * 300))
            $pres.Close()
            Write-Host ("{0}  Fig {1}  {2} chart(s)  {3:N0} x {4:N0} pt" -f $name, $e.fig, $cos.Count, $W, $slideH)
        } catch {
            Write-Host "FAILED $name : $($_.Exception.Message)"
            Write-Host $_.ScriptStackTrace
            throw
        } finally {
            $wb.Close($false)
        }
    }
} finally {
    $ppt.Quit(); $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
