<#
.SYNOPSIS
    Aligns the numbers in a table's numeric columns on their decimal point,
    using a decimal tab stop.

.DESCRIPTION
    A converter leaves table cells left-aligned, so a column of numbers comes
    out ragged and the decimal points wander:

        4.13          4.13
        101.97      101.97
        1750.00    1750.00
        273           273

    Right-aligning fixes it only while every value in the column carries the
    same number of decimals, and the moment someone types 40.4 next to 40.40
    it breaks again. A decimal tab stop is the real fix: Word puts the decimal
    point at the tab position whatever the value looks like, and an integer
    lands with its last digit there.

    The script finds numeric columns by looking at what is in them, computes
    where the decimal point has to sit for the longest fractional part to fit
    inside the cell, writes a decimal tab stop at that position, and puts a
    leading tab in each numeric cell to drive it.

    It only ever touches w:pPr and adds a tab run. Run properties are left
    exactly as they are, so bolding, italics, superscript footnote markers,
    and any hand adjustment survive.

    Cells that are not numeric are skipped, so a column header reading "Mean"
    or "%" keeps its own alignment. Placeholder cells - a dash, n/a - are
    tabbed with the numbers so they line up on the same edge.

    A decimal tab is an absolute position inside the cell, so the cell has to
    have a width Word will actually honour. A table left on auto layout is
    re-sized from its contents at render time and its w:gridCol values are
    only hints, so a stop computed from them lands in the wrong place. Those
    tables are skipped and reported; pin them first with
    Set-TableLook.ps1 -FixedWidth, or pass -AllowAutoLayout to take the hint
    at face value anyway.

    Needs System.Drawing to measure the font, so run it on Windows.

    Idempotent. Backs up to <name>.decimaltabs.bak before writing.

.PARAMETER File
    Single Word file to process.

.PARAMETER Folder
    Folder of Word files. All .docx/.docm files are processed.

.PARAMETER MinNumericFraction
    How much of a column has to be numeric before the column is treated as
    numeric, between 0 and 1. Default 0.6. The header row, full-width merged
    rows, and rows on -ExcludeStyles are not counted, and neither are empty or
    placeholder cells.

.PARAMETER ExcludeStyles
    Paragraph styleIds that mark a row as not part of the data grid. Defaults
    to TableandFigureFootnotes, the style Merge-TableFootnotes.ps1 puts a
    table's note rows on.

.PARAMETER RightPadInches
    Space left between the last character of the longest value and the cell's
    right text edge. Default 0.02.

.PARAMETER MinColumns
    Ignore tables with fewer columns than this. Default 2.

.PARAMETER FontName
    Override the font used for measuring. By default it is resolved from the
    -TableTextStyle chain in styles.xml, which on the NIST template gives
    Source Sans Pro 9pt.

.PARAMETER FontSize
    Override the point size used for measuring.

.PARAMETER TableTextStyle
    Style to resolve the measuring font from. Default TableText.

.PARAMETER AllowAutoLayout
    Also process tables left on auto layout, taking their w:gridCol values as
    the true column widths. Off by default, because they usually are not.

.PARAMETER Remove
    Strip the decimal tab stops and the leading tab runs this script added,
    instead of writing them.

.PARAMETER ListOnly
    Report which columns would be treated as numeric, and where the tab would
    land, and write nothing.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-TableDecimalTabs.ps1 -File "Chapter-01.docx" -ListOnly

.EXAMPLE
    .\Set-TableDecimalTabs.ps1 -File "Chapter-01.docx" -Force

.EXAMPLE
    .\Set-TableDecimalTabs.ps1 -File "Chapter-01.docx" -Remove -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [ValidateRange(0.0, 1.0)]
    [double]$MinNumericFraction = 0.6,

    [string[]]$ExcludeStyles = @('TableandFigureFootnotes'),

    [ValidateRange(0.0, 0.5)]
    [double]$RightPadInches = 0.02,

    [ValidateRange(1, 64)]
    [int]$MinColumns = 2,

    [string]$FontName = '',

    [double]$FontSize = 0,

    [string]$TableTextStyle = 'TableText',

    [switch]$AllowAutoLayout,

    [switch]$Remove,

    [switch]$ListOnly,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# CT_PPr elements that must come AFTER w:tabs.
$PPR_AFTER = @('suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
               'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
               'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
               'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
               'textDirection', 'textAlignment', 'textboxTightWrap',
               'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange')

# A value that should carry the tab but says nothing about the column's type.
$PLACEHOLDER_RX = '^(?:[-\u2010\u2011\u2012\u2013\u2014\u2015]+|[Nn]/?[Aa]|NA|n\.a\.|\u2026)$'
# A number, with optional thousands separators, sign, currency, and a trailing
# unit. The fractional part is captured so its width can be measured.
$NUMBER_RX = '^[\$\u00A3\u20AC]?[+\u2212-]?\d{1,3}(?:,\d{3})*(\.\d+)?\s*%?$|^[\$\u00A3\u20AC]?[+\u2212-]?\d+(\.\d+)?\s*%?$'

# ---------------------------------------------------------------------------
# Targets
# ---------------------------------------------------------------------------

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $resolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($resolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|docm)$' -and $_.Name -notlike '~$*' } |
        ForEach-Object { $_.FullName } | Sort-Object
}
if ($targets.Count -eq 0) { throw "No Word files found." }

if (-not $Remove) {
    try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop }
    catch { throw "This needs System.Drawing to measure the font, which means running it on Windows." }
}

Write-Host "Set Table Decimal Tabs" -ForegroundColor Cyan
if ($Remove) {
    Write-Host "  Removing decimal tab stops and their leading tab runs." -ForegroundColor Cyan
} else {
    Write-Host ("  Numeric columns: at least {0:P0} of counted cells." -f $MinNumericFraction) -ForegroundColor Cyan
    Write-Host ("  Right pad: {0}in   Rows skipped: {1}" -f $RightPadInches, ($ExcludeStyles -join ', ')) -ForegroundColor Cyan
}
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""
if ($ListOnly) { Write-Host "  -ListOnly: reporting matches, writing nothing." -ForegroundColor Yellow }

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

function Read-Package {
    param([string]$Path)
    $names   = New-Object System.Collections.ArrayList
    $entries = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $stream  = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
                [void]$names.Add($e.FullName)
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
    return ,[pscustomobject]@{ Names = $names; Data = $entries }
}

function Get-CellText {
    param([string]$Xml)
    $j = ([regex]::Matches($Xml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value }) -join ''
    return ([regex]::Replace([System.Net.WebUtility]::HtmlDecode($j), '\s+', ' ')).Trim()
}

function Get-TableTextFont {
    # Resolve the run font and size the table text actually inherits, by walking
    # the basedOn chain. Falls back to the NIST table default.
    param([string]$StylesXml, [string]$StartId = 'TableText')
    $name = $null; $size = 0
    $id = $StartId
    for ($hop = 0; $hop -lt 8 -and $id; $hop++) {
        $m = [regex]::Match($StylesXml, '(?s)<w:style\b[^>]*w:styleId="' + [regex]::Escape($id) + '"[^>]*>.*?</w:style>')
        if (-not $m.Success) { break }
        if (-not $name) {
            $f = [regex]::Match($m.Value, '<w:rFonts\b[^>]*w:ascii="([^"]+)"')
            if ($f.Success) { $name = $f.Groups[1].Value }
        }
        if ($size -le 0) {
            $z = [regex]::Match($m.Value, '<w:sz\b[^>]*w:val="(\d+)"')
            if ($z.Success) { $size = [int]$z.Groups[1].Value / 2 }
        }
        if ($name -and $size -gt 0) { break }
        $b = [regex]::Match($m.Value, '<w:basedOn\b[^>]*w:val="([^"]+)"')
        $id = if ($b.Success) { $b.Groups[1].Value } else { $null }
    }
    if (-not $name)  { $name = 'Source Sans Pro' }
    if ($size -le 0) { $size = 9 }
    return ,@($name, $size)
}

function Get-CellMarginTwips {
    # Left + right cell margin, which is width the text cannot use.
    param([string]$StylesXml, [string]$TableStyleId)
    $l = 58; $r = 29
    if ($TableStyleId) {
        $m = [regex]::Match($StylesXml, '(?s)<w:style\b[^>]*w:styleId="' + [regex]::Escape($TableStyleId) + '"[^>]*>.*?</w:style>')
        if ($m.Success) {
            $cm = [regex]::Match($m.Value, '(?s)<w:tblCellMar>.*?</w:tblCellMar>')
            if ($cm.Success) {
                $lm = [regex]::Match($cm.Value, '<w:left\b[^>]*w:w="(\d+)"')
                $rm = [regex]::Match($cm.Value, '<w:right\b[^>]*w:w="(\d+)"')
                if ($lm.Success) { $l = [int]$lm.Groups[1].Value }
                if ($rm.Success) { $r = [int]$rm.Groups[1].Value }
            }
        }
    }
    return ,@($l, $r)
}

# Writes the decimal tab stop into a paragraph's pPr, in schema order.
function Set-DecimalTabStop {
    param([string]$Para, [int]$Pos)
    $stop = '<w:tab w:val="decimal" w:pos="' + $Pos + '"/>'

    $tm = [regex]::Match($Para, '(?s)<w:tabs>.*?</w:tabs>|<w:tabs\s*/>')
    if ($tm.Success) {
        $inner = [regex]::Match($tm.Value, '(?s)^<w:tabs>(.*)</w:tabs>$')
        $body  = if ($inner.Success) { $inner.Groups[1].Value } else { '' }
        # Replace any decimal stop we own; keep the author's other stops.
        $body  = [regex]::Replace($body, '<w:tab\b[^>]*w:val="decimal"[^>]*/>', '')
        $block = '<w:tabs>' + $stop + $body + '</w:tabs>'
        return $Para.Substring(0, $tm.Index) + $block + $Para.Substring($tm.Index + $tm.Length)
    }

    $block = '<w:tabs>' + $stop + '</w:tabs>'
    $pm = [regex]::Match($Para, '(?s)<w:pPr(?:\s[^>]*)?>(.*?)</w:pPr>')
    if ($pm.Success) {
        $body = $pm.Groups[1].Value
        $at   = $body.Length
        foreach ($e in [regex]::Matches($body, '<w:([a-zA-Z0-9]+)(?:\s[^>]*)?(?:/>|>)')) {
            if ($PPR_AFTER -contains $e.Groups[1].Value) { $at = $e.Index; break }
        }
        $newPr = '<w:pPr>' + $body.Substring(0, $at) + $block + $body.Substring($at) + '</w:pPr>'
        return $Para.Substring(0, $pm.Index) + $newPr + $Para.Substring($pm.Index + $pm.Length)
    }
    if ($Para -match '<w:pPr\s*/>') {
        return [regex]::Replace($Para, '<w:pPr\s*/>', "<w:pPr>$block</w:pPr>", 1)
    }
    return [regex]::Replace($Para, '^(<w:p(?:\s[^>]*)?>)', ('${1}<w:pPr>' + $block + '</w:pPr>'), 1)
}

function Remove-DecimalTabStop {
    param([string]$Para)
    $tm = [regex]::Match($Para, '(?s)<w:tabs>.*?</w:tabs>')
    if (-not $tm.Success) { return $Para }
    $inner = [regex]::Match($tm.Value, '(?s)^<w:tabs>(.*)</w:tabs>$')
    $body  = if ($inner.Success) { $inner.Groups[1].Value } else { '' }
    $body  = [regex]::Replace($body, '<w:tab\b[^>]*w:val="decimal"[^>]*/>', '')
    $block = if ($body -match '<w:tab\b') { '<w:tabs>' + $body + '</w:tabs>' } else { '' }
    return $Para.Substring(0, $tm.Index) + $block + $Para.Substring($tm.Index + $tm.Length)
}

# True when the paragraph's first run already leads with a tab.
function Test-LeadingTabRun {
    param([string]$Para)
    $rm = [regex]::Match($Para, '(?s)<w:r(?:\s[^>]*)?>(?:(?!</w:r>).)*</w:r>')
    if (-not $rm.Success) { return $false }
    $body = [regex]::Replace($rm.Value, '(?s)<w:rPr(?:\s[^>]*)?>.*?</w:rPr>|<w:rPr\s*/>', '')
    return ($body -match '(?s)^<w:r(?:\s[^>]*)?>\s*<w:tab\s*/>')
}

function Add-LeadingTabRun {
    param([string]$Para)
    if (Test-LeadingTabRun -Para $Para) { return $Para }
    $run = '<w:r><w:tab/></w:r>'
    $pm = [regex]::Match($Para, '(?s)<w:pPr(?:\s[^>]*)?>.*?</w:pPr>|<w:pPr\s*/>')
    if ($pm.Success) {
        $at = $pm.Index + $pm.Length
        return $Para.Substring(0, $at) + $run + $Para.Substring($at)
    }
    return [regex]::Replace($Para, '^(<w:p(?:\s[^>]*)?>)', ('${1}' + $run), 1)
}

function Remove-LeadingTabRun {
    param([string]$Para)
    # Only a run that is nothing but a tab - that is the one this script adds.
    return [regex]::Replace($Para, '<w:r(?:\s[^>]*)?>(?:<w:rPr(?:\s[^>]*)?>.*?</w:rPr>)?<w:tab\s*/></w:r>', '', 1)
}

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    if (-not $pkg.Data.ContainsKey('word/document.xml')) { Write-Warning "  No word/document.xml. Skipping."; continue }

    $raw    = $utf8NoBom.GetString($pkg.Data['word/document.xml'])
    # Not StartsWith: PowerShell coerces the char to a string and .NET's
    # default comparison is culture-sensitive, which treats U+FEFF as
    # ignorable - so it answers True on a document that has no BOM, and the
    # rewrite then adds one that was never there.
    $hadBom = $raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF
    $doc    = $raw.TrimStart([char]0xFEFF)
    $styles = if ($pkg.Data.ContainsKey('word/styles.xml')) { $utf8NoBom.GetString($pkg.Data['word/styles.xml']).TrimStart([char]0xFEFF) } else { '' }

    $bodyAt = $doc.IndexOf('<w:body>')
    if ($bodyAt -lt 0) { Write-Warning "  No <w:body>. Skipping."; continue }
    $head = $doc.Substring(0, $bodyAt)
    $body = $doc.Substring($bodyAt)

    $gfx = $null; $fnt = $null; $tf = $null; $box = $null
    if (-not $Remove) {
        $fi      = Get-TableTextFont -StylesXml $styles -StartId $TableTextStyle
        $useFont = if ($FontName) { $FontName } else { $fi[0] }
        $useSize = if ($FontSize -gt 0) { $FontSize } else { $fi[1] }
        $installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
        if ($installed -notcontains $useFont) {
            Write-Warning "  '$useFont' is not installed; measuring with Segoe UI, so treat the positions as approximate."
            $useFont = 'Segoe UI'
        }
        Write-Host ("  Measuring with {0} {1}pt" -f $useFont, $useSize) -ForegroundColor Gray
        $bmp = New-Object System.Drawing.Bitmap 1, 1
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        $gfx.PageUnit = [System.Drawing.GraphicsUnit]::Inch
        $tf  = [System.Drawing.StringFormat]::GenericTypographic
        $fnt = New-Object System.Drawing.Font($useFont, $useSize, [System.Drawing.FontStyle]::Regular)
        $box = New-Object System.Drawing.SizeF(400, 400)
    }

    $edits     = New-Object System.Collections.ArrayList
    $autoSkip  = 0
    $cellCount = 0
    $colCount  = 0
    $tblCount  = 0
    $shown     = 0

    foreach ($tm in [regex]::Matches($body, '(?s)<w:tbl(?:\s[^>]*)?>.*?</w:tbl>')) {
        $table = $tm.Value

        $gridCols = [regex]::Matches($table, '<w:gridCol\b[^>]*w:w="(\d+)"') |
                    ForEach-Object { [int]$_.Groups[1].Value }
        $ncols = ([regex]::Matches($table, '<w:gridCol\b')).Count
        if ($ncols -lt $MinColumns) { continue }

        $rows = [regex]::Matches($table, '(?s)<w:tr(?:\s[^>]*)?>(?:(?!</w:tr>).)*</w:tr>')
        if ($rows.Count -lt 2) { continue }

        # A tab stop is an absolute position, so the column width has to be one
        # Word will honour. On auto layout the gridCol values are only hints.
        if (-not $Remove -and -not $AllowAutoLayout) {
            $prm = [regex]::Match($table, '(?s)<w:tblPr(?:\s[^>]*)?>.*?</w:tblPr>')
            $pr  = if ($prm.Success) { $prm.Value } else { '' }
            $fixed = ($pr -match '<w:tblLayout\s[^>]*w:type="fixed"')
            $wm    = [regex]::Match($pr, '<w:tblW\b[^>]*/>')
            $hasW  = $wm.Success -and ($wm.Value -match 'w:type="dxa"') -and
                     ($wm.Value -match 'w:w="(\d+)"') -and ([int]$Matches[1] -gt 0)
            if (-not ($fixed -and $hasW)) { $autoSkip++; continue }
        }

        $tblStyle = ''
        $sm = [regex]::Match($table, '<w:tblStyle\b[^>]*w:val="([^"]+)"')
        if ($sm.Success) { $tblStyle = $sm.Groups[1].Value }
        $mar = Get-CellMarginTwips -StylesXml $styles -TableStyleId $tblStyle

        # Walk the body rows once, recording every single-column cell with its
        # column index, its text, and where it sits in $body.
        $seen = @{}
        for ($i = 1; $i -lt $rows.Count; $i++) {
            $row = $rows[$i].Value

            $skip = $false
            if ($ExcludeStyles.Count -gt 0) {
                foreach ($psm in [regex]::Matches($row, '<w:pStyle\s[^>]*w:val="([^"]+)"')) {
                    if ($ExcludeStyles -contains $psm.Groups[1].Value) { $skip = $true; break }
                }
            }
            if ($skip) { continue }

            $col = 0
            foreach ($cm in [regex]::Matches($row, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')) {
                $span = 1
                $gm = [regex]::Match($cm.Value, '<w:gridSpan\s[^>]*w:val="(\d+)"')
                if ($gm.Success) { $span = [int]$gm.Groups[1].Value }
                if ($span -eq 1 -and $col -lt $ncols) {
                    if (-not $seen.ContainsKey($col)) { $seen[$col] = New-Object System.Collections.ArrayList }
                    [void]$seen[$col].Add([pscustomobject]@{
                        Text  = Get-CellText -Xml $cm.Value
                        Start = $tm.Index + $rows[$i].Index + $cm.Index
                        End   = $tm.Index + $rows[$i].Index + $cm.Index + $cm.Length
                        Xml   = $cm.Value
                    })
                }
                $col += $span
            }
        }

        $tableTouched = $false
        foreach ($col in ($seen.Keys | Sort-Object)) {
            $items   = $seen[$col]
            $numeric = @($items | Where-Object { $_.Text -match $NUMBER_RX })
            $holder  = @($items | Where-Object { $_.Text -match $PLACEHOLDER_RX })
            $counted = @($items | Where-Object { $_.Text -and ($_.Text -notmatch $PLACEHOLDER_RX) })
            if ($counted.Count -eq 0) { continue }
            if ($numeric.Count / [double]$counted.Count -lt $MinNumericFraction) { continue }

            $tabPos = 0
            if (-not $Remove) {
                if ($col -ge $gridCols.Count) { continue }
                $usableTw = $gridCols[$col] - $mar[0] - $mar[1]
                if ($usableTw -le 0) { continue }
                $usableIn = $usableTw / 1440.0

                # Widest fractional part, decimal point included. An integer
                # contributes nothing, so a whole-number column right-aligns.
                $maxFrac = 0.0
                foreach ($it in $numeric) {
                    $dot = $it.Text.IndexOf('.')
                    $tail = if ($dot -ge 0) { $it.Text.Substring($dot) } else { '' }
                    if (-not $tail) { continue }
                    $w = $gfx.MeasureString($tail, $fnt, $box, $tf).Width
                    if ($w -gt $maxFrac) { $maxFrac = $w }
                }

                $posIn = $usableIn - $RightPadInches - $maxFrac
                # Widest integer part has to fit to the left of the stop, or
                # Word pushes the value past it and the column looks broken.
                $maxInt = 0.0
                foreach ($it in $numeric) {
                    $dot = $it.Text.IndexOf('.')
                    $headTxt = if ($dot -ge 0) { $it.Text.Substring(0, $dot) } else { $it.Text }
                    $w = $gfx.MeasureString($headTxt, $fnt, $box, $tf).Width
                    if ($w -gt $maxInt) { $maxInt = $w }
                }
                if ($posIn -lt $maxInt) { $posIn = [Math]::Min($maxInt, $usableIn) }
                if ($posIn -le 0) { continue }
                $tabPos = [int][Math]::Round($posIn * 1440)
            }

            $targetsInCol = @($numeric + $holder)
            if ($targetsInCol.Count -eq 0) { continue }

            $changedHere = 0
            foreach ($it in $targetsInCol) {
                $new = $it.Xml
                # Every paragraph in the cell, so a wrapped value keeps its stop.
                $new = [regex]::Replace($new, '(?s)<w:p(?:\s[^>]*)?>(?:(?!</w:p>).)*</w:p>', {
                    param($pm2)
                    $p = $pm2.Value
                    if ($p -notmatch '<w:t[\s>]') { return $p }
                    if ($Remove) {
                        $p = Remove-DecimalTabStop -Para $p
                        return (Remove-LeadingTabRun -Para $p)
                    }
                    $p = Set-DecimalTabStop -Para $p -Pos $tabPos
                    return (Add-LeadingTabRun -Para $p)
                })
                if ($new -eq $it.Xml) { continue }
                $changedHere++
                [void]$edits.Add([pscustomobject]@{ Start = $it.Start; End = $it.End; Xml = $new })
            }

            if ($changedHere -gt 0) {
                $cellCount += $changedHere
                $colCount++
                $tableTouched = $true
                if ($shown -lt 8 -and -not $Remove) {
                    Write-Host ("    table {0} col {1}: {2} cells, stop at {3:N2}in" -f
                        ($tblCount + 1), ($col + 1), $changedHere, ($tabPos / 1440.0)) -ForegroundColor DarkGray
                    $shown++
                }
            }
        }
        if ($tableTouched) { $tblCount++ }
    }

    Write-Host ("  Tables: {0}  |  columns: {1}  |  cells: {2}" -f $tblCount, $colCount, $cellCount) -ForegroundColor DarkGray
    if ($autoSkip -gt 0) {
        Write-Warning ("  {0} table(s) skipped: still on auto layout, so their column widths are not fixed. Pin them with Set-TableLook.ps1 -FixedWidth, or pass -AllowAutoLayout." -f $autoSkip)
    }

    if ($edits.Count -eq 0) {
        Write-Host "  Nothing to change. File unchanged." -ForegroundColor DarkGray
        continue
    }
    if ($ListOnly) { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow; continue }

    foreach ($e in ($edits | Sort-Object Start -Descending)) {
        $body = $body.Substring(0, $e.Start) + $e.Xml + $body.Substring($e.End)
    }
    $newDoc = $head + $body

    $probe = New-Object System.Xml.XmlDocument
    $probe.PreserveWhitespace = $true
    try { $probe.LoadXml($newDoc) }
    catch { throw "Rewrite produced malformed XML for ${target}: $($_.Exception.Message)" }

    if (-not $Force -and (Test-Path -LiteralPath "$target.decimaltabs.bak")) {
        Write-Warning "  Backup $target.decimaltabs.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.decimaltabs.bak" -Force
    }

    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [char]0xFEFF + $newDoc } else { $newDoc }))

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "dectabs-$([guid]::NewGuid()).zip")
    $outStream = [System.IO.File]::Create($tempZip)
    try {
        $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $pkg.Names) {
                $ne = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try { $os.Write($pkg.Data[$name], 0, $pkg.Data[$name].Length) } finally { $os.Dispose() }
            }
        } finally { $outZip.Dispose() }
    } finally { $outStream.Dispose() }

    Move-Item -LiteralPath $tempZip -Destination $target -Force
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
