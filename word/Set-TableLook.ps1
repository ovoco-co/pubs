<#
.SYNOPSIS
    Normalizes table style options (tblLook) and enables repeating header
    rows on every table in a Word document.

.DESCRIPTION
    Word stores per-table style option checkboxes (Table Design > Table
    Style Options) in w:tblLook inside w:tbl/w:tblPr. This script
    normalizes every table to formal report standard defaults:

        Header Row     (firstRow)     = checked
        Total Row      (lastRow)      = unchecked
        First Column   (firstColumn)  = unchecked
        Last Column    (lastColumn)   = unchecked
        Banded Rows    (noHBand=0)    = checked
        Banded Columns (noVBand=1)    = unchecked

    Every one of those six is a parameter, so a document set that wants a
    different look asks for it directly - for example a converted statistical
    table, where the style's header-row and banding emphasis fights the
    formatting already in the cells:

        -HeaderRow $false -TotalRow $false -BandedRows $false -BandedColumns $true

    The legacy w:val bitmask is computed from whatever the six are set to; it
    is not hardcoded. With the defaults it comes out "0420" (firstRow +
    noVBand), which is what it has always been.

    Also writes w:tblHeader into every first row's w:trPr so the header
    row repeats across page breaks. That is independent of the Header Row
    checkbox - repeating is a page-break behavior, the checkbox is style
    emphasis - so it stays on unless -RepeatHeaderRow $false turns it off.

    Tables that use the Sidebar Table style are treated as floating callout
    boxes instead: they are stamped with the canonical right-margin float and
    text-wrapping (w:tblpPr + w:tblOverlap) and a 3" preferred width, and are
    NOT given a repeating header row. This is done here because the schema does
    not allow the float (w:tblpPr) to live in a table style.
    Operates directly on the ZIP package via System.IO.Compression - no
    Word installation required. Idempotent: re-running produces the same
    result.

.PARAMETER File
    Path to a single Word file.

.PARAMETER Folder
    Folder of Word files. All .docx/.dotx/.docm/.dotm files are processed.

.PARAMETER HeaderRow
    Header Row checkbox (w:firstRow). Default $true.

.PARAMETER TotalRow
    Total Row checkbox (w:lastRow). Default $false.

.PARAMETER FirstColumn
    First Column checkbox (w:firstColumn). Default $false.

.PARAMETER LastColumn
    Last Column checkbox (w:lastColumn). Default $false.

.PARAMETER BandedRows
    Banded Rows checkbox. Default $true. Stored inverted, as w:noHBand.

.PARAMETER BandedColumns
    Banded Columns checkbox. Default $false. Stored inverted, as w:noVBand.

.PARAMETER RepeatHeaderRow
    Write w:tblHeader into the first row so it repeats across page breaks.
    Default $true. Independent of -HeaderRow.

.PARAMETER FixedWidth
    Pin every table to a real width instead of letting Word autofit it.

    A converted table usually arrives with <w:tblW w:type="auto" w:w="0"/>,
    which tells Word to size the table from its contents. The w:gridCol widths
    are then only a hint, so tables that share an identical grid still render
    at different widths down the page. This writes <w:tblLayout w:type="fixed"/>
    and an explicit w:tblW, so the grid is what decides.

    Without -TableWidthInches the table keeps the width its own grid already
    adds up to, which is usually what the source intended.

.PARAMETER ColumnWidthsInches
    Exact per-column widths, in inches, for tables whose column count matches
    the list length. Use with -MatchHeader to hit one family of tables and
    leave the rest alone. The statistical tables in the NCSTAR and NWIRP
    chapters use 1.5 / 0.6 / 0.6.

.PARAMETER MatchHeader
    Only apply -ColumnWidthsInches to tables whose first row reads exactly
    this, cell for cell: -MatchHeader Characteristic,N,%

.PARAMETER TableWidthInches
    Target width for -FixedWidth, in inches. Column widths are rescaled in
    proportion, so the relative shape of the table is kept. The NIST text
    column is 6.5 inches. Ignored without -FixedWidth.

.PARAMETER ReportWidths
    Measure what the tables actually need and print it, writing nothing.

    Column widths are a guess unless the text is measured, and guessing has a
    wide error bar: two characters-per-inch estimates for the same content
    differed by 44 percent, which is the difference between fitting and
    wrapping. This measures every cell with the real installed font and reports
    the narrowest each column can be.

    The number that matters is usually not a column's own content. A row whose
    cell spans every column - a group-label banner - constrains the TOTAL
    width, so one long banner sets how wide the whole table has to be while
    each individual column would have fitted in half of it.

    Needs System.Drawing, so run it on Windows. It reports and stops; nothing
    is written whatever else is passed.

.PARAMETER ExcludeSpanStyles
    Paragraph styles whose full-width rows are prose and are SUPPOSED to wrap,
    so they must not drag the table's minimum width up. Defaults to the
    footnote and caption styles. Without this the merged note row at the foot
    of a table - every note in one cell - reports a table 42 inches wide.

.PARAMETER MaxTableInches
    The widest a table can be, which is the text column: 6.5 inches for NIST.
    A suggestion is never wider than this; if the content genuinely needs more,
    it is reported as rows that will wrap.

.PARAMETER MinColumnInches
    Floor for the narrow columns in the suggestion. Default 0.6. Numeric
    columns measure far narrower than that, and a 0.25 inch column looks wrong
    next to its heading.

.PARAMETER HeadroomInches
    Slack added to the measured minimum in the suggestion. Default 0.15, about
    a character and a half, so a table is not one glyph away from wrapping.

.PARAMETER FontName
    Font to measure with. Default: read from the table text style in the file.

.PARAMETER FontSize
    Point size to measure with. Default: read from the table text style.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-TableLook.ps1 -Folder "C:\templates"

.EXAMPLE
    .\Set-TableLook.ps1 -File "Report.docx" -Force

.EXAMPLE
    # Converted statistical tables: no header/total emphasis, columns banded
    # instead of rows.
    .\Set-TableLook.ps1 -File "Chapter-05.docx" -HeaderRow $false -TotalRow $false `
        -BandedRows $false -BandedColumns $true -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [bool]$HeaderRow       = $true,
    [bool]$TotalRow        = $false,
    [bool]$FirstColumn     = $false,
    [bool]$LastColumn      = $false,
    [bool]$BandedRows      = $true,
    [bool]$BandedColumns   = $false,
    [bool]$RepeatHeaderRow = $true,

    [switch]$FixedWidth,

    [double]$TableWidthInches = 0,

    [double[]]$ColumnWidthsInches = @(),

    [string[]]$MatchHeader = @(),

    [switch]$ReportWidths,

    [string[]]$ExcludeSpanStyles = @('TableandFigureFootnotes', 'CaptionTable'),

    [double]$MaxTableInches = 6.5,

    [double]$MinColumnInches = 0.6,

    [double]$HeadroomInches = 0.15,

    [string]$FontName,

    [double]$FontSize = 0,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Canonical tblLook, built from the six parameters.
#
# Banded Rows and Banded Columns are stored INVERTED in the file: the
# attributes are noHBand and noVBand, so banding on means the "no" flag is 0.
# The legacy w:val bitmask carries the same six bits for readers that predate
# the named attributes, and both have to agree.
# ---------------------------------------------------------------------------

$BIT = @{ firstRow = 0x0020; lastRow = 0x0040; firstColumn = 0x0080
          lastColumn = 0x0100; noHBand = 0x0200; noVBand = 0x0400 }

$look = [ordered]@{
    firstRow    = [int]$HeaderRow
    lastRow     = [int]$TotalRow
    firstColumn = [int]$FirstColumn
    lastColumn  = [int]$LastColumn
    noHBand     = [int](-not $BandedRows)
    noVBand     = [int](-not $BandedColumns)
}

$mask = 0
foreach ($k in $look.Keys) { if ($look[$k]) { $mask = $mask -bor $BIT[$k] } }

$CANONICAL_TBL_LOOK = '<w:tblLook w:val="{0:X4}"{1}/>' -f $mask,
    (($look.Keys | ForEach-Object { ' w:{0}="{1}"' -f $_, $look[$_] }) -join '')

Write-Host "Table Style Options to apply:" -ForegroundColor Cyan
Write-Host ("  Header Row     {0}" -f $(if ($HeaderRow)     { 'checked' } else { 'unchecked' }))
Write-Host ("  Total Row      {0}" -f $(if ($TotalRow)      { 'checked' } else { 'unchecked' }))
Write-Host ("  First Column   {0}" -f $(if ($FirstColumn)   { 'checked' } else { 'unchecked' }))
Write-Host ("  Last Column    {0}" -f $(if ($LastColumn)    { 'checked' } else { 'unchecked' }))
Write-Host ("  Banded Rows    {0}" -f $(if ($BandedRows)    { 'checked' } else { 'unchecked' }))
Write-Host ("  Banded Columns {0}" -f $(if ($BandedColumns) { 'checked' } else { 'unchecked' }))
Write-Host ("  Repeat header row across pages: {0}" -f $(if ($RepeatHeaderRow) { 'yes' } else { 'no' }))
Write-Host ("  w:tblLook -> {0}" -f $CANONICAL_TBL_LOOK) -ForegroundColor DarkGray
Write-Host ""

# ---------------------------------------------------------------------------
# Build target list
# ---------------------------------------------------------------------------

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|dotx|docm|dotm)$' } |
        ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }

Write-Host "Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "  $_" }
Write-Host ""

if (-not $Force -and -not $ReportWidths) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# Sidebar callout float: right margin, text wraps with a 0.125" gap on each side,
# tables never overlap. Applied to tables using the Sidebar Table style.
$SIDEBAR_FLOAT = '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="-5"/><w:tblOverlap w:val="never"/>'
$SIDEBAR_STYLE = '<w:tblStyle w:val="SidebarTable"/>'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Rewrite an existing <w:tblLook .../> tag in place with the canonical values.
$tblLookEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    return $CANONICAL_TBL_LOOK
}

# For each <w:tbl>...<w:tblPr>...</w:tblPr>...</w:tbl>: if tblPr lacks tblLook,
# insert canonical one before </w:tblPr>.
$tblPrEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $whole = $m.Value
    if ($whole -match '<w:tblLook\b') { return $whole }
    return ($whole -replace '</w:tblPr>', "$CANONICAL_TBL_LOOK</w:tblPr>")
}

# For each <w:tbl> using the Sidebar Table style: stamp the canonical float +
# 3" preferred width onto its tblPr. Idempotent - existing float/overlap are
# replaced. Tables not using the style are returned unchanged.
$sidebarEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $table = $m.Value
    if ($table -notmatch [regex]::Escape($SIDEBAR_STYLE)) { return $table }
    $prMatch = [regex]::Match($table, '(?s)<w:tblPr\b[^>]*>.*?</w:tblPr>')
    if (-not $prMatch.Success) { return $table }
    $pr = $prMatch.Value
    # Drop any existing float / overlap so re-runs stay idempotent.
    $np = [regex]::Replace($pr, '<w:tblpPr\b[^>]*/>', '')
    $np = [regex]::Replace($np, '<w:tblOverlap\b[^>]*/>', '')
    # Insert the canonical float + overlap immediately after the tblStyle
    # (schema order: tblStyle, tblpPr, tblOverlap, ... , tblW, ... , tblLook).
    $np = [regex]::Replace($np, "($([regex]::Escape($SIDEBAR_STYLE)))", "`$1$SIDEBAR_FLOAT", 1)
    # Force the preferred width to 3" (4325 twips): replace an instance tblW
    # (usually auto, which would override the style) or insert after tblOverlap.
    if ($np -match '<w:tblW\b[^>]*/>') {
        $np = [regex]::Replace($np, '<w:tblW\b[^>]*/>', '<w:tblW w:w="4325" w:type="dxa"/>', 1)
    } else {
        $np = [regex]::Replace($np, '(<w:tblOverlap\b[^>]*/>)', '$1<w:tblW w:w="4325" w:type="dxa"/>', 1)
    }
    return $table.Substring(0, $prMatch.Index) + $np + $table.Substring($prMatch.Index + $pr.Length)
}

# For each <w:tbl>: locate the first <w:tr>, ensure <w:tblHeader/> is in its trPr.
# Sidebar callout tables are skipped - a floating callout has no repeating header.
$firstRowEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $table = $m.Value
    if ($table -match [regex]::Escape($SIDEBAR_STYLE)) { return $table }
    # Find the first <w:tr>...</w:tr> block in this table
    $trMatch = [regex]::Match($table, '(?s)<w:tr\b[^>]*>(.*?)</w:tr>')
    if (-not $trMatch.Success) { return $table }
    $tr = $trMatch.Value
    if ($tr -match '<w:tblHeader\b') { return $table }  # already set
    # Insert <w:tblHeader/> inside <w:trPr> (create trPr if absent).
    if ($tr -match '<w:trPr\b[^>]*>') {
        $newTr = [regex]::Replace($tr, '(<w:trPr\b[^>]*>)', '$1<w:tblHeader/>', 1)
    } elseif ($tr -match '<w:trPr\s*/>') {
        $newTr = [regex]::Replace($tr, '<w:trPr\s*/>', '<w:trPr><w:tblHeader/></w:trPr>', 1)
    } else {
        # No trPr at all — insert right after the opening <w:tr ...> tag.
        $newTr = [regex]::Replace($tr, '(<w:tr\b[^>]*>)', '$1<w:trPr><w:tblHeader/></w:trPr>', 1)
    }
    # Replace just that one tr inside the table string (use literal Substring
    # math to avoid regex hazards in the surrounding content).
    $idx = $table.IndexOf($tr)
    if ($idx -lt 0) { return $table }
    return $table.Substring(0, $idx) + $newTr + $table.Substring($idx + $tr.Length)
}

# For each <w:tbl>: pin the layout so Word stops re-deriving column widths.
#
# Three levels have to agree or Word recomputes them from the contents:
# the tblGrid, the table's own tblW, and every cell's tcW. A cell spanning N
# columns takes the sum of those N, so merged rows still add up.
#
# Sidebar callouts are skipped - they carry their own 3" float width.
$fixedWidthEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $table = $m.Value
    if ($table -match [regex]::Escape($SIDEBAR_STYLE)) { return $table }

    $gridMatch = [regex]::Match($table, '(?s)<w:tblGrid>.*?</w:tblGrid>')
    if (-not $gridMatch.Success) { return $table }
    $cols = [regex]::Matches($gridMatch.Value, '<w:gridCol\b[^>]*w:w="(\d+)"')
    if ($cols.Count -eq 0) { return $table }
    $widths = @($cols | ForEach-Object { [int]$_.Groups[1].Value })

    # Explicit per-column widths, optionally gated on the header row.
    if ($script:ExplicitTwips) {
        if ($widths.Count -ne $script:ExplicitTwips.Count) { return $table }
        if ($script:HeaderMatch) {
            $row1 = [regex]::Match($table, '(?s)<w:tr(?:\s[^>]*)?>(?:(?!</w:tr>).)*</w:tr>')
            if (-not $row1.Success) { return $table }
            $cells = [regex]::Matches($row1.Value, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')
            if ($cells.Count -ne $script:HeaderMatch.Count) { return $table }
            for ($i = 0; $i -lt $cells.Count; $i++) {
                $t = (([regex]::Matches($cells[$i].Value, '<w:t(?:\s[^>]*)?>(.*?)</w:t>') |
                       ForEach-Object { $_.Groups[1].Value }) -join '')
                $t = ([regex]::Replace([System.Net.WebUtility]::HtmlDecode($t), '\s+', ' ')).Trim()
                if ($t -ne $script:HeaderMatch[$i]) { return $table }
            }
        }
        $target = $script:ExplicitTwips
    }
    elseif ($script:TargetTwips -gt 0) {
        # Rescale in proportion; rounding drift goes on the last column so the
        # columns still add up to the target exactly.
        $total = ($widths | Measure-Object -Sum).Sum
        if ($total -le 0) { return $table }
        $target = @()
        foreach ($w in $widths) { $target += [int][Math]::Round($w * $script:TargetTwips / $total) }
        $target[$target.Count - 1] += ($script:TargetTwips - (($target | Measure-Object -Sum).Sum))
    }
    else {
        $target = $widths          # keep the grid, just stop autofitting it
    }

    $script:tablesPinned++
    $totalTwips = ($target | Measure-Object -Sum).Sum

    # 1. tblGrid
    $newGrid = '<w:tblGrid>' + (($target | ForEach-Object { '<w:gridCol w:w="' + $_ + '"/>' }) -join '') + '</w:tblGrid>'
    $newTable = $table.Substring(0, $gridMatch.Index) + $newGrid +
                $table.Substring($gridMatch.Index + $gridMatch.Length)

    # 2. every cell's tcW, walking each row and honouring gridSpan
    $newTable = [regex]::Replace($newTable, '(?s)<w:tr(?:\s[^>]*)?>(?:(?!</w:tr>).)*</w:tr>', {
        param($rm)
        $row = $rm.Value
        $out = New-Object System.Text.StringBuilder
        $pos = 0
        $ci  = 0
        foreach ($cm in [regex]::Matches($row, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')) {
            $tc = $cm.Value
            $sp = [regex]::Match($tc, '<w:gridSpan\s[^>]*w:val="(\d+)"')
            $n  = if ($sp.Success) { [int]$sp.Groups[1].Value } else { 1 }
            $w  = 0
            for ($k = $ci; $k -lt [Math]::Min($ci + $n, $target.Count); $k++) { $w += $target[$k] }
            if ($w -le 0) { $w = $target[$target.Count - 1] }
            $tcW = '<w:tcW w:w="' + $w + '" w:type="dxa"/>'

            if ($tc -match '<w:tcW\b[^>]*/>') {
                $ntc = [regex]::Replace($tc, '<w:tcW\b[^>]*/>', $tcW, 1)
            } elseif ($tc -match '(?s)<w:tcPr(?:\s[^>]*)?>') {
                $ntc = [regex]::Replace($tc, '(<w:tcPr(?:\s[^>]*)?>)', ('${1}' + $tcW), 1)
            } elseif ($tc -match '<w:tcPr\s*/>') {
                $ntc = [regex]::Replace($tc, '<w:tcPr\s*/>', "<w:tcPr>$tcW</w:tcPr>", 1)
            } else {
                $ntc = [regex]::Replace($tc, '^(<w:tc(?:\s[^>]*)?>)', ('${1}<w:tcPr>' + $tcW + '</w:tcPr>'), 1)
            }
            [void]$out.Append($row.Substring($pos, $cm.Index - $pos))
            [void]$out.Append($ntc)
            $pos = $cm.Index + $cm.Length
            $ci += $n
        }
        [void]$out.Append($row.Substring($pos))
        return $out.ToString()
    })

    # 3. the table's own width and a fixed layout
    $tblW = '<w:tblW w:w="' + $totalTwips + '" w:type="dxa"/>'
    $lay  = '<w:tblLayout w:type="fixed"/>'
    $prMatch = [regex]::Match($newTable, '(?s)<w:tblPr\b[^>]*>.*?</w:tblPr>')
    if (-not $prMatch.Success) { return $newTable }
    $pr = $prMatch.Value
    $np = [regex]::Replace($pr, '<w:tblW\b[^>]*/>', '')
    $np = [regex]::Replace($np, '<w:tblLayout\b[^>]*/>', '')
    # Schema order inside tblPr: ... tblW, tblLayout, ... tblLook last.
    if ($np -match '<w:tblLook\b') {
        $np = [regex]::Replace($np, '(<w:tblLook\b)', ($tblW + $lay + '$1'), 1)
    } else {
        $np = [regex]::Replace($np, '</w:tblPr>', ($tblW + $lay + '</w:tblPr>'), 1)
    }
    return $newTable.Substring(0, $prMatch.Index) + $np + $newTable.Substring($prMatch.Index + $pr.Length)
}
# No .GetNewClosure() here on purpose: it snapshots the variables as they stand
# at definition time, and $script:ExplicitTwips / $script:HeaderMatch are set
# later, by the driver, per file. Closing over them would freeze them as null.

# ---------------------------------------------------------------------------
# Width reporting
# ---------------------------------------------------------------------------

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
    if (-not $name)   { $name = 'Source Sans Pro' }
    if ($size -le 0)  { $size = 9 }
    return ,@($name, $size)
}

function Get-CellMarginInches {
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
    return ($l + $r) / 1440.0
}

# ---------------------------------------------------------------------------
# Process each file
# ---------------------------------------------------------------------------

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $documentXml = $null
    $stream = [System.IO.File]::Open($target, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $docEntry = $zip.GetEntry('word/document.xml')
            if (-not $docEntry) { Write-Warning "  No word/document.xml. Skipping."; continue }
            $reader = New-Object System.IO.StreamReader($docEntry.Open())
            try { $documentXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }

    if (-not $documentXml) { continue }

    # ---- -ReportWidths: measure, print, write nothing ----
    if ($ReportWidths) {
        try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop }
        catch { throw "-ReportWidths needs System.Drawing, which means running this on Windows." }

        $stylesXml = ''
        $st = [System.IO.File]::Open($target, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        try {
            $z = New-Object System.IO.Compression.ZipArchive($st, [System.IO.Compression.ZipArchiveMode]::Read)
            try {
                $se = $z.GetEntry('word/styles.xml')
                if ($se) { $sr = New-Object System.IO.StreamReader($se.Open()); try { $stylesXml = $sr.ReadToEnd() } finally { $sr.Dispose() } }
            } finally { $z.Dispose() }
        } finally { $st.Dispose() }

        $fi = Get-TableTextFont -StylesXml $stylesXml
        $useFont = if ($FontName) { $FontName } else { $fi[0] }
        $useSize = if ($FontSize -gt 0) { $FontSize } else { $fi[1] }

        $installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
        if ($installed -notcontains $useFont) {
            Write-Warning "  '$useFont' is not installed; measuring with Segoe UI, so treat the numbers as approximate."
            $useFont = 'Segoe UI'
        }
        Write-Host ("  Measuring with {0} {1}pt" -f $useFont, $useSize) -ForegroundColor Gray

        $bmp = New-Object System.Drawing.Bitmap 1, 1
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        $gfx.PageUnit = [System.Drawing.GraphicsUnit]::Inch
        $tf  = [System.Drawing.StringFormat]::GenericTypographic
        $fReg = New-Object System.Drawing.Font($useFont, $useSize, [System.Drawing.FontStyle]::Regular)
        $fBld = New-Object System.Drawing.Font($useFont, $useSize, [System.Drawing.FontStyle]::Bold)
        $box = New-Object System.Drawing.SizeF(400, 400)

        # Group tables by column count + header signature, so one report line
        # covers a whole family.
        $groups = @{}
        foreach ($tm in [regex]::Matches($documentXml, '(?s)<w:tbl\b[^>]*>.*?</w:tbl>')) {
            $tbl = $tm.Value
            if ($tbl -match [regex]::Escape($SIDEBAR_STYLE)) { continue }
            $ncols = ([regex]::Matches($tbl, '<w:gridCol\b')).Count
            if ($ncols -lt 1) { continue }
            $rows = [regex]::Matches($tbl, '(?s)<w:tr(?:\s[^>]*)?>(?:(?!</w:tr>).)*</w:tr>')
            if ($rows.Count -eq 0) { continue }

            $cellText = {
                param($cellXml)
                $t = (([regex]::Matches($cellXml, '<w:t(?:\s[^>]*)?>(.*?)</w:t>') |
                       ForEach-Object { $_.Groups[1].Value }) -join '')
                return ([regex]::Replace([System.Net.WebUtility]::HtmlDecode($t), '\s+', ' ')).Trim()
            }

            $hdrCells = [regex]::Matches($rows[0].Value, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')
            $sig = (($hdrCells | ForEach-Object { & $cellText $_.Value }) -join ' | ')
            if ($MatchHeader.Count -gt 0 -and $sig -ne ($MatchHeader -join ' | ')) { continue }

            $tsm = [regex]::Match($tbl, '<w:tblStyle\b[^>]*w:val="([^"]+)"')
            $margin = Get-CellMarginInches -StylesXml $stylesXml -TableStyleId $(if ($tsm.Success) { $tsm.Groups[1].Value } else { $null })

            $key = "$ncols|$sig"
            if (-not $groups.ContainsKey($key)) {
                $groups[$key] = [pscustomobject]@{
                    Cols = $ncols; Sig = $sig; Count = 0; Margin = $margin
                    ColMin = @(0.0) * $ncols          # widest non-spanning cell per column
                    SpanMin = 0.0                      # widest cell spanning every column
                    SpanText = ''
                }
            }
            $grp = $groups[$key]
            $grp.Count++

            foreach ($r in $rows) {
                $ci = 0
                foreach ($cm in [regex]::Matches($r.Value, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')) {
                    $cx = $cm.Value
                    $sp = [regex]::Match($cx, '<w:gridSpan\s[^>]*w:val="(\d+)"')
                    $n  = if ($sp.Success) { [int]$sp.Groups[1].Value } else { 1 }
                    $txt = & $cellText $cx
                    if ($txt) {
                        $f = if ($cx -match '<w:b\s*/>' -or $cx -match '<w:b\s+w:val="(?:1|true|on)"') { $fBld } else { $fReg }
                        $w = $gfx.MeasureString($txt, $f, $box, $tf).Width
                        if ($n -ge $ncols) {
                            # Spans the table: constrains the TOTAL, not a column -
                            # unless it is prose that is meant to wrap.
                            $prose = $false
                            foreach ($ps in [regex]::Matches($cx, '<w:pStyle\s[^>]*w:val="([^"]+)"')) {
                                if ($ExcludeSpanStyles -contains $ps.Groups[1].Value) { $prose = $true; break }
                            }
                            if (-not $prose -and $w -gt $grp.SpanMin) { $grp.SpanMin = $w; $grp.SpanText = $txt }
                        } elseif ($n -eq 1 -and $ci -lt $ncols) {
                            if ($w -gt $grp.ColMin[$ci]) { $grp.ColMin[$ci] = $w }
                        }
                    }
                    $ci += $n
                }
            }
        }

        if ($groups.Count -eq 0) {
            Write-Host "  No tables matched." -ForegroundColor DarkGray
            continue
        }

        $ceil05 = { param($v) [Math]::Ceiling($v * 20) / 20 }

        foreach ($key in ($groups.Keys | Sort-Object)) {
            $g = $groups[$key]
            Write-Host ""
            Write-Host ("  {0}-column tables [{1}]  x{2}" -f $g.Cols, $g.Sig, $g.Count) -ForegroundColor Yellow
            for ($i = 0; $i -lt $g.Cols; $i++) {
                Write-Host ("    column {0}: widest text {1,6:N3}`"  -> needs {2,6:N3}`"" -f
                    ($i + 1), $g.ColMin[$i], ($g.ColMin[$i] + $g.Margin)) -ForegroundColor DarkGray
            }
            $needTotal = 0.0
            if ($g.SpanMin -gt 0) {
                $needTotal = $g.SpanMin + $g.Margin
                Write-Host ("    full-width rows: widest {0,6:N3}`" -> the TABLE must reach {1:N3}`"" -f $g.SpanMin, $needTotal) -ForegroundColor DarkGray
                Write-Host ("      longest: {0}" -f $g.SpanText) -ForegroundColor DarkGray
            }

            # Suggestion. Two shapes of table need different treatment:
            #
            #  - one where the columns comfortably fit, and the width is really
            #    set by the full-width banner rows. Trailing columns go to a
            #    common width, and column 1 absorbs whatever the banners need.
            #  - one where the columns alone already overflow the text column.
            #    Then nothing "absorbs" anything: the space is shared out in
            #    proportion to what each column actually needs, and the report
            #    says plainly that content will wrap.
            $minCol = @()
            for ($i = 0; $i -lt $g.Cols; $i++) { $minCol += ($g.ColMin[$i] + $g.Margin) }

            $small = 0.0
            if ($g.Cols -ge 2) {
                $widestTrailing = (($minCol[1..($g.Cols - 1)]) | Measure-Object -Maximum).Maximum
                $small = & $ceil05 ([Math]::Max($MinColumnInches, $widestTrailing))
            }
            $first  = & $ceil05 $minCol[0]
            $target = $needTotal + $HeadroomInches
            if ($g.Cols -ge 2 -and ($first + $small * ($g.Cols - 1)) -lt $target) {
                $first = & $ceil05 ($target - $small * ($g.Cols - 1))
            }
            $widths = @($first)
            for ($i = 1; $i -lt $g.Cols; $i++) { $widths += $small }
            $total = ($widths | Measure-Object -Sum).Sum

            $notes = @()
            if ($total -gt $MaxTableInches) {
                # Not enough room. Do NOT share it out in proportion: that
                # shrinks a money or count column below the width of its own
                # digits, so every figure wraps, while the prose column - the
                # one that can wrap harmlessly - is the one given the slack.
                #
                # Give every column exactly what it measured, and make the
                # WIDEST column absorb the shortfall. A long name wrapping to
                # two lines is normal; "$78,199,438,009" wrapping is not.
                $sumMin  = ($minCol | Measure-Object -Sum).Sum
                $floor   = 0.35
                $widest  = 0
                for ($i = 1; $i -lt $g.Cols; $i++) { if ($minCol[$i] -gt $minCol[$widest]) { $widest = $i } }

                $widths = @()
                foreach ($m in $minCol) { $widths += [Math]::Round($m, 2) }
                $over = (($widths | Measure-Object -Sum).Sum) - $MaxTableInches
                $widths[$widest] = [Math]::Round($widths[$widest] - $over, 2)

                # If taking it all off the widest column would leave it below
                # the floor, fall back to proportional - nothing else fits.
                if ($widths[$widest] -lt $floor) {
                    $widths = @()
                    foreach ($m in $minCol) {
                        $widths += [Math]::Max($floor, [Math]::Round($MaxTableInches * $m / $sumMin, 2))
                    }
                    $drift = $MaxTableInches - (($widths | Measure-Object -Sum).Sum)
                    $widths[$widest] = [Math]::Round($widths[$widest] + $drift, 2)
                    if ($widths[$widest] -lt $floor) { $widths[$widest] = $floor }
                }
                $total = ($widths | Measure-Object -Sum).Sum

                if ($sumMin -gt $MaxTableInches) {
                    $notes += ("does not fit: the columns alone need {0:N2}`", the text column is {1}`"" -f $sumMin, $MaxTableInches)
                } else {
                    $notes += ("shared out proportionally: equal trailing columns would have overflowed the {0}`" text column" -f $MaxTableInches)
                }
                # 0.01 tolerance, or rounding alone reports a column as short.
                for ($i = 0; $i -lt $g.Cols; $i++) {
                    if (($minCol[$i] - $widths[$i]) -gt 0.01) {
                        $notes += ("  column {0} gets {1:N2}`" but wants {2:N2}`" - its text will wrap" -f ($i + 1), $widths[$i], $minCol[$i])
                    }
                }
            }
            if ($needTotal -gt 0 -and ($needTotal - $total) -gt 0.01) {
                $notes += ("  the full-width rows want {0:N2}`" and will wrap" -f $needTotal)
            }
            foreach ($nte in $notes) { Write-Host ("    $nte") -ForegroundColor Yellow }

            Write-Host ("    suggested: -ColumnWidthsInches {0} -MatchHeader {1}   (total {2:N2}`")" -f
                (($widths | ForEach-Object { $_.ToString() }) -join ','),
                (($g.Sig -split ' \| ') -join ','), $total) -ForegroundColor Green
        }

        Write-Host ""
        Write-Host "  -ReportWidths: nothing written." -ForegroundColor Yellow
        continue
    }

    # 1. Normalize every existing <w:tblLook ...> tag.
    $tblLookCountBefore = ([regex]::Matches($documentXml, '<w:tblLook\b')).Count
    $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
        $documentXml,
        '<w:tblLook\b[^/]*/>',
        $tblLookEvaluator)
    Write-Host "  Normalized $tblLookCountBefore existing tblLook tag(s)" -ForegroundColor DarkGray

    # 2. Insert a canonical tblLook into every <w:tbl>/<w:tblPr> that lacks one.
    #    Match <w:tbl>...<w:tblPr>...</w:tblPr> non-greedily.
    $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
        $documentXml,
        '(?s)<w:tbl\b[^>]*>.*?<w:tblPr\b[^>]*>.*?</w:tblPr>',
        $tblPrEvaluator)

    # 3. Ensure first <w:tr> of every <w:tbl> carries <w:tblHeader/> in trPr.
    $tableCount = ([regex]::Matches($documentXml, '(?s)<w:tbl\b[^>]*>.*?</w:tbl>')).Count
    if ($RepeatHeaderRow) {
        $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
            $documentXml,
            '(?s)<w:tbl\b[^>]*>.*?</w:tbl>',
            $firstRowEvaluator)
        Write-Host "  Visited $tableCount table(s); ensured Repeat-Header-Row" -ForegroundColor DarkGray
    } else {
        Write-Host "  Visited $tableCount table(s); Repeat-Header-Row left alone" -ForegroundColor DarkGray
    }

    # 4. Sidebar callout tables: stamp the float + 3" width.
    $sidebarCount = ([regex]::Matches($documentXml, "(?s)<w:tbl\b[^>]*>(?:(?!</w:tbl>).)*?$([regex]::Escape($SIDEBAR_STYLE))")).Count
    $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
        $documentXml,
        '(?s)<w:tbl\b[^>]*>.*?</w:tbl>',
        $sidebarEvaluator)
    if ($sidebarCount -gt 0) {
        Write-Host "  Sidebar Table callouts floated (right margin, 3`"): $sidebarCount" -ForegroundColor DarkGray
    }

    # 4b. Pin table widths so Word stops autofitting them to their contents.
    if ($FixedWidth -or $ColumnWidthsInches.Count -gt 0) {
        $script:TargetTwips   = [int][Math]::Round($TableWidthInches * 1440)
        $script:ExplicitTwips = if ($ColumnWidthsInches.Count -gt 0) {
            @($ColumnWidthsInches | ForEach-Object { [int][Math]::Round($_ * 1440) })
        } else { $null }
        $script:HeaderMatch   = if ($MatchHeader.Count -gt 0) { $MatchHeader } else { $null }
        $script:tablesPinned  = 0

        $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
            $documentXml,
            '(?s)<w:tbl\b[^>]*>.*?</w:tbl>',
            $fixedWidthEvaluator)

        if ($script:ExplicitTwips) {
            $ins = ($ColumnWidthsInches | ForEach-Object { '{0}"' -f $_ }) -join ' / '
            $hdr = if ($script:HeaderMatch) { " matching [$($MatchHeader -join ' | ')]" } else { '' }
            Write-Host ("  Columns set to $ins on $script:tablesPinned table(s)$hdr") -ForegroundColor DarkGray
        } elseif ($script:TargetTwips -gt 0) {
            Write-Host ("  {0} table(s) pinned to {1}`" (fixed layout)" -f $script:tablesPinned, $TableWidthInches) -ForegroundColor DarkGray
        } else {
            Write-Host ("  {0} table(s) pinned to their own grid (fixed layout)" -f $script:tablesPinned) -ForegroundColor DarkGray
        }
    }

    # 5. Write back
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "tbllook-$([guid]::NewGuid()).zip")
    $srcStream = [System.IO.File]::OpenRead($target)
    try {
        $srcZip = New-Object System.IO.Compression.ZipArchive($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $outStream = [System.IO.File]::Create($tempZip)
            try {
                $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
                try {
                    foreach ($e in $srcZip.Entries) {
                        $newEntry = $outZip.CreateEntry($e.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
                        if ($e.FullName -eq 'word/document.xml') {
                            $w = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
                            try { $w.Write($documentXml) } finally { $w.Dispose() }
                        } else {
                            $es = $e.Open()
                            try {
                                $ns = $newEntry.Open()
                                try { $es.CopyTo($ns) } finally { $ns.Dispose() }
                            } finally { $ns.Dispose() }
                        }
                    }
                } finally { $outZip.Dispose() }
            } finally { $outStream.Dispose() }
        } finally { $srcZip.Dispose() }
    } finally { $srcStream.Dispose() }

    Move-Item -LiteralPath $tempZip -Destination $target -Force
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
if ($ReportWidths) {
    Write-Host "Done. Nothing written - pass the suggested -ColumnWidthsInches back in to apply it." -ForegroundColor Green
} else {
    Write-Host "Done. Open in Word, click any table, and verify under Table Design > Table Style Options." -ForegroundColor Green
}
