<#
.SYNOPSIS
    Runs a colored bar down one edge of every row in a table, so a reader can
    tell at a glance which population or category the table belongs to.

.DESCRIPTION
    A chapter that reports the same measures for two populations - schools and
    hospitals, say - is a wall of tables that all look alike. The NCSTAR
    chapters solve that with a thick colored border on the first cell of every
    row, which reads as a continuous spine down the left edge of the table in
    that population's assigned color:

        |  Characteristic          N        %
        |  Ownership
        |    Public              196    65.99
        |    Private             101    34.01

    Unlike the banner-row tint, the spine does not depend on the table having
    banner rows, so it carries the code through the small summary tables too.
    It is drawn at full strength (D55E00, 009E73) rather than as a tint,
    because a 4pt rule needs the saturation to register.

    Which color a table gets comes from -ColorRules, matched against the
    caption above it. A table with no caption of its own inherits the nearest
    caption above it, which is what a continuation table needs.

    The footnote row at the foot of a table is skipped. It is a merged
    full-width row that sits outside the data grid, so a spine through it
    reads as the table continuing past its own last row. -ExcludeStyles
    controls that, and a row it names has any existing bar stripped rather
    than merely left alone, so the rule applies to a document the bar has
    already been drawn on.

    Idempotent: an existing bar on the same edge is replaced, and borders on
    the other edges are left alone. Backs up to <name>.accentbar.bak.

    Uses System.IO.Compression only. Word is not required.

.PARAMETER File
    Single Word file to process.

.PARAMETER Folder
    Folder of Word files. All .docx/.docm files are processed.

.PARAMETER Color
    Bar color as six hex digits, no hash. The fallback for any table no
    -ColorRules entry claims. Default 00257A, Resolution Blue.

.PARAMETER ColorRules
    Caption-driven colors. Each entry is 'Pattern=HEX', where Pattern is a
    regex matched case-insensitively against the caption above the table. The
    first entry that matches wins, so order them most specific first.

        -ColorRules 'Hospital=009E73','School=D55E00'

    Anything unmatched falls back to -Color.

.PARAMETER CaptionStyles
    Paragraph styleIds -ColorRules reads a table's caption from. Defaults to
    CaptionTable, Caption, and Heading1 through Heading3.

.PARAMETER ExcludeStyles
    Paragraph styleIds that mark a row as not part of the data grid, so it
    gets no bar. Defaults to TableandFigureFootnotes, the style
    Merge-TableFootnotes.ps1 puts a table's note rows on. Pass an empty array
    to bar every row.

.PARAMETER Edge
    Which edge of the table the bar runs down. Default left.

.PARAMETER WidthPoints
    Bar thickness in points. Default 4, which is what the NCSTAR chapters use.
    Word stores it as eighths of a point, so this is multiplied by 8.

.PARAMETER Column
    Which column carries the bar, 1-based. Default 1. A full-width merged
    banner row has only one cell, so column 1 always lands on it.

.PARAMETER SkipHeaderRow
    Leave the header row's cell alone, so the bar starts below the header.
    Off by default: the NCSTAR spine runs the full height of the table.

.PARAMETER Remove
    Strip the bar from the chosen edge instead of drawing it. Other borders
    are left alone.

.PARAMETER ListOnly
    Report what would change and write nothing.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-TableAccentBar.ps1 -File "Chapter-05.docx" -ColorRules 'Hospital=009E73','School=D55E00' -ListOnly

.EXAMPLE
    .\Set-TableAccentBar.ps1 -File "Chapter-05.docx" -ColorRules 'Hospital=009E73','School=D55E00' -Force

.EXAMPLE
    .\Set-TableAccentBar.ps1 -File "Chapter-05.docx" -Remove -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$Color = '00257A',

    [string[]]$ColorRules = @(),

    [string[]]$CaptionStyles = @('CaptionTable', 'Caption', 'Heading1', 'Heading2', 'Heading3'),

    [string[]]$ExcludeStyles = @('TableandFigureFootnotes'),

    [ValidateSet('left', 'right', 'top', 'bottom')]
    [string]$Edge = 'left',

    [ValidateRange(0.25, 12)]
    [double]$WidthPoints = 4,

    [ValidateRange(1, 64)]
    [int]$Column = 1,

    [switch]$SkipHeaderRow,

    [switch]$Remove,

    [switch]$ListOnly,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if ($Color -notmatch '^[0-9A-Fa-f]{6}$') {
    throw "-Color must be six hex digits with no hash. Got: $Color"
}

# powershell.exe -File does not split a comma-separated argument into an
# array, so 'A=1','B=2' arrives as the single string 'A=1,B=2'. Split it back.
$ruleEntries = @()
foreach ($raw in $ColorRules) {
    if (-not $raw) { continue }
    $ruleEntries += [regex]::Split($raw, ',(?=[^,]*=)')
}

# Ordered, so the first matching pattern wins.
$rules = New-Object System.Collections.ArrayList
foreach ($entry in $ruleEntries) {
    if (-not $entry) { continue }
    $eq = $entry.LastIndexOf('=')
    if ($eq -lt 1 -or $eq -eq $entry.Length - 1) {
        throw "-ColorRules entries must look like 'Pattern=HEX'. Got: $entry"
    }
    $pattern = $entry.Substring(0, $eq).Trim()
    $hex     = $entry.Substring($eq + 1).Trim()
    if ($hex -notmatch '^[0-9A-Fa-f]{6}$') {
        throw "-ColorRules color must be six hex digits with no hash. Got: $hex (in '$entry')"
    }
    try { [void][regex]::new($pattern) }
    catch { throw "-ColorRules pattern is not a valid regex: $pattern (in '$entry')" }
    [void]$rules.Add([pscustomobject]@{ Pattern = $pattern; Color = $hex })
}

$captionStyleAlt = ($CaptionStyles | ForEach-Object { [regex]::Escape($_) }) -join '|'

# Word stores border width in eighths of a point.
$sz = [int][Math]::Round($WidthPoints * 8)

# CT_TcBorders child order. The bar element has to be inserted in the right
# place or Word rejects the part.
$BORDER_ORDER = @('top', 'start', 'left', 'bottom', 'end', 'right',
                  'insideH', 'insideV', 'tl2br', 'tr2bl')

# CT_TcPr elements that must come AFTER tcBorders.
$TCPR_AFTER = @('shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText',
                'vAlign', 'hideMark')

# The transitional and strict names for the same edge.
$edgeAliases = switch ($Edge) {
    'left'  { @('left', 'start') }
    'right' { @('right', 'end') }
    default { @($Edge) }
}
$edgeAlt = ($edgeAliases | ForEach-Object { [regex]::Escape($_) }) -join '|'
$edgeRx  = '<w:(?:' + $edgeAlt + ')(?:\s[^>]*)?(?:/>|>.*?</w:(?:' + $edgeAlt + ')>)'

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

Write-Host "Set Table Accent Bar" -ForegroundColor Cyan
if ($Remove) {
    Write-Host ("  Removing the {0} border from column {1} of every row." -f $Edge, $Column) -ForegroundColor Cyan
} else {
    Write-Host ("  {0} border, {1}pt, on column {2} of every row." -f $Edge, $WidthPoints, $Column) -ForegroundColor Cyan
    if ($rules.Count -gt 0) {
        Write-Host ("  Caption rules (first match wins), read from: {0}" -f ($CaptionStyles -join ', ')) -ForegroundColor Cyan
        foreach ($r in $rules) { Write-Host ("    /{0}/ -> {1}" -f $r.Pattern, $r.Color) -ForegroundColor Cyan }
        Write-Host ("    unmatched -> {0}" -f $Color) -ForegroundColor Cyan
    } else {
        Write-Host ("  Color: {0}" -f $Color) -ForegroundColor Cyan
    }
}
if ($SkipHeaderRow) { Write-Host "  Header row left alone" -ForegroundColor Cyan }
if (-not $Remove -and $ExcludeStyles.Count -gt 0) {
    Write-Host ("  Rows cleared, not barred: {0}" -f ($ExcludeStyles -join ', ')) -ForegroundColor Cyan
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

function Get-RowText {
    param([string]$Xml)
    $j = ([regex]::Matches($Xml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value }) -join ''
    return ([regex]::Replace([System.Net.WebUtility]::HtmlDecode($j), '\s+', ' ')).Trim()
}

# Puts the bar element into a tcBorders block in schema order.
function Add-EdgeToBorders {
    param([string]$Borders, [string]$Element)
    $inner = [regex]::Match($Borders, '(?s)^<w:tcBorders(?:\s[^>]*)?>(.*)</w:tcBorders>$')
    $body  = if ($inner.Success) { $inner.Groups[1].Value } else { '' }
    $mine  = [array]::IndexOf($BORDER_ORDER, $Edge)
    $at    = $body.Length
    foreach ($m in [regex]::Matches($body, '<w:([a-zA-Z0-9]+)(?:\s[^>]*)?(?:/>|>)')) {
        $name = $m.Groups[1].Value
        $idx  = [array]::IndexOf($BORDER_ORDER, $name)
        if ($idx -ge 0 -and $idx -gt $mine) { $at = $m.Index; break }
    }
    return '<w:tcBorders>' + $body.Substring(0, $at) + $Element + $body.Substring($at) + '</w:tcBorders>'
}

# Puts a tcBorders block into a tcPr in schema order.
function Add-BordersToCellProps {
    param([string]$Cell, [string]$Borders)
    if ($Cell -match '(?s)<w:tcPr(?:\s[^>]*)?>.*?</w:tcPr>') {
        $m     = [regex]::Match($Cell, '(?s)<w:tcPr(?:\s[^>]*)?>(.*?)</w:tcPr>')
        $body  = $m.Groups[1].Value
        $at    = $body.Length
        foreach ($e in [regex]::Matches($body, '<w:([a-zA-Z0-9]+)(?:\s[^>]*)?(?:/>|>)')) {
            if ($TCPR_AFTER -contains $e.Groups[1].Value) { $at = $e.Index; break }
        }
        $newPr = '<w:tcPr>' + $body.Substring(0, $at) + $Borders + $body.Substring($at) + '</w:tcPr>'
        return $Cell.Substring(0, $m.Index) + $newPr + $Cell.Substring($m.Index + $m.Length)
    }
    if ($Cell -match '<w:tcPr\s*/>') {
        return [regex]::Replace($Cell, '<w:tcPr\s*/>', "<w:tcPr>$Borders</w:tcPr>", 1)
    }
    return [regex]::Replace($Cell, '^(<w:tc(?:\s[^>]*)?>)', ('${1}<w:tcPr>' + $Borders + '</w:tcPr>'), 1)
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

    $bodyAt = $doc.IndexOf('<w:body>')
    if ($bodyAt -lt 0) { Write-Warning "  No <w:body>. Skipping."; continue }
    $head = $doc.Substring(0, $bodyAt)
    $body = $doc.Substring($bodyAt)

    # Caption paragraphs, in document order, so each table can find the one
    # above it. A table with no caption of its own inherits the nearest
    # caption above it, which is what a continuation table needs.
    $captions = New-Object System.Collections.ArrayList
    if ($rules.Count -gt 0 -and -not $Remove) {
        foreach ($pm in [regex]::Matches($body, '(?s)<w:p(?:\s[^>]*)?>(?:(?!</w:p>).)*</w:p>')) {
            if ($pm.Value -notmatch ('<w:pStyle\s[^>]*w:val="(?:' + $captionStyleAlt + ')"')) { continue }
            $ct = Get-RowText -Xml $pm.Value
            if ($ct) { [void]$captions.Add([pscustomobject]@{ Index = $pm.Index; Text = $ct }) }
        }
    }

    $edits   = New-Object System.Collections.ArrayList
    $cells   = 0
    $cleared = 0
    $tables  = 0
    $groups  = @{}

    foreach ($tm in [regex]::Matches($body, '(?s)<w:tbl(?:\s[^>]*)?>.*?</w:tbl>')) {
        $table = $tm.Value
        if (([regex]::Matches($table, '<w:gridCol\b')).Count -lt 1) { continue }

        $rows = [regex]::Matches($table, '(?s)<w:tr(?:\s[^>]*)?>(?:(?!</w:tr>).)*</w:tr>')
        if ($rows.Count -lt 1) { continue }

        # Which color this table's bar gets. Fall back to -Color.
        $tableColor = $Color
        $tableRule  = ''
        if ($rules.Count -gt 0 -and -not $Remove) {
            $cap = ''
            for ($c = $captions.Count - 1; $c -ge 0; $c--) {
                if ($captions[$c].Index -lt $tm.Index) { $cap = $captions[$c].Text; break }
            }
            foreach ($r in $rules) {
                if ($cap -match $r.Pattern) { $tableColor = $r.Color; $tableRule = $r.Pattern; break }
            }
        }
        $element = '<w:' + $Edge + ' w:val="single" w:sz="' + $sz +
                   '" w:space="0" w:color="' + $tableColor.ToUpperInvariant() + '"/>'

        $touched = $false
        for ($i = 0; $i -lt $rows.Count; $i++) {
            if ($SkipHeaderRow -and $i -eq 0) { continue }
            $row = $rows[$i].Value

            # A row on an excluded style is not part of the data grid. Strip
            # any bar it already has rather than just skipping it, so the rule
            # holds on a document the bar has already been drawn on.
            $excluded = $false
            if ($ExcludeStyles.Count -gt 0) {
                foreach ($sm in [regex]::Matches($row, '<w:pStyle\s[^>]*w:val="([^"]+)"')) {
                    if ($ExcludeStyles -contains $sm.Groups[1].Value) { $excluded = $true; break }
                }
            }

            $tcs = [regex]::Matches($row, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')
            if ($tcs.Count -lt $Column) { continue }
            $tc  = $tcs[$Column - 1]
            $old = $tc.Value

            # Drop any existing bar on this edge, wherever it sits.
            $new = $old
            $bm  = [regex]::Match($new, '(?s)<w:tcBorders(?:\s[^>]*)?>.*?</w:tcBorders>')
            if ($bm.Success) {
                $stripped = [regex]::Replace($bm.Value, $edgeRx, '')
                # An empty tcBorders is legal but pointless; drop it.
                if ($stripped -match '(?s)^<w:tcBorders(?:\s[^>]*)?>\s*</w:tcBorders>$') { $stripped = '' }
                $new = $new.Substring(0, $bm.Index) + $stripped + $new.Substring($bm.Index + $bm.Length)
            }

            if (-not $Remove -and -not $excluded) {
                $bm2 = [regex]::Match($new, '(?s)<w:tcBorders(?:\s[^>]*)?>.*?</w:tcBorders>')
                if ($bm2.Success) {
                    $merged = Add-EdgeToBorders -Borders $bm2.Value -Element $element
                    $new = $new.Substring(0, $bm2.Index) + $merged + $new.Substring($bm2.Index + $bm2.Length)
                } else {
                    $new = Add-BordersToCellProps -Cell $new -Borders ('<w:tcBorders>' + $element + '</w:tcBorders>')
                }
            }

            if ($new -eq $old) { continue }
            if ($excluded -and -not $Remove) { $cleared++ } else { $cells++; $touched = $true }
            [void]$edits.Add([pscustomobject]@{
                Start = $tm.Index + $rows[$i].Index + $tc.Index
                End   = $tm.Index + $rows[$i].Index + $tc.Index + $tc.Length
                Xml   = $new
            })
        }

        if ($touched) {
            $tables++
            if (-not $Remove) {
                $key = if ($tableRule) { "$tableRule -> $tableColor" } else { "(unmatched) -> $tableColor" }
                if (-not $groups.ContainsKey($key)) { $groups[$key] = 0 }
                $groups[$key]++
            }
        }
    }

    Write-Host ("  Tables changed: {0}  |  cells barred: {1}  |  excluded rows cleared: {2}" -f $tables, $cells, $cleared) -ForegroundColor DarkGray
    foreach ($k in ($groups.Keys | Sort-Object)) {
        Write-Host ("    {0}  {1} table(s)" -f $k, $groups[$k]) -ForegroundColor DarkGray
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

    if (-not $Force -and (Test-Path -LiteralPath "$target.accentbar.bak")) {
        Write-Warning "  Backup $target.accentbar.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.accentbar.bak" -Force
    }

    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [char]0xFEFF + $newDoc } else { $newDoc }))

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "accentbar-$([guid]::NewGuid()).zip")
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
