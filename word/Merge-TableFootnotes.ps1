<#
.SYNOPSIS
    Collapses the footnote rows at the bottom of each table into a single
    full-width, unshaded row on the table-footnote style.

.DESCRIPTION
    Statistical tables carry their notes as extra rows at the foot of the
    table - the a/b/c qualifiers, the significance codes, the "Reference
    levels:" legend. A converter emits each note as its own row, and because
    those rows are full-width merges they collect whatever the table style
    paints on them: the banding stripe, or a header/banner fill if something
    upstream mistook a merged row for a banner.

    For every table this script walks UP from the last row, gathers the run of
    footnote rows it finds, and replaces the whole run with one row:

      - one cell spanning every column, so the notes read across the table,
      - w:fill="auto" plus a zeroed w:cnfStyle on the row, so neither the
        style's banding nor its first/last-row emphasis can paint it, wherever
        the row happens to land in the stripe pattern,
      - one paragraph per note, stacked, on the -Style paragraph style,
      - bold stripped and the indent dropped, since notes sit flush left.

    Superscript and subscript markers are left alone - they are run properties
    and nothing here touches w:rPr beyond the bold.

    Identifying the rows. By default a trailing row counts as a footnote row
    when it is a full-width merge (one cell spanning all the columns) and has
    text in it. The scan stops at the first row that is not one, so ordinary
    body rows are never touched. Pass -SourceHtml to match against the
    <tfoot> blocks of the HTML the tables came from instead, which is exact
    rather than structural, or -Marker to require the text to look like a note.

    Run -ListOnly first. A table whose real last row is a full-width banner or
    a merged "Total" line matches the same structural rule as a footnote.

    Idempotent: the row this writes is recognized and rebuilt identically on a
    re-run. Backs up to <name>.footnotes.bak before writing.

    Uses System.IO.Compression and System.Xml only. Word is not required.

.PARAMETER File
    Single Word file to process.

.PARAMETER Folder
    Folder of Word files. All .docx/.docm files are processed.

.PARAMETER SourceHtml
    The HTML the tables were converted from. Its <tfoot> rows are read and a
    trailing row is treated as a footnote row only when its text matches one
    of them. Exact, where the default rule is structural. Use it when a table
    legitimately ends in a full-width row that is not a note.

.PARAMETER Style
    Paragraph style for the merged notes. Defaults to TableandFigureFootnotes,
    which is the NIST template's style for exactly this.

.PARAMETER Marker
    Regex the row text must match for the row to count as a footnote. Applied
    on top of the full-width test. Unset by default (no text filter). For the
    lettered-note convention: -Marker '^[a-z][A-Z(]' .

.PARAMETER KeepBold
    Leave bold runs in the notes alone. By default bold is stripped, because a
    converter that mistook the merged row for a banner will have bolded it.

.PARAMETER KeepIndent
    Leave w:ind alone. By default it is dropped so the notes sit flush left.

.PARAMETER ListOnly
    Report what would be merged and write nothing.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Merge-TableFootnotes.ps1 -File "Chapter-05.docx" -ListOnly

.EXAMPLE
    .\Merge-TableFootnotes.ps1 -File "Chapter-05.docx" -Force

.EXAMPLE
    # Match the notes against the source HTML rather than the row shape.
    .\Merge-TableFootnotes.ps1 -File "Chapter-05.docx" `
        -SourceHtml "descriptivetabs_RB_August17.html" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$SourceHtml,

    [string]$Style = 'TableandFigureFootnotes',

    [string]$Marker,

    [switch]$KeepBold,

    [switch]$KeepIndent,

    [switch]$ListOnly,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

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
        Where-Object { $_.Extension -match '^\.(docx|docm)$' -and $_.Name -notlike '~$*' } |
        ForEach-Object { $_.FullName } | Sort-Object
}
if ($targets.Count -eq 0) { throw "No Word files found." }

# ---------------------------------------------------------------------------
# Optional: the <tfoot> text of the source HTML
# ---------------------------------------------------------------------------

$footTexts = New-Object 'System.Collections.Generic.HashSet[string]'
if ($SourceHtml) {
    if (-not (Test-Path -LiteralPath $SourceHtml -PathType Leaf)) { throw "Source HTML not found: $SourceHtml" }
    $html = Get-Content -LiteralPath $SourceHtml -Raw -Encoding UTF8
    foreach ($tf in [regex]::Matches($html, '(?is)<tfoot\b.*?</tfoot>')) {
        foreach ($tr in [regex]::Matches($tf.Value, '(?is)<tr\b.*?</tr>')) {
            $bare = [regex]::Replace($tr.Value, '(?s)<[^>]+>', '')
            $bare = [System.Net.WebUtility]::HtmlDecode($bare)
            $bare = ([regex]::Replace($bare, '\s+', ' ')).Trim()
            if ($bare) { [void]$footTexts.Add($bare) }
        }
    }
    if ($footTexts.Count -eq 0) { throw "No <tfoot> rows found in $SourceHtml." }
}

# ---------------------------------------------------------------------------
# Announce
# ---------------------------------------------------------------------------

Write-Host "Merge Table Footnotes" -ForegroundColor Cyan
Write-Host "  Note paragraph style: $Style" -ForegroundColor Cyan
if ($SourceHtml) {
    Write-Host "  Footnote rows identified from: $SourceHtml ($($footTexts.Count) distinct <tfoot> row(s))" -ForegroundColor Cyan
} else {
    Write-Host "  Footnote rows identified structurally: trailing full-width merged rows" -ForegroundColor Cyan
}
if ($Marker)     { Write-Host "  Row text must also match: $Marker" -ForegroundColor Cyan }
if ($KeepBold)   { Write-Host "  Bold left in place" -ForegroundColor Cyan }
if ($KeepIndent) { Write-Host "  Indentation left in place" -ForegroundColor Cyan }
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""
if ($ListOnly) { Write-Host "  -ListOnly: reporting matches, writing nothing." -ForegroundColor Yellow }

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Constants
#
# A zeroed cnfStyle tells Word this row is none of the conditional-format
# slots, so Table Style Options - banding, first row, last row - paint nothing
# on it. Without it the row picks up whatever stripe it lands on.
# ---------------------------------------------------------------------------

$CNF_NONE = '<w:cnfStyle w:val="000000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" ' +
            'w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" ' +
            'w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" ' +
            'w:lastRowLastColumn="0"/>'

$NO_FILL = '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'

# <w:tbl> and <w:tr> boundaries, written so the character after the name has to
# be whitespace or '>'. That keeps <w:tblPr>, <w:tblGrid>, <w:tblLook> and
# <w:trPr> out of the scan.
$TOKEN_RX = '</?w:(?:tbl|tr)(?:\s[^>]*)?>'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-TableSpans {
    # Every <w:tbl> in the text, nested ones included, as absolute spans.
    param([string]$Text)
    $spans = New-Object System.Collections.ArrayList
    $stack = New-Object System.Collections.Stack
    foreach ($m in [regex]::Matches($Text, $TOKEN_RX)) {
        if ($m.Value -like '<w:tbl*')       { $stack.Push($m.Index) }
        elseif ($m.Value -eq '</w:tbl>') {
            if ($stack.Count -gt 0) {
                $start = $stack.Pop()
                [void]$spans.Add([pscustomobject]@{ Start = $start; End = $m.Index + $m.Length })
            }
        }
    }
    return ,$spans
}

function Get-DirectRowSpans {
    # Rows that are direct children of this table, as spans relative to it.
    # Rows belonging to a nested table sit at a deeper level and are skipped.
    param([string]$Table)
    $spans = New-Object System.Collections.ArrayList
    $depth = 0
    $rowStart = -1
    foreach ($m in [regex]::Matches($Table, $TOKEN_RX)) {
        $v = $m.Value
        if     ($v -like '<w:tbl*')  { $depth++ }
        elseif ($v -eq '</w:tbl>')   { $depth-- }
        elseif ($v -like '<w:tr*')   { if ($depth -eq 1) { $rowStart = $m.Index } }
        elseif ($v -eq '</w:tr>')    {
            if ($depth -eq 1 -and $rowStart -ge 0) {
                [void]$spans.Add([pscustomobject]@{ Start = $rowStart; End = $m.Index + $m.Length })
                $rowStart = -1
            }
        }
    }
    return ,$spans
}

function Get-RowText {
    param([string]$Xml)
    $joined = ([regex]::Matches($Xml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') |
        ForEach-Object { $_.Groups[1].Value }) -join ''
    $joined = [System.Net.WebUtility]::HtmlDecode($joined)
    return ([regex]::Replace($joined, '\s+', ' ')).Trim()
}

function Test-FullWidthRow {
    # One cell, and it spans every column in the grid. Counting every w:tc in
    # the row, nested ones included, is deliberate: a trailing row that holds a
    # nested table comes back with more than one cell and is left alone, which
    # is the safe answer for a row this script would otherwise flatten.
    param([string]$Row, [int]$ColumnCount)
    $cells = [regex]::Matches($Row, '(?s)<w:tc(?:\s[^>]*)?>').Count
    if ($cells -ne 1) { return $false }
    $span = [regex]::Match($Row, '<w:gridSpan\s[^>]*w:val="(\d+)"')
    if ($span.Success) { return ([int]$span.Groups[1].Value -ge $ColumnCount) }
    return ($ColumnCount -le 1)
}

function ConvertTo-NotePara {
    # Put one note paragraph on the footnote style, drop the bold a banner
    # misdetection would have added, and un-indent it. w:rPr is otherwise
    # untouched, so superscript markers survive.
    param([string]$Para)

    $p = $Para
    if (-not $KeepBold) {
        $p = [regex]::Replace($p, '<w:b(?:\s+w:val="(?:1|true|on)")?\s*/>', '')
        $p = [regex]::Replace($p, '<w:bCs(?:\s+w:val="(?:1|true|on)")?\s*/>', '')
    }
    if (-not $KeepIndent) {
        $p = [regex]::Replace($p, '<w:ind\b[^>]*/>', '')
    }

    $newStyle = '<w:pStyle w:val="' + $Style + '"/>'
    $openTag  = [regex]::Match($p, '^<w:p(?:\s[^>]*)?>').Value
    $pPrMatch = [regex]::Match($p, '(?s)<w:pPr(?:\s[^>]*)?(?:/>|>.*?</w:pPr>)')

    if (-not $pPrMatch.Success) {
        # pPr has to be the first child of w:p.
        return $openTag + '<w:pPr>' + $newStyle + '</w:pPr>' + $p.Substring($openTag.Length)
    }

    $pPr = $pPrMatch.Value
    $styleMatch = [regex]::Match($pPr, '<w:pStyle\s[^>]*w:val="[^"]*"[^>]*/>')
    if ($styleMatch.Success) {
        $newPPr = $pPr.Remove($styleMatch.Index, $styleMatch.Length).Insert($styleMatch.Index, $newStyle)
    } elseif ($pPr -match '/>$') {
        $newPPr = '<w:pPr>' + $newStyle + '</w:pPr>'
    } else {
        # pStyle has to be pPr's first child.
        $pPrOpen = [regex]::Match($pPr, '^<w:pPr(?:\s[^>]*)?>').Value
        $newPPr  = $pPrOpen + $newStyle + $pPr.Substring($pPrOpen.Length)
    }
    return $p.Remove($pPrMatch.Index, $pPrMatch.Length).Insert($pPrMatch.Index, $newPPr)
}

function Read-Entries {
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

# ---------------------------------------------------------------------------
# Process each file
# ---------------------------------------------------------------------------

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $pkg = Read-Entries -Path $target
    if (-not $pkg.Data.ContainsKey('word/document.xml')) {
        Write-Warning "  No word/document.xml. Skipping."
        continue
    }

    # Word writes document.xml with a UTF-8 BOM; GetString keeps it as a
    # leading U+FEFF. Strip it, and put it back only if it was there.
    $raw = $utf8NoBom.GetString($pkg.Data['word/document.xml'])
    $hadBom = $raw.StartsWith([char]0xFEFF)
    $doc = $raw.TrimStart([char]0xFEFF)

    $bodyAt = $doc.IndexOf('<w:body>')
    if ($bodyAt -lt 0) { Write-Warning "  No <w:body>. Skipping."; continue }
    $head = $doc.Substring(0, $bodyAt)
    $body = $doc.Substring($bodyAt)

    $tables = Get-TableSpans -Text $body
    $edits  = New-Object System.Collections.ArrayList

    $tablesWithNotes = 0
    $multiRowTables  = 0
    $rowsCollapsed   = 0
    $tableIndex      = 0

    foreach ($t in ($tables | Sort-Object Start)) {
        $tableIndex++
        $table = $body.Substring($t.Start, $t.End - $t.Start)

        $grid = [regex]::Matches($table, '<w:gridCol\b[^>]*w:w="(\d+)"') |
            ForEach-Object { [int]$_.Groups[1].Value }
        $ncols = @($grid).Count
        if ($ncols -eq 0) { $ncols = 1 }

        $rows = Get-DirectRowSpans -Table $table
        if ($rows.Count -lt 2) { continue }

        # Walk up from the last row and take the run of footnote rows.
        $notes = New-Object System.Collections.ArrayList
        for ($i = $rows.Count - 1; $i -ge 0; $i--) {
            $rowXml = $table.Substring($rows[$i].Start, $rows[$i].End - $rows[$i].Start)
            $text   = Get-RowText -Xml $rowXml
            if (-not $text) { break }

            if ($SourceHtml) {
                if (-not $footTexts.Contains($text)) { break }
            } else {
                if (-not (Test-FullWidthRow -Row $rowXml -ColumnCount $ncols)) { break }
                if ($Marker -and $text -notmatch $Marker) { break }
            }
            [void]$notes.Insert(0, [pscustomobject]@{ Span = $rows[$i]; Xml = $rowXml; Text = $text })
        }

        if ($notes.Count -eq 0) { continue }
        if ($notes.Count -ge $rows.Count) {
            # Every row matched. That is a one-column table, not a table with
            # notes at the foot of it.
            Write-Host "    Table $tableIndex : every row matched; skipped" -ForegroundColor DarkYellow
            continue
        }

        $tablesWithNotes++
        if ($notes.Count -gt 1) { $multiRowTables++; $rowsCollapsed += $notes.Count }

        Write-Host ("    Table {0} : {1} row(s) -> 1  {2}" -f `
            $tableIndex, $notes.Count, $notes[0].Text.Substring(0, [Math]::Min(60, $notes[0].Text.Length))) -ForegroundColor DarkGray

        # Build the single replacement row.
        $paras = New-Object System.Collections.ArrayList
        foreach ($n in $notes) {
            foreach ($pm in [regex]::Matches($n.Xml, '(?s)<w:p(?:\s[^>]*)?>.*?</w:p>')) {
                [void]$paras.Add((ConvertTo-NotePara -Para $pm.Value))
            }
        }
        if ($paras.Count -eq 0) { continue }

        $totalWidth = 0
        foreach ($g in $grid) { $totalWidth += $g }
        $tcW = if ($totalWidth -gt 0) { '<w:tcW w:w="' + $totalWidth + '" w:type="dxa"/>' }
               else                   { '<w:tcW w:w="0" w:type="auto"/>' }
        $gridSpan = if ($ncols -gt 1) { '<w:gridSpan w:val="' + $ncols + '"/>' } else { '' }

        $newRow = '<w:tr><w:trPr>' + $CNF_NONE + '</w:trPr><w:tc><w:tcPr>' +
                  $tcW + $gridSpan + $NO_FILL + '</w:tcPr>' +
                  ($paras -join '') + '</w:tc></w:tr>'

        [void]$edits.Add([pscustomobject]@{
            Start = $t.Start + $notes[0].Span.Start
            End   = $t.Start + $notes[$notes.Count - 1].Span.End
            Xml   = $newRow
        })
    }

    Write-Host ("  Tables scanned: {0}  |  with footnote rows: {1}  |  of those, multi-row sets collapsed: {2} ({3} rows -> {2})" -f `
        $tables.Count, $tablesWithNotes, $multiRowTables, $rowsCollapsed) -ForegroundColor DarkGray

    if ($edits.Count -eq 0) {
        Write-Host "  No footnote rows found. File unchanged." -ForegroundColor DarkGray
        continue
    }
    if ($ListOnly) {
        Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow
        continue
    }

    # Apply back to front so the earlier offsets stay valid. A nested table's
    # own footnote row would sit inside an outer row being rewritten; drop the
    # inner edit rather than write over it twice.
    $applied = 0
    $lastStart = [int]::MaxValue
    foreach ($e in ($edits | Sort-Object Start -Descending)) {
        if ($e.End -gt $lastStart) {
            Write-Warning "  Overlapping footnote rows (nested table); inner one left alone."
            continue
        }
        $body = $body.Substring(0, $e.Start) + $e.Xml + $body.Substring($e.End)
        $lastStart = $e.Start
        $applied++
    }

    $newDoc = $head + $body

    # Fail loudly rather than write a file Word will refuse to open.
    $probe = New-Object System.Xml.XmlDocument
    $probe.PreserveWhitespace = $true
    try { $probe.LoadXml($newDoc) }
    catch { throw "Rewrite produced malformed XML for ${target}: $($_.Exception.Message)" }

    if (-not $Force -and (Test-Path -LiteralPath "$target.footnotes.bak")) {
        Write-Warning "  Backup $target.footnotes.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.footnotes.bak" -Force
    }

    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [char]0xFEFF + $newDoc } else { $newDoc }))

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "tblfootnotes-$([guid]::NewGuid()).zip")
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
    Write-Host "  Merged $applied footnote block(s). Saved." -ForegroundColor Green
}

Write-Host ""
if ($ListOnly) {
    Write-Host "Done. Nothing written. Check the listing above: a table whose last row is a" -ForegroundColor Green
    Write-Host "full-width banner or a merged total matches the same rule as a footnote." -ForegroundColor Green
} else {
    Write-Host "Done. The notes now sit in one unshaded full-width row per table." -ForegroundColor Green
}
