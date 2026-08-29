<#
.SYNOPSIS
    Puts every paragraph inside a table onto one paragraph style, without
    touching any of the formatting that sits on top of the style.

.DESCRIPTION
    Rewrites the <w:pStyle> of each paragraph inside every <w:tbl> to -Style
    (Table Text by default) and changes nothing else. This is deliberately
    narrower than Reapply-Styles.ps1, which wipes <w:rPr> off every run:

      Reapply-Styles.ps1          this script
      -----------------------     ---------------------------------------
      clears direct run           leaves <w:rPr> completely alone, so
      formatting everywhere       bold, italic, superscript, subscript,
                                  highlighting, and character styles all
                                  survive
      does not change pStyle      changes pStyle, and only pStyle

    Everything else in <w:pPr> survives too, which matters most for <w:ind>:
    a table whose row hierarchy is carried by indentation keeps it. Alignment
    (<w:jc>), list numbering (<w:numPr>), spacing overrides, and paragraph
    borders are all left as they are.

    Use it after a conversion, where cell paragraphs land on whatever style the
    converter chose (Body Text, or Quote via pandoc's Block Text) and need to
    be moved onto the table style without flattening the author's emphasis.

    Nested tables are included: paragraphs are matched inside the outermost
    <w:tbl> through its matching close tag, so tables inside cells are covered
    by the same pass.

    The style must already exist in styles.xml, by styleId. The script checks
    first and stops rather than writing a reference to a style that is not
    there, which would silently fall back to Normal.

    Only word/document.xml is processed. Tables in headers and footers are left
    alone, since those are template furniture rather than content.

    Idempotent: re-running changes nothing. Backs up to <name>.tablestyle.bak
    before writing.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Folder of Word files. All .docx/.dotx/.docm/.dotm in it are processed.

.PARAMETER Style
    The styleId to apply. Defaults to TableText (the "Table Text" style in the
    NIST template). This is a styleId, not a style name.

.PARAMETER UnboldColumns
    Column numbers (1-based) to strip direct bold from, in body rows only. The
    header row keeps its bold. This is the one thing the script changes beyond
    w:pStyle, and it only runs when asked for. Converted statistical tables
    carry bold group labels in column 1 that the NIST table style has no way to
    turn off, because the bold is direct run formatting rather than the style's
    First Column option.

.PARAMETER UnboldIncludesHeader
    Include the first row when stripping bold with -UnboldColumns.

.PARAMETER KeepStyles
    StyleIds to leave alone when they are already on a table paragraph. Use for
    captions or footnotes that live inside a table and should keep their own
    style. Compared case-insensitively.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-TableParagraphStyle.ps1 -File "Chapter-05.docx"

.EXAMPLE
    # Leave table captions and table footnotes on their own styles.
    .\Set-TableParagraphStyle.ps1 -Folder "C:\out" -KeepStyles CaptionTable,TableandFigureFootnotes

.EXAMPLE
    # Any other template's table style.
    .\Set-TableParagraphStyle.ps1 -File "report.docx" -Style TableGridText
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$Style = 'TableText',

    [int[]]$UnboldColumns = @(),

    [switch]$UnboldIncludesHeader,

    [string[]]$KeepStyles = @(),

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# -------- Target list --------
$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|dotx|docm|dotm)$' } |
        Sort-Object Name |
        ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }

$keepSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($k in $KeepStyles) { [void]$keepSet.Add($k) }

# -------- Confirm --------
Write-Host "Set Table Paragraph Style" -ForegroundColor Cyan
Write-Host "  Style applied to every paragraph inside a table: $Style" -ForegroundColor Cyan
Write-Host "  Left untouched: run formatting (bold, italic, super/subscript)," -ForegroundColor Cyan
Write-Host "                  indentation, alignment, numbering, spacing." -ForegroundColor Cyan
if ($keepSet.Count -gt 0) {
    Write-Host "  Styles kept as-is: $($KeepStyles -join ', ')" -ForegroundColor Cyan
}
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# w:tbl never carries attributes in practice, but allow them so the boundary
# scan cannot be fooled. The negative side of this pattern matters more: it must
# NOT match <w:tblPr>, <w:tblGrid>, or <w:tblLook>, so the character after
# "<w:tbl" has to be whitespace or the closing bracket.
$tblOpen  = '<w:tbl(?:\s[^>]*)?>'
$tblClose = '</w:tbl>'

# Paragraphs, same boundary logic so <w:pPr> is never mistaken for <w:p>.
# Paragraphs do not nest, so a lazy match to </w:p> is exact.
$paraPattern = '(?s)<w:p(?:\s[^>]*)?>.*?</w:p>'

$script:applied = 0
$script:already = 0
$script:kept    = 0

$paraEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $para    = $m.Value
    $openTag = [regex]::Match($para, '^<w:p(?:\s[^>]*)?>').Value
    $newStyle = "<w:pStyle w:val=`"$Style`"/>"

    $pPrMatch = [regex]::Match($para, '(?s)<w:pPr(?:\s[^>]*)?(?:/>|>.*?</w:pPr>)')
    if (-not $pPrMatch.Success) {
        # No pPr at all. pPr must be the first child of w:p.
        $script:applied++
        return $openTag + "<w:pPr>$newStyle</w:pPr>" + $para.Substring($openTag.Length)
    }

    $pPr = $pPrMatch.Value
    $styleMatch = [regex]::Match($pPr, '<w:pStyle\s[^>]*w:val="([^"]*)"[^>]*/>')
    if ($styleMatch.Success) {
        $current = $styleMatch.Groups[1].Value
        if ($keepSet.Contains($current)) { $script:kept++; return $para }
        if ($current -ceq $Style)        { $script:already++; return $para }
        $script:applied++
        $newPPr = $pPr.Remove($styleMatch.Index, $styleMatch.Length).Insert($styleMatch.Index, $newStyle)
    } elseif ($pPr -match '/>$') {
        # <w:pPr/> - self-closing, nothing in it to preserve.
        $script:applied++
        $newPPr = "<w:pPr>$newStyle</w:pPr>"
    } else {
        # pPr with content but no pStyle. pStyle must be pPr's first child.
        $script:applied++
        $pPrOpen = [regex]::Match($pPr, '^<w:pPr(?:\s[^>]*)?>').Value
        $newPPr  = $pPrOpen + $newStyle + $pPr.Substring($pPrOpen.Length)
    }

    return $para.Remove($pPrMatch.Index, $pPrMatch.Length).Insert($pPrMatch.Index, $newPPr)
}

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $entries   = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $srcStream = [System.IO.File]::OpenRead($target)
    try {
        $srcZip = New-Object System.IO.Compression.ZipArchive($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $srcZip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
            }
        } finally { $srcZip.Dispose() }
    } finally { $srcStream.Dispose() }

    if (-not $entries.ContainsKey('word/document.xml')) { throw "No word/document.xml in $target" }

    # A reference to a style that styles.xml does not define renders as Normal,
    # silently. Check before writing anything.
    if ($entries.ContainsKey('word/styles.xml')) {
        $stylesXml = $utf8NoBom.GetString($entries['word/styles.xml'])
        $defined = [regex]::Match($stylesXml, '<w:style\b[^>]*\bw:styleId="' + [regex]::Escape($Style) + '"')
        if (-not $defined.Success) {
            throw "Style '$Style' is not defined in $target. Apply the template's styles first (Set-NISTStyleArchitecture.ps1), or pass a -Style that exists."
        }
    } else {
        throw "No word/styles.xml in $target"
    }

    $xml = $utf8NoBom.GetString($entries['word/document.xml'])

    # Walk w:tbl open/close tags in order, tracking depth, and record the span of
    # each outermost table. Nested tables fall inside those spans, so their
    # paragraphs are covered without being visited twice.
    $boundary = [regex]::Matches($xml, "$tblOpen|$tblClose")
    $spans = New-Object System.Collections.ArrayList
    $depth = 0
    $start = -1
    foreach ($b in $boundary) {
        if ($b.Value -eq '</w:tbl>') {
            $depth--
            if ($depth -eq 0 -and $start -ge 0) {
                [void]$spans.Add(@{ Start = $start; End = $b.Index + $b.Length })
                $start = -1
            }
        } else {
            if ($depth -eq 0) { $start = $b.Index }
            $depth++
        }
    }
    if ($depth -ne 0) { throw "Unbalanced <w:tbl> tags in $target (depth $depth). File may be corrupt." }

    if ($spans.Count -eq 0) {
        Write-Host "  No tables. File unchanged." -ForegroundColor DarkGray
        continue
    }

    $script:applied = 0
    $script:already = 0
    $script:kept    = 0

    # Rebuild back to front so the recorded offsets stay valid.
    $sb = New-Object System.Text.StringBuilder $xml
    for ($i = $spans.Count - 1; $i -ge 0; $i--) {
        $s   = $spans[$i]
        $len = $s.End - $s.Start
        $chunk = $xml.Substring($s.Start, $len)
        $newChunk = [regex]::Replace($chunk, $paraPattern, $paraEvaluator)
        [void]$sb.Remove($s.Start, $len)
        [void]$sb.Insert($s.Start, $newChunk)
    }
    $newXml = $sb.ToString()

    # ---- Optional second pass: strip direct bold from named columns ----
    # Walks rows and cells so a column number means something. Plain loops, no
    # MatchEvaluator: a closure gets its own scope, so a $script: counter
    # incremented inside one never reaches the caller, and the file then looks
    # like it had nothing to change.
    $unbolded = 0
    if ($UnboldColumns.Count -gt 0) {
        $wanted = @{}
        foreach ($c in $UnboldColumns) { $wanted[[int]$c] = $true }

        $sb2  = New-Object System.Text.StringBuilder $newXml
        $base = $newXml
        # Table spans were measured on $xml. If the restyle pass changed the
        # length they no longer line up, so re-find the tables in $base.
        $tableSpans = @()
        foreach ($tm in [regex]::Matches($base, '(?s)<w:tbl(?:\s[^>]*)?>.*?</w:tbl>')) {
            $tableSpans += [pscustomobject]@{ Start = $tm.Index; Length = $tm.Length }
        }

        for ($i = $tableSpans.Count - 1; $i -ge 0; $i--) {
            $ts    = $tableSpans[$i]
            $chunk = $base.Substring($ts.Start, $ts.Length)
            $rows  = [regex]::Matches($chunk, '(?s)<w:tr(?:\s[^>]*)?>(?:(?!</w:tr>).)*</w:tr>')
            if ($rows.Count -eq 0) { continue }

            $newChunk = New-Object System.Text.StringBuilder
            $pos      = 0
            $rowIndex = 0
            foreach ($r in $rows) {
                $rowIndex++
                [void]$newChunk.Append($chunk.Substring($pos, $r.Index - $pos))
                $pos = $r.Index + $r.Length

                if ($rowIndex -eq 1 -and -not $UnboldIncludesHeader) {
                    [void]$newChunk.Append($r.Value)
                    continue
                }

                $row     = $r.Value
                $rowOut  = New-Object System.Text.StringBuilder
                $cellPos = 0
                $col     = 0
                foreach ($cm in [regex]::Matches($row, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')) {
                    $col++
                    $cell = $cm.Value
                    if ($wanted.ContainsKey($col)) {
                        $stripped = [regex]::Replace($cell, '<w:b(?:\s+w:val="(?:1|true|on)")?\s*/>', '')
                        $stripped = [regex]::Replace($stripped, '<w:bCs(?:\s+w:val="(?:1|true|on)")?\s*/>', '')
                        if ($stripped -ne $cell) { $unbolded++; $cell = $stripped }
                    }
                    [void]$rowOut.Append($row.Substring($cellPos, $cm.Index - $cellPos))
                    [void]$rowOut.Append($cell)
                    $cellPos = $cm.Index + $cm.Length
                    # A cell spanning N columns advances the count by N.
                    $gs = [regex]::Match($cm.Value, '<w:gridSpan\s[^>]*w:val="(\d+)"')
                    if ($gs.Success) { $col += ([int]$gs.Groups[1].Value - 1) }
                }
                [void]$rowOut.Append($row.Substring($cellPos))
                [void]$newChunk.Append($rowOut.ToString())
            }
            [void]$newChunk.Append($chunk.Substring($pos))

            [void]$sb2.Remove($ts.Start, $ts.Length)
            [void]$sb2.Insert($ts.Start, $newChunk.ToString())
        }
        $newXml = $sb2.ToString()
    }

    Write-Host "  Tables: $($spans.Count)" -ForegroundColor Gray
    if ($UnboldColumns.Count -gt 0) {
        Write-Host "  Bold stripped from column(s) $($UnboldColumns -join ', '): $unbolded cell(s)" -ForegroundColor Gray
    }
    Write-Host "  Paragraphs restyled to ${Style}: $script:applied" -ForegroundColor Gray
    Write-Host "  Already on ${Style}: $script:already" -ForegroundColor DarkGray
    if ($script:kept -gt 0) { Write-Host "  Left on a -KeepStyles style: $script:kept" -ForegroundColor DarkGray }

    if ($script:applied -eq 0 -and $unbolded -eq 0) {
        Write-Host "  Nothing to change. File unchanged." -ForegroundColor DarkGray
        continue
    }

    if (-not $Force -and (Test-Path -LiteralPath "$target.tablestyle.bak")) {
        Write-Warning "  Backup $target.tablestyle.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.tablestyle.bak" -Force
    }

    $entries['word/document.xml'] = $utf8NoBom.GetBytes($newXml)

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "tablestyle-$([guid]::NewGuid()).zip")
    $outStream = [System.IO.File]::Create($tempZip)
    try {
        $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $entries.Keys) {
                $ne = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try { $os.Write($entries[$name], 0, $entries[$name].Length) } finally { $os.Dispose() }
            }
        } finally { $outZip.Dispose() }
    } finally { $outStream.Dispose() }

    Move-Item -LiteralPath $tempZip -Destination $target -Force
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
