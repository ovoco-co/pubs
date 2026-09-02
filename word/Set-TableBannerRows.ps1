<#
.SYNOPSIS
    Shades and bolds the banner rows inside a table - the full-width merged
    rows that carry a group label rather than data.

.DESCRIPTION
    A statistical table breaks its body into groups, and the group label sits
    in a row of its own that spans every column:

        Ownership (n = 297)
          Public                 96      21.29
          Private                58      12.86

    Those banner rows read as headings inside the table, so they carry a fill
    and bold. A converter emits them as plain merged rows with neither.

    The catch is that a table's FOOTNOTE rows are full-width merges too, and
    they must not be shaded - the note under a table is not a heading. This
    script tells them apart by paragraph style: run Merge-TableFootnotes.ps1
    first, which puts the notes on TableandFigureFootnotes, and everything
    still full-width and unclaimed after that is a banner. -ExcludeStyles adds
    more styles to that exclusion, and -SkipTrailing also leaves the rows at
    the very foot of a table alone, whatever style they are on, for a document
    where the footnote merge has not been run.

    Run -ListOnly first. The rule is structural: a merged "Total" row at the
    bottom of a table looks exactly like a group label to it.

    Idempotent. Backs up to <name>.banners.bak before writing.

    Uses System.IO.Compression only. Word is not required.

.PARAMETER File
    Single Word file to process.

.PARAMETER Folder
    Folder of Word files. All .docx/.docm files are processed.

.PARAMETER Fill
    Cell fill for a banner row, as six hex digits, no hash. Default CCECE3.
    Pass 'auto' to apply no fill and only set the bold. This is the fallback
    for any table no -FillRules entry claims.

.PARAMETER FillRules
    Caption-driven fills, so one document can colour-code its tables by which
    population they describe. Each entry is 'Pattern=HEX', where Pattern is a
    regex matched case-insensitively against the caption above the table and
    HEX is six hex digits or 'auto'. The first entry that matches wins, so
    order them most specific first. A table with no caption of its own
    inherits the nearest caption above it, which is what a continuation table
    needs.

        -FillRules 'Schools=F7DFCC','Hospitals=CCECE3'

    Anything unmatched falls back to -Fill.

.PARAMETER CaptionStyles
    Paragraph styleIds -FillRules reads a table's caption from. Defaults to
    CaptionTable, Caption, and Heading1 through Heading3.

.PARAMETER NoBold
    Leave the run formatting alone and only apply the fill.

.PARAMETER ExcludeStyles
    Paragraph styleIds that mark a full-width row as something other than a
    banner. Defaults to TableandFigureFootnotes and CaptionTable.

.PARAMETER SkipTrailing
    Also leave the run of full-width rows at the foot of each table alone.
    Use on a document whose footnote rows have not been put on their own style
    by Merge-TableFootnotes.ps1 yet.

.PARAMETER ListOnly
    Report which rows would be treated as banners and write nothing.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-TableBannerRows.ps1 -File "Chapter-05.docx" -ListOnly

.EXAMPLE
    .\Set-TableBannerRows.ps1 -File "Chapter-05.docx" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$Fill = 'CCECE3',

    [string[]]$FillRules = @(),

    [string[]]$CaptionStyles = @('CaptionTable', 'Caption', 'Heading1', 'Heading2', 'Heading3'),

    [switch]$NoBold,

    [string[]]$ExcludeStyles = @('TableandFigureFootnotes', 'CaptionTable'),

    [switch]$SkipTrailing,

    [switch]$ListOnly,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if ($Fill -ne 'auto' -and $Fill -notmatch '^[0-9A-Fa-f]{6}$') {
    throw "-Fill must be six hex digits with no hash, or 'auto'. Got: $Fill"
}

# Ordered, so the first matching pattern wins.
$rules = New-Object System.Collections.ArrayList
# powershell.exe -File does not split a comma-separated argument into an
# array, so 'A=1','B=2' arrives as the single string 'A=1,B=2'. Split it back.
$ruleEntries = @()
foreach ($raw in $FillRules) {
    if (-not $raw) { continue }
    $ruleEntries += [regex]::Split($raw, ',(?=[^,]*=)')
}

foreach ($entry in $ruleEntries) {
    if (-not $entry) { continue }
    $eq = $entry.LastIndexOf('=')
    if ($eq -lt 1 -or $eq -eq $entry.Length - 1) {
        throw "-FillRules entries must look like 'Pattern=HEX'. Got: $entry"
    }
    $pattern = $entry.Substring(0, $eq).Trim()
    $hex     = $entry.Substring($eq + 1).Trim()
    if ($hex -ne 'auto' -and $hex -notmatch '^[0-9A-Fa-f]{6}$') {
        throw "-FillRules fill must be six hex digits with no hash, or 'auto'. Got: $hex (in '$entry')"
    }
    try { [void][regex]::new($pattern) }
    catch { throw "-FillRules pattern is not a valid regex: $pattern (in '$entry')" }
    [void]$rules.Add([pscustomobject]@{ Pattern = $pattern; Fill = $hex })
}

$captionStyleAlt = ($CaptionStyles | ForEach-Object { [regex]::Escape($_) }) -join '|'

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

Write-Host "Set Table Banner Rows" -ForegroundColor Cyan
Write-Host "  Full-width merged rows that are not footnotes or captions." -ForegroundColor Cyan
Write-Host ("  Fill: {0}   Bold: {1}" -f $Fill, $(if ($NoBold) { 'left alone' } else { 'applied' })) -ForegroundColor Cyan
Write-Host ("  Excluded styles: {0}" -f ($ExcludeStyles -join ', ')) -ForegroundColor Cyan
if ($rules.Count -gt 0) {
    Write-Host ("  Caption rules (first match wins), read from: {0}" -f ($CaptionStyles -join ', ')) -ForegroundColor Cyan
    foreach ($r in $rules) { Write-Host ("    /{0}/ -> {1}" -f $r.Pattern, $r.Fill) -ForegroundColor Cyan }
    Write-Host ("    unmatched -> {0}" -f $Fill) -ForegroundColor Cyan
}
if ($SkipTrailing) { Write-Host "  Trailing full-width rows left alone" -ForegroundColor Cyan }
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""
if ($ListOnly) { Write-Host "  -ListOnly: reporting matches, writing nothing." -ForegroundColor Yellow }

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

function New-Shd {
    param([string]$Hex)
    if ($Hex -eq 'auto') { return '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>' }
    return '<w:shd w:val="clear" w:color="auto" w:fill="' + $Hex.ToUpperInvariant() + '"/>'
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
    if ($rules.Count -gt 0) {
        foreach ($pm in [regex]::Matches($body, '(?s)<w:p(?:\s[^>]*)?>(?:(?!</w:p>).)*</w:p>')) {
            if ($pm.Value -notmatch ('<w:pStyle\s[^>]*w:val="(?:' + $captionStyleAlt + ')"')) { continue }
            $ct = Get-RowText -Xml $pm.Value
            if ($ct) { [void]$captions.Add([pscustomobject]@{ Index = $pm.Index; Text = $ct }) }
        }
    }

    $edits   = New-Object System.Collections.ArrayList
    $banners = 0
    $shown   = 0
    $groups  = @{}

    foreach ($tm in [regex]::Matches($body, '(?s)<w:tbl(?:\s[^>]*)?>.*?</w:tbl>')) {
        $table = $tm.Value
        $ncols = ([regex]::Matches($table, '<w:gridCol\b')).Count
        if ($ncols -lt 2) { continue }

        $rows = [regex]::Matches($table, '(?s)<w:tr(?:\s[^>]*)?>(?:(?!</w:tr>).)*</w:tr>')
        if ($rows.Count -lt 2) { continue }

        # Which fill this table's banners get. Fall back to -Fill.
        $tableFill  = $Fill
        $tableRule  = ''
        $tableCap   = ''
        if ($rules.Count -gt 0) {
            for ($c = $captions.Count - 1; $c -ge 0; $c--) {
                if ($captions[$c].Index -lt $tm.Index) { $tableCap = $captions[$c].Text; break }
            }
            foreach ($r in $rules) {
                if ($tableCap -match $r.Pattern) { $tableFill = $r.Fill; $tableRule = $r.Pattern; break }
            }
        }
        $shd = New-Shd -Hex $tableFill

        # Which trailing rows to leave alone when the footnote style is absent.
        $trailingFrom = $rows.Count
        if ($SkipTrailing) {
            for ($i = $rows.Count - 1; $i -ge 0; $i--) {
                $rx = $rows[$i].Value
                $cs = ([regex]::Matches($rx, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')).Count
                $gs = [regex]::Match($rx, '<w:gridSpan\s[^>]*w:val="(\d+)"')
                if ($cs -eq 1 -and $gs.Success -and [int]$gs.Groups[1].Value -ge $ncols) { $trailingFrom = $i }
                else { break }
            }
        }

        for ($i = 0; $i -lt $rows.Count; $i++) {
            if ($i -eq 0) { continue }                    # header row is not a banner
            if ($i -ge $trailingFrom) { continue }
            $row = $rows[$i].Value

            $cells = [regex]::Matches($row, '(?s)<w:tc(?:\s[^>]*)?>(?:(?!</w:tc>).)*</w:tc>')
            if ($cells.Count -ne 1) { continue }
            $gs = [regex]::Match($row, '<w:gridSpan\s[^>]*w:val="(\d+)"')
            if (-not $gs.Success -or [int]$gs.Groups[1].Value -lt $ncols) { continue }

            $text = Get-RowText -Xml $row
            if (-not $text) { continue }

            $styles = [regex]::Matches($row, '<w:pStyle\s[^>]*w:val="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
            $skip = $false
            foreach ($st in $styles) { if ($ExcludeStyles -contains $st) { $skip = $true; break } }
            if ($skip) { continue }

            $banners++
            if ($rules.Count -gt 0) {
                $key = if ($tableRule) { "$tableRule -> $tableFill" } else { "(unmatched) -> $tableFill" }
                if (-not $groups.ContainsKey($key)) { $groups[$key] = 0 }
                $groups[$key]++
            }
            if ($shown -lt 6) {
                Write-Host ("    {0}" -f $text.Substring(0, [Math]::Min(64, $text.Length))) -ForegroundColor DarkGray
                $shown++
            }

            # Rebuild the one cell: canonical fill, and bold on every run.
            $cell = $cells[0].Value
            $new  = [regex]::Replace($cell, '<w:shd\b[^>]*/>', '')
            if ($new -match '(?s)<w:tcPr(?:\s[^>]*)?>.*?</w:tcPr>') {
                $new = [regex]::Replace($new, '(?s)(<w:tcPr(?:\s[^>]*)?>)(.*?)(</w:tcPr>)', ('${1}${2}' + $shd + '${3}'), 1)
            } elseif ($new -match '<w:tcPr\s*/>') {
                $new = [regex]::Replace($new, '<w:tcPr\s*/>', "<w:tcPr>$shd</w:tcPr>", 1)
            } else {
                $new = [regex]::Replace($new, '^(<w:tc(?:\s[^>]*)?>)', ('${1}<w:tcPr>' + $shd + '</w:tcPr>'), 1)
            }

            if (-not $NoBold) {
                # w:b is the first child of w:rPr in schema order, and rPr must
                # be the first child of w:r.
                $new = [regex]::Replace($new, '(?s)<w:r(?:\s[^>]*)?>(?:(?!</w:r>).)*</w:r>', {
                    param($rm)
                    $run = $rm.Value
                    if ($run -notmatch '<w:t[\s>]') { return $run }       # no text, leave it
                    if ($run -match '<w:b\s*/>' -or $run -match '<w:b\s+w:val=') { return $run }
                    if ($run -match '(?s)<w:rPr(?:\s[^>]*)?>.*?</w:rPr>') {
                        return [regex]::Replace($run, '(<w:rPr(?:\s[^>]*)?>)', '${1}<w:b/><w:bCs/>', 1)
                    }
                    if ($run -match '<w:rPr\s*/>') {
                        return [regex]::Replace($run, '<w:rPr\s*/>', '<w:rPr><w:b/><w:bCs/></w:rPr>', 1)
                    }
                    return [regex]::Replace($run, '^(<w:r(?:\s[^>]*)?>)', '${1}<w:rPr><w:b/><w:bCs/></w:rPr>', 1)
                })
            }

            if ($new -ne $cell) {
                [void]$edits.Add([pscustomobject]@{
                    Start = $tm.Index + $rows[$i].Index + $cells[0].Index
                    End   = $tm.Index + $rows[$i].Index + $cells[0].Index + $cells[0].Length
                    Xml   = $new
                })
            }
        }
    }

    Write-Host ("  Banner rows found: {0}  |  rewritten: {1}" -f $banners, $edits.Count) -ForegroundColor DarkGray
    foreach ($k in ($groups.Keys | Sort-Object)) {
        Write-Host ("    {0}  {1} row(s)" -f $k, $groups[$k]) -ForegroundColor DarkGray
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

    if (-not $Force -and (Test-Path -LiteralPath "$target.banners.bak")) {
        Write-Warning "  Backup $target.banners.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.banners.bak" -Force
    }

    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [char]0xFEFF + $newDoc } else { $newDoc }))

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "banners-$([guid]::NewGuid()).zip")
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
