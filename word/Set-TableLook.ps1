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

    Hex bitmask for the legacy w:val attribute is "0420" (firstRow + noVBand).

    Also writes w:tblHeader into every first row's w:trPr so the header
    row repeats across page breaks.

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

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-TableLook.ps1 -Folder "C:\templates"

.EXAMPLE
    .\Set-TableLook.ps1 -File "Report.docx" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

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

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Canonical tblLook
# ---------------------------------------------------------------------------

$CANONICAL_TBL_LOOK = '<w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'

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
    $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
        $documentXml,
        '(?s)<w:tbl\b[^>]*>.*?</w:tbl>',
        $firstRowEvaluator)
    Write-Host "  Visited $tableCount table(s); ensured Repeat-Header-Row" -ForegroundColor DarkGray

    # 4. Sidebar callout tables: stamp the float + 3" width.
    $sidebarCount = ([regex]::Matches($documentXml, "(?s)<w:tbl\b[^>]*>(?:(?!</w:tbl>).)*?$([regex]::Escape($SIDEBAR_STYLE))")).Count
    $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
        $documentXml,
        '(?s)<w:tbl\b[^>]*>.*?</w:tbl>',
        $sidebarEvaluator)
    if ($sidebarCount -gt 0) {
        Write-Host "  Sidebar Table callouts floated (right margin, 3`"): $sidebarCount" -ForegroundColor DarkGray
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
Write-Host "Done. Open in Word, click any table, and verify under Table Design > Table Style Options." -ForegroundColor Green
