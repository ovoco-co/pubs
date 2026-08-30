<#
.SYNOPSIS
    Fixes typography in a Word file: converts straight quotes to curly
    (typographic) quotes, and optionally straightens other marks. Operates
    only on visible text, never on field codes or markup.

.DESCRIPTION
    Authored or pasted Word content often has straight quotes (" and ')
    instead of the curly quotes a published report wants. This converts them:

      "  ->  U+201C / U+201D  (left / right double quotation mark)
      '  ->  U+2018 / U+2019  (left / right single quotation mark)

    Opening vs closing is decided by context, tracked ACROSS run and paragraph
    boundaries so a quote that starts a run (or follows a tab or line break)
    still opens correctly:
      - Start of a paragraph, or after whitespace / ( [ { / an opening curly
        quote / an en or em dash  ->  opening quote.
      - Otherwise (after a letter, digit, or closing punctuation)  ->  closing
        quote, i.e. also the apostrophe form for contractions and possessives
        (it's, teachers').
      - A single quote immediately before a digit (the '90s) is treated as an
        apostrophe (U+2019), not an opening quote.

    Only the text inside <w:t> runs is changed. Field codes (<w:instrText>,
    e.g. STYLEREF "Heading 1"), element tags, and attributes are never touched,
    so quotes that are syntax stay straight. Already-curly quotes are left
    alone, so re-running is a no-op.

    Applies to the text-bearing parts: document body, headers, footers,
    footnotes, endnotes. styles.xml is never touched.

    Backs up the original to <name>.typography.bak before writing.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. Applies to every supported Word file in the folder.
    Either this or -File is required.

.PARAMETER Ellipses
    Also convert three straight dots (...) to a single ellipsis character
    (U+2026). Off by default.

.PARAMETER Force
    Overwrite without prompting and skip the existing-backup check.

.EXAMPLE
    .\Repair-Typography.ps1 -File "Chapter-02.docx"

.EXAMPLE
    .\Repair-Typography.ps1 -Folder "C:\Users\Artist\Documents\260729" -Force

.EXAMPLE
    .\Repair-Typography.ps1 -File "Chapter-02.docx" -Ellipses
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$Ellipses,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# Curly-quote code points.
$LDQUO = [char]0x201C; $RDQUO = [char]0x201D    # " "
$LSQUO = [char]0x2018; $RSQUO = [char]0x2019    # ' '
$HELLIP = [char]0x2026                          # ...

# Characters after which a quote OPENS (in addition to whitespace / start).
$opensAfter = "([{" + [char]0x2018 + [char]0x201C + [char]0x2013 + [char]0x2014 + [char]0x00A0

function Test-QuoteOpens {
    param([char]$Prev)
    if ($Prev -eq [char]0) { return $true }              # start of paragraph / segment
    if ([char]::IsWhiteSpace($Prev)) { return $true }
    return ($opensAfter.IndexOf($Prev) -ge 0)
}

function Convert-QuoteRun {
    # Convert straight quotes (and optionally ellipses) in one run's text,
    # threading the "previous character" through $St so context carries between
    # runs. $St = @{ Prev=<char>; Count=<int> }.
    param([string]$Text, [hashtable]$St)
    $sb = New-Object System.Text.StringBuilder
    $n = $Text.Length
    for ($i = 0; $i -lt $n; $i++) {
        $ch = $Text[$i]
        if ($ch -eq '"') {
            $o = if (Test-QuoteOpens $St.Prev) { $LDQUO } else { $RDQUO }
            [void]$sb.Append($o); $St.Prev = $o; $St.Count++
        } elseif ($ch -eq "'") {
            $nx = if ($i + 1 -lt $n) { $Text[$i + 1] } else { [char]0 }
            $o = if ([char]::IsDigit($nx)) { $RSQUO }
                 elseif (Test-QuoteOpens $St.Prev) { $LSQUO }
                 else { $RSQUO }
            [void]$sb.Append($o); $St.Prev = $o; $St.Count++
        } else {
            [void]$sb.Append($ch); $St.Prev = $ch
        }
    }
    $r = $sb.ToString()
    if ($Ellipses) {
        $e = ([regex]::Matches($r, '\.\.\.')).Count
        if ($e -gt 0) { $r = $r -replace '\.\.\.', $HELLIP; $St.Count += $e }
    }
    return $r
}

function Convert-Part {
    # Walk one part's XML, converting quotes inside <w:t> runs only. Paragraph
    # starts and line breaks / tabs reset the "previous character" so the next
    # quote opens. Everything else (markup, field codes) is copied verbatim.
    param([string]$Xml)
    $pattern = '(?s)<w:t\b[^>]*>.*?</w:t>|<w:p\b[^>]*>|</w:p>|<w:br\b[^>]*?/?>|<w:tab\b[^>]*?/?>|<w:cr\b[^>]*?/?>'
    $matches = [regex]::Matches($Xml, $pattern)
    if ($matches.Count -eq 0) { return @{ Xml = $Xml; Count = 0 } }

    $sb  = New-Object System.Text.StringBuilder
    $pos = 0
    $St  = @{ Prev = [char]0; Count = 0 }

    foreach ($m in $matches) {
        if ($m.Index -gt $pos) { [void]$sb.Append($Xml.Substring($pos, $m.Index - $pos)) }
        $v = $m.Value
        if ($v.EndsWith('</w:t>')) {
            $mm    = [regex]::Match($v, '(?s)^(<w:t\b[^>]*>)(.*)(</w:t>)$')
            $open  = $mm.Groups[1].Value
            $inner = $mm.Groups[2].Value
            $close = $mm.Groups[3].Value
            [void]$sb.Append($open)
            [void]$sb.Append((Convert-QuoteRun -Text $inner -St $St))
            [void]$sb.Append($close)
        } else {
            # Paragraph boundary or tab / break: reset context so the next quote opens.
            $St.Prev = [char]0
            [void]$sb.Append($v)
        }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $Xml.Length) { [void]$sb.Append($Xml.Substring($pos)) }
    return @{ Xml = $sb.ToString(); Count = $St.Count }
}

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
        ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }

# -------- Confirm --------
$fixes = @('straight quotes -> curly')
if ($Ellipses) { $fixes += '... -> ellipsis' }
Write-Host "Repair Typography" -ForegroundColor Cyan
Write-Host "  Fixes: $($fixes -join ', ')." -ForegroundColor Cyan
Write-Host "  Only <w:t> text is changed; field codes and markup are left alone." -ForegroundColor Cyan
Write-Host "  Parts: document body, headers, footers, footnotes, endnotes." -ForegroundColor Cyan
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
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

    $textParts = @($entries.Keys | Where-Object {
        $_ -match '^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$'
    })

    $totalFixed = 0
    foreach ($part in $textParts) {
        $xml    = $utf8NoBom.GetString($entries[$part])
        $result = Convert-Part -Xml $xml
        if ($result.Count -gt 0) {
            $entries[$part] = $utf8NoBom.GetBytes($result.Xml)
            $totalFixed += $result.Count
            Write-Host "  $part : fixed $($result.Count)" -ForegroundColor Gray
        }
    }

    if ($totalFixed -eq 0) {
        Write-Host "  Nothing to fix. File unchanged." -ForegroundColor DarkGray
        continue
    }
    Write-Host "  Total fixed: $totalFixed" -ForegroundColor Green

    if (-not $Force -and (Test-Path -LiteralPath "$target.typography.bak")) {
        Write-Warning "  Backup $target.typography.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.typography.bak" -Force
    }

    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "typography-$([guid]::NewGuid()).zip")
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
