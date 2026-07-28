<#
.SYNOPSIS
    Removes direct (locally applied) shading from text and paragraphs in a Word
    file, so runs and paragraphs show no background. Style-defined shading is
    kept; table-cell fills are kept unless -IncludeTables is given.

.DESCRIPTION
    Word documents pick up stray shading applied as DIRECT formatting - most
    often "Pattern: Clear (White)", a solid fill behind text or a bullet - from
    pasting, conversions, or template imports. It shows as a blank box and is
    almost always a mistake: text does not need a background.

    The rule this script follows:
      - Direct shading on a run (local character formatting) -> removed.
      - Direct shading on a paragraph (local paragraph formatting) -> removed.
      - Shading that comes from a paragraph or character STYLE -> kept, because
        it lives in styles.xml, which this script never touches. If shading is
        part of a style, it belongs.
      - Table-CELL fills (w:tcPr) -> kept by default, since a cell fill may be
        intentional. Pass -IncludeTables to strip those too.

    Applies to the text-bearing parts: document body, headers, footers,
    footnotes, endnotes. The <w:shd> element is always empty (self-closing),
    so removing it only drops the background; no text or structure changes.

    Backs up the original to <name>.shading.bak before writing.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. Applies to every supported Word file in the folder.
    Either this or -File is required.

.PARAMETER IncludeTables
    Also remove direct shading from table cells and whole-table properties,
    not just runs and paragraphs.

.PARAMETER Force
    Overwrite without prompting and skip the existing-backup check.

.EXAMPLE
    .\Remove-Shading.ps1 -File "Chapter-02.docx"

.EXAMPLE
    .\Remove-Shading.ps1 -Folder "C:\Users\Artist\Documents\260726" -Force

.EXAMPLE
    .\Remove-Shading.ps1 -File "Chapter-02.docx" -IncludeTables
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$IncludeTables,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# A <w:shd> is always an empty element, so only the self-closing form exists.
$shdPattern = '(?is)<w:shd\b[^>]*/>'
# Strip every <w:shd> found inside a matched block (rPr / pPr / tcPr).
$stripShd = { param($m) [regex]::Replace($m.Value, $shdPattern, '') }

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
$scope = if ($IncludeTables) { "runs, paragraphs, and table cells" } else { "runs and paragraphs (table cells kept)" }
Write-Host "Remove Shading" -ForegroundColor Cyan
Write-Host "  Removes direct shading from: $scope." -ForegroundColor Cyan
Write-Host "  Style-defined shading is kept (styles.xml is never touched)." -ForegroundColor Cyan
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

    # Read all entries into memory
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

    # Text-bearing parts only. Never touch styles.xml.
    $textParts = @($entries.Keys | Where-Object {
        $_ -match '^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$'
    })

    $totalRemoved = 0
    foreach ($part in $textParts) {
        $xml    = $utf8NoBom.GetString($entries[$part])
        $before = ([regex]::Matches($xml, '<w:shd\b')).Count

        # Direct run shading (local character formatting), including paragraph marks.
        $xml = [regex]::Replace($xml, '(?is)<w:rPr\b[^>]*>.*?</w:rPr>', $stripShd)
        # Direct paragraph shading (local paragraph formatting).
        $xml = [regex]::Replace($xml, '(?is)<w:pPr\b[^>]*>.*?</w:pPr>', $stripShd)
        # Optionally table-cell and whole-table shading.
        if ($IncludeTables) {
            $xml = [regex]::Replace($xml, $shdPattern, '')
        }

        $after   = ([regex]::Matches($xml, '<w:shd\b')).Count
        $removed = $before - $after
        if ($removed -gt 0) {
            $entries[$part] = $utf8NoBom.GetBytes($xml)
            $totalRemoved += $removed
            Write-Host "  $part : removed $removed" -ForegroundColor Gray
        }
    }

    if ($totalRemoved -eq 0) {
        Write-Host "  No direct text/paragraph shading found. File unchanged." -ForegroundColor DarkGray
        continue
    }
    Write-Host "  Total removed: $totalRemoved" -ForegroundColor Green

    # Back up, then write the package back
    if (-not $Force -and (Test-Path -LiteralPath "$target.shading.bak")) {
        Write-Warning "  Backup $target.shading.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.shading.bak" -Force
    }

    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "shading-$([guid]::NewGuid()).zip")
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
