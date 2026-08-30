<#
.SYNOPSIS
    Removes embedded (bundled) fonts from a Word file, so the document relies on
    the fonts installed on whatever machine opens or prints it.

.DESCRIPTION
    Word can bundle font files inside the .docx package (Options > Save > "Embed
    fonts in the file"). With "Embed only the characters used" off, or with
    system-font embedding on, this balloons the package: whole copies of stock
    Windows fonts (Times New Roman, Arial, Courier New, Segoe UI Emoji, ...) get
    carried along even though every Windows machine already has them.

    That bloat is a common cause of slow or timed-out PDF conversion (Acrobat
    Distiller / PDFMaker have to process every embedded font). Removing the
    embedding is safe whenever the document's fonts are installed on the machine
    doing the conversion, which is the normal case for an authoring workstation.

    This script:
      - deletes every word/fonts/*.odttf part,
      - strips the <w:embedRegular/Bold/Italic/BoldItalic> refs from
        word/fontTable.xml,
      - removes the matching font relationships from fontTable.xml.rels,
      - clears the <w:embedTrueTypeFonts>, <w:embedSystemFonts>, and
        <w:saveSubsetFonts> flags in word/settings.xml so Word does not
        re-embed on the next save,
      - drops the now-unused "odttf" Default from [Content_Types].xml.

    It never touches text, styles, or layout. The rendered document is identical
    as long as the fonts are present on the machine.

    Backs up the original to <name>.fonts.bak before writing.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. Applies to every supported Word file in the folder.
    Either this or -File is required.

.PARAMETER Force
    Overwrite without prompting and skip the existing-backup check.

.EXAMPLE
    .\Remove-EmbeddedFonts.ps1 -File "NCSTAR_4-3_Combined_nomarkup.docx"

.EXAMPLE
    .\Remove-EmbeddedFonts.ps1 -Folder "C:\Users\Artist\Documents\260729" -Force
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
        ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }

# -------- Confirm --------
Write-Host "Remove Embedded Fonts" -ForegroundColor Cyan
Write-Host "  Removes bundled font files and clears font-embedding flags." -ForegroundColor Cyan
Write-Host "  Safe when the document's fonts are installed on this machine." -ForegroundColor Cyan
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
    $order     = New-Object 'System.Collections.Generic.List[string]'
    $srcStream = [System.IO.File]::OpenRead($target)
    try {
        $srcZip = New-Object System.IO.Compression.ZipArchive($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $srcZip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
                $order.Add($e.FullName)
            }
        } finally { $srcZip.Dispose() }
    } finally { $srcStream.Dispose() }

    # 1. Drop the embedded font parts.
    $fontParts = @($entries.Keys | Where-Object { $_ -match '(?i)^word/fonts/.*\.odttf$' })
    if ($fontParts.Count -eq 0) {
        Write-Host "  No embedded fonts found. File unchanged." -ForegroundColor DarkGray
        continue
    }
    $freed = 0
    foreach ($fp in $fontParts) {
        $freed += $entries[$fp].Length
        [void]$entries.Remove($fp)
        [void]$order.Remove($fp)
    }
    Write-Host "  Removed $($fontParts.Count) embedded font part(s), freed $([math]::Round($freed/1MB,2)) MB" -ForegroundColor Gray

    # 2. Strip the embed refs from fontTable.xml.
    if ($entries.ContainsKey('word/fontTable.xml')) {
        $ft = $utf8NoBom.GetString($entries['word/fontTable.xml'])
        $ft = [regex]::Replace($ft, '(?is)<w:embed(?:Regular|Bold|Italic|BoldItalic)\b[^>]*/>', '')
        $entries['word/fontTable.xml'] = $utf8NoBom.GetBytes($ft)
    }

    # 3. Remove the font relationships.
    if ($entries.ContainsKey('word/_rels/fontTable.xml.rels')) {
        $rels = $utf8NoBom.GetString($entries['word/_rels/fontTable.xml.rels'])
        $rels = [regex]::Replace($rels, '(?is)<Relationship\b[^>]*Type="[^"]*/font"[^>]*/>', '')
        $entries['word/_rels/fontTable.xml.rels'] = $utf8NoBom.GetBytes($rels)
    }

    # 4. Clear the embedding flags in settings.xml so a later save does not
    #    re-embed. All three are empty elements (self-closing).
    if ($entries.ContainsKey('word/settings.xml')) {
        $st = $utf8NoBom.GetString($entries['word/settings.xml'])
        foreach ($flag in 'embedTrueTypeFonts','embedSystemFonts','saveSubsetFonts') {
            $st = [regex]::Replace($st, '(?is)<w:' + $flag + '\b[^>]*/>', '')
        }
        $entries['word/settings.xml'] = $utf8NoBom.GetBytes($st)
    }

    # 5. Drop the now-unused odttf Default from [Content_Types].xml.
    if ($entries.ContainsKey('[Content_Types].xml')) {
        $ct = $utf8NoBom.GetString($entries['[Content_Types].xml'])
        $ct = [regex]::Replace($ct, '(?is)<Default\b[^>]*Extension="odttf"[^>]*/>', '')
        $entries['[Content_Types].xml'] = $utf8NoBom.GetBytes($ct)
    }

    # Back up, then write the package back
    if (-not $Force -and (Test-Path -LiteralPath "$target.fonts.bak")) {
        Write-Warning "  Backup $target.fonts.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.fonts.bak" -Force
    }

    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "fonts-$([guid]::NewGuid()).zip")
    $outStream = [System.IO.File]::Create($tempZip)
    try {
        $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $order) {
                $ne = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try { $os.Write($entries[$name], 0, $entries[$name].Length) } finally { $os.Dispose() }
            }
        } finally { $outZip.Dispose() }
    } finally { $outStream.Dispose() }

    $newSize = (Get-Item -LiteralPath $tempZip).Length
    # Overwrite in place. Move-Item -Force cannot replace an existing file across
    # the Windows drive mount (DrvFs), so copy-over-then-delete instead, clearing
    # any read-only bit on the target first.
    $ti = Get-Item -LiteralPath $target -Force
    if ($ti.IsReadOnly) { $ti.IsReadOnly = $false }
    [System.IO.File]::Copy($tempZip, $target, $true)
    Remove-Item -LiteralPath $tempZip -Force
    Write-Host "  Saved. New size $([math]::Round($newSize/1MB,2)) MB" -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
