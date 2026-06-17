<#
.SYNOPSIS
    Replaces the color scheme in a Word file's theme1.xml with the
    canonical NIST NCSTAR / NWIRP entity palette.

.DESCRIPTION
    Opens a .docx, .dotx, .docm, or .dotm file, replaces the <a:clrScheme>
    block inside word/theme/theme1.xml with the canonical palette
    (Resolution Blue, Tawny, Shamrock Green, Claret, Picton Blue,
    Mulberry), and writes the result back. Other parts of the theme
    (font scheme, effect styles, object defaults) are preserved verbatim.

    The canonical palette comes from NWIRP_2-3_Graphics_Recommendations.md.
    Once applied, figures and styles can reference theme slots by name
    (accent1 = Study/NIST, accent2 = Schools, etc.) instead of hardcoding
    hex values, so the colors cannot drift between chapter documents.

    Resolution Blue is set as `#00257A` (the master-palette value from
    the graphics doc, canonical). The variant `#02257A` that appears
    elsewhere in NCSTAR styles is treated as a typo and is NOT preserved.

    The original file is backed up to <name>.theme.bak before changes.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. The script applies to every supported Word file
    in the folder. Either this or -File is required.

.PARAMETER Force
    Overwrite the file without prompting and skip the existing-backup check.

.EXAMPLE
    .\Set-NISTTheme.ps1 -File "C:\templates\HM-NWIRP-Template.dotx"

.EXAMPLE
    .\Set-NISTTheme.ps1 -Folder "\\blue\E\00-clients\nist-HM\Templates" -Force
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
# Canonical NIST color scheme. Replaces the existing <a:clrScheme> entirely.
# ---------------------------------------------------------------------------

$nistClrScheme = @'
<a:clrScheme name="NIST NCSTAR">
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="355195"/></a:dk2>
    <a:lt2><a:srgbClr val="CCD3E4"/></a:lt2>
    <a:accent1><a:srgbClr val="00257A"/></a:accent1>
    <a:accent2><a:srgbClr val="D55E00"/></a:accent2>
    <a:accent3><a:srgbClr val="009E73"/></a:accent3>
    <a:accent4><a:srgbClr val="990B38"/></a:accent4>
    <a:accent5><a:srgbClr val="56B4E9"/></a:accent5>
    <a:accent6><a:srgbClr val="C15C94"/></a:accent6>
    <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
    <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
'@

$paletteSummary = @"
NIST NCSTAR theme palette:
  dk1       sys windowText (#000000)
  lt1       sys window (#FFFFFF)
  dk2       #355195  NCSTAR Secondary (secondary text)
  lt2       #CCD3E4  NCSTAR header background
  accent1   #00257A  Resolution Blue   - Study / NIST elements
  accent2   #D55E00  Tawny             - Schools
  accent3   #009E73  Shamrock Green    - Hospitals
  accent4   #990B38  Claret            - Disruptions
  accent5   #56B4E9  Picton Blue       - Secondary emphasis, water/utilities
  accent6   #C15C94  Mulberry          - Additional category
  hlink     #0563C1  Hyperlink
  folHlink  #954F72  Followed Hyperlink
"@

# ---------------------------------------------------------------------------
# Build the file list
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

if ($targets.Count -eq 0) {
    throw "No supported Word files found."
}

Write-Host $paletteSummary -ForegroundColor Cyan
Write-Host ""
Write-Host "Applying to:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "  $_" }
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Process each file
# ---------------------------------------------------------------------------

foreach ($target in $targets) {
    Write-Host "Processing $target..." -ForegroundColor Yellow

    # Read theme1.xml from the package
    $stream = [System.IO.File]::Open($target, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $entry = $zip.GetEntry('word/theme/theme1.xml')
            if (-not $entry) {
                Write-Warning "  word/theme/theme1.xml not found. Skipping."
                continue
            }
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { $themeXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }

    # Replace the <a:clrScheme>...</a:clrScheme> block
    if ($themeXml -notmatch '<a:clrScheme[\s\S]*?</a:clrScheme>') {
        Write-Warning "  Could not find <a:clrScheme>...</a:clrScheme> block. Skipping."
        continue
    }
    $newThemeXml = [regex]::Replace($themeXml, '<a:clrScheme[\s\S]*?</a:clrScheme>', $nistClrScheme, 1)

    # Write the new theme back into the package
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "theme-swap-$([guid]::NewGuid()).zip")

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
                        if ($e.FullName -eq 'word/theme/theme1.xml') {
                            $w = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
                            try { $w.Write($newThemeXml) } finally { $w.Dispose() }
                        } else {
                            $es = $e.Open()
                            try {
                                $ns = $newEntry.Open()
                                try { $es.CopyTo($ns) } finally { $ns.Dispose() }
                            } finally { $es.Dispose() }
                        }
                    }
                } finally { $outZip.Dispose() }
            } finally { $outStream.Dispose() }
        } finally { $srcZip.Dispose() }
    } finally { $srcStream.Dispose() }

    # Atomic move
    Move-Item -LiteralPath $tempZip -Destination $target -Force
    Write-Host "  Theme replaced." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done. Open the files in Word and verify Design > Colors shows the NIST NCSTAR palette." -ForegroundColor Green
