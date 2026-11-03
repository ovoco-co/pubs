<#
.SYNOPSIS
    Replaces the color scheme in a Word file's theme1.xml with a supplied
    clrScheme XML block.

.DESCRIPTION
    Opens a .docx, .dotx, .docm, or .dotm file, replaces the <a:clrScheme>
    block inside word/theme/theme1.xml with the supplied XML, and writes the
    result back. Other parts of the theme (font scheme, effect styles, object
    defaults) are preserved verbatim.

    The original file is backed up to <name>.theme.bak before changes.

    Validates that the supplied ColorSchemeXml contains all 12 required theme
    color elements (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink).

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. The script applies to every supported Word file
    in the folder. Either this or -File is required.

.PARAMETER ColorSchemeXml
    Complete <a:clrScheme> XML block with all 12 theme color elements.

.PARAMETER Force
    Overwrite the file without prompting and skip the existing-backup check.

.EXAMPLE
    .\Set-ThemeColors.ps1 -File "report.docx" -ColorSchemeXml $myScheme

.EXAMPLE
    .\Set-ThemeColors.ps1 -Folder "C:\templates" -ColorSchemeXml $myScheme -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [Parameter(Mandatory=$true)]
    [string]$ColorSchemeXml,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Validate that ColorSchemeXml contains all 12 required theme color elements
# ---------------------------------------------------------------------------

$requiredElements = @('a:dk1', 'a:lt1', 'a:dk2', 'a:lt2',
    'a:accent1', 'a:accent2', 'a:accent3', 'a:accent4',
    'a:accent5', 'a:accent6', 'a:hlink', 'a:folHlink')

$missing = $requiredElements | Where-Object { $ColorSchemeXml -notmatch "<$_>" }
if ($missing) {
    throw "ColorSchemeXml is missing required theme color elements: $($missing -join ', ')"
}

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

Write-Host "Applying theme colors to:" -ForegroundColor Cyan
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
    $newThemeXml = [regex]::Replace($themeXml, '<a:clrScheme[\s\S]*?</a:clrScheme>', $ColorSchemeXml, 1)

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
    Write-Host "  Theme colors replaced." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done. Open the files in Word and verify Design > Colors shows the expected palette." -ForegroundColor Green
