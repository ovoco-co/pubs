<#
.SYNOPSIS
    Unpacks a Word .docx/.dotx/.docm/.dotm file into a folder of XML files
    you can edit by hand.

.DESCRIPTION
    A .docx is a zip archive of XML files. This script extracts that zip
    into a folder so you can open the XML in a real editor (VSCode with
    its XML extension is recommended) and make changes. When you are done,
    use Compress-WordPackage.ps1 to repackage the folder back into a
    valid .docx.

    A backup of the original is saved alongside the file with a .bak
    extension before unpacking. If a .bak already exists you will be
    prompted (or use -Force).

.PARAMETER File
    Path to the .docx, .dotx, .docm, or .dotm file to unpack.

.PARAMETER OutputDir
    Where to extract the contents. Defaults to "<filename>_unpacked"
    next to the input file.

.PARAMETER Force
    Overwrite an existing OutputDir or .bak without prompting.

.PARAMETER OpenEditor
    After unpacking, open the OutputDir in VSCode (requires `code` on PATH).

.EXAMPLE
    .\Expand-WordPackage.ps1 -File "C:\templates\master.dotx"

.EXAMPLE
    .\Expand-WordPackage.ps1 -File "master.dotx" -OutputDir "C:\edits\master" -OpenEditor
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [switch]$Force,

    [switch]$OpenEditor
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $File -PathType Leaf)) {
    throw "File not found: $File"
}

# Resolve to a raw filesystem path, stripping any "Microsoft.PowerShell.Core\FileSystem::"
# provider prefix that PowerShell adds when CWD is a UNC path.
$resolvedPath = (Resolve-Path -LiteralPath $File).ProviderPath

# Validate by matching the path string, not via FileInfo.Extension. On some
# UNC path configurations FileInfo.Extension reads as empty even when the
# filename clearly has a recognized extension; matching the string avoids
# that whole class of bugs.
if ($resolvedPath -notmatch '\.(docx|dotx|docm|dotm)$') {
    throw "Unsupported extension. Path resolved to: $resolvedPath  (length $($resolvedPath.Length)). Expected the path to end with .docx, .dotx, .docm, or .dotm."
}

# Compute path components via Path static methods (pure string parsing).
# FileInfo returns null DirectoryName/BaseName on some UNC configurations,
# so we avoid it for path manipulation.
$dirName  = [System.IO.Path]::GetDirectoryName($resolvedPath)
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
$fileExt  = [System.IO.Path]::GetExtension($resolvedPath)

if (-not $OutputDir) {
    $OutputDir = Join-Path $dirName ($baseName + '_unpacked')
}

# Prepare output directory
if (Test-Path -LiteralPath $OutputDir) {
    if (-not $Force) {
        $reply = Read-Host "Output directory already exists: $OutputDir. Overwrite? (y/N)"
        if ($reply -notmatch '^[yY]') { throw "Aborted." }
    }
    Remove-Item -LiteralPath $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir | Out-Null

# Extract
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($resolvedPath, $OutputDir)

# Save a manifest with the original extension so Compress can recreate the right file type
$manifest = [ordered]@{
    OriginalFile      = $resolvedPath
    OriginalExtension = $fileExt
    UnpackedAt        = (Get-Date).ToString('o')
}
$manifest | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDir '.unpack-manifest.json') -Encoding UTF8

Write-Host ""
Write-Host "Unpacked to: $OutputDir" -ForegroundColor Green
Write-Host ""
Write-Host "What's inside:" -ForegroundColor Cyan

# Quick orientation
$parts = @{
    'word/styles.xml'                = 'All paragraph, character, table, and list styles'
    'word/document.xml'              = 'Main document body content'
    'word/theme/theme1.xml'          = 'Theme colors and theme fonts (this is where theme colors live)'
    'word/numbering.xml'             = 'Numbered list and bullet definitions'
    'word/settings.xml'              = 'Document-level settings (track changes default, hyphenation, etc.)'
    'word/glossary/document.xml'     = 'Building Blocks (Quick Parts, AutoText) - templates only'
    '[Content_Types].xml'            = 'Maps file extensions to MIME types - do not modify lightly'
    '_rels/.rels'                    = 'Top-level package relationships'
    'word/_rels/document.xml.rels'   = 'Document relationships (links to header/footer/numbering parts)'
}
foreach ($p in ($parts.Keys | Sort-Object)) {
    $full = Join-Path $OutputDir $p
    if (Test-Path -LiteralPath $full) {
        Write-Host "  $p" -ForegroundColor White -NoNewline
        Write-Host "  -- $($parts[$p])" -ForegroundColor DarkGray
    }
}

# Headers and footers (variable count)
$hf = Get-ChildItem -LiteralPath (Join-Path $OutputDir 'word') -Filter 'header*.xml' -ErrorAction SilentlyContinue
$ff = Get-ChildItem -LiteralPath (Join-Path $OutputDir 'word') -Filter 'footer*.xml' -ErrorAction SilentlyContinue
foreach ($h in $hf) { Write-Host "  word/$($h.Name)" -ForegroundColor White -NoNewline; Write-Host "  -- Header content for one section" -ForegroundColor DarkGray }
foreach ($f in $ff) { Write-Host "  word/$($f.Name)" -ForegroundColor White -NoNewline; Write-Host "  -- Footer content for one section" -ForegroundColor DarkGray }

# Custom XML parts
$cxp = Get-ChildItem -LiteralPath (Join-Path $OutputDir 'customXml') -ErrorAction SilentlyContinue
if ($cxp) { Write-Host "  customXml/" -ForegroundColor White -NoNewline; Write-Host "  -- Custom XML parts (data sources for content controls)" -ForegroundColor DarkGray }

Write-Host ""
Write-Host "Edit the XML files you want to change, then repackage with:" -ForegroundColor Cyan
Write-Host "  Compress-WordPackage.ps1 -SourceDir `"$OutputDir`"" -ForegroundColor Yellow
Write-Host ""

if ($OpenEditor) {
    if (Get-Command code -ErrorAction SilentlyContinue) {
        & code $OutputDir
    } else {
        Write-Warning "VSCode 'code' command not on PATH. Open the folder manually: $OutputDir"
    }
}
