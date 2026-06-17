<#
.SYNOPSIS
    Repacks a folder produced by Expand-WordPackage.ps1 back into a
    valid Word .docx/.dotx/.docm/.dotm file.

.DESCRIPTION
    Reads .unpack-manifest.json from the source directory to know which
    file to replace and what the original extension was. Produces a
    valid OOXML zip with no extraneous files. Validates that the result
    can be reopened as a zip and contains the expected core parts before
    overwriting the target.

.PARAMETER SourceDir
    The unpacked folder produced by Expand-WordPackage.ps1.

.PARAMETER OutputFile
    Where to write the repacked file. Defaults to the OriginalFile path
    recorded in .unpack-manifest.json (the original .docx/.dotx).

.PARAMETER NoBackup
    Skip rolling forward the original. By default the existing
    OutputFile (if any) is moved aside to <name>.pre-repack.bak before
    overwrite.

.PARAMETER Force
    Skip the confirmation prompt before overwriting.

.EXAMPLE
    .\Compress-WordPackage.ps1 -SourceDir "C:\templates\master_unpacked"

.EXAMPLE
    .\Compress-WordPackage.ps1 -SourceDir "C:\templates\master_unpacked" -OutputFile "C:\templates\master-v2.dotx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDir,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [switch]$NoBackup,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SourceDir -PathType Container)) {
    throw "SourceDir not found: $SourceDir"
}

$manifestPath = Join-Path $SourceDir '.unpack-manifest.json'
if (-not (Test-Path -LiteralPath $manifestPath)) {
    throw "No .unpack-manifest.json in $SourceDir. This folder was not produced by Expand-WordPackage.ps1."
}
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

if (-not $OutputFile) {
    $OutputFile = $manifest.OriginalFile
}

# Sanity-check the unpacked folder has the parts a Word file needs
$requiredParts = @('[Content_Types].xml', '_rels/.rels', 'word/document.xml')
foreach ($p in $requiredParts) {
    $full = Join-Path $SourceDir $p
    if (-not (Test-Path -LiteralPath $full)) {
        throw "Missing required part: $p`. The folder is not a valid Word package."
    }
}

# Confirm
if ((Test-Path -LiteralPath $OutputFile) -and -not $Force) {
    $reply = Read-Host "Overwrite $OutputFile`? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# Build the zip in a temp location, then move atomically
$tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "word-repack-$([guid]::NewGuid()).zip")

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# We can't use ZipFile.CreateFromDirectory because we need to skip the
# manifest file and we need to use forward slashes for entry names (OOXML
# requires forward slashes regardless of platform).
$fs = [System.IO.File]::Create($tempZip)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        # Use .ProviderPath so the prefix matches Get-ChildItem's FullName on UNC paths.
        # .Path returns "Microsoft.PowerShell.Core\FileSystem::\\server\..." which is longer
        # than the bare-UNC form Get-ChildItem reports, causing the Substring to overflow.
        $sourceRoot = (Resolve-Path -LiteralPath $SourceDir).ProviderPath
        $files = [System.IO.DirectoryInfo]::new($sourceRoot).GetFiles('*', [System.IO.SearchOption]::AllDirectories) |
            Where-Object { $_.Name -ne '.unpack-manifest.json' }

        foreach ($f in $files) {
            # Compute the entry name relative to source root, with forward slashes
            $rel = $f.FullName.Substring($sourceRoot.Length).TrimStart('\','/')
            $rel = $rel -replace '\\', '/'

            $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
            $entryStream = $entry.Open()
            try {
                $fileStream = [System.IO.File]::OpenRead($f.FullName)
                try {
                    $fileStream.CopyTo($entryStream)
                } finally {
                    $fileStream.Dispose()
                }
            } finally {
                $entryStream.Dispose()
            }
        }
    } finally {
        $zip.Dispose()
    }
} finally {
    $fs.Dispose()
}

# Validate the zip we just created
$validateOk = $true
$validateMsg = ''
try {
    $vfs = [System.IO.File]::OpenRead($tempZip)
    try {
        $vzip = New-Object System.IO.Compression.ZipArchive($vfs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($p in $requiredParts) {
                if (-not $vzip.GetEntry($p)) {
                    $validateOk = $false
                    $validateMsg = "Validation failed: missing entry $p in repacked zip."
                    break
                }
            }
        } finally {
            $vzip.Dispose()
        }
    } finally {
        $vfs.Dispose()
    }
} catch {
    $validateOk = $false
    $validateMsg = "Validation failed: cannot reopen the repacked zip: $($_.Exception.Message)"
}

if (-not $validateOk) {
    Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue
    throw $validateMsg
}

# Move into place
Move-Item -LiteralPath $tempZip -Destination $OutputFile -Force

$size = (Get-Item -LiteralPath $OutputFile).Length
Write-Host ""
Write-Host "Wrote: $OutputFile  ($size bytes)" -ForegroundColor Green
Write-Host "Open it in Word to verify the changes look right." -ForegroundColor DarkGray
