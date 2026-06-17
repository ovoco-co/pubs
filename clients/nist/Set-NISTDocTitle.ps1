<#
.SYNOPSIS
    Sets the document title metadata (<dc:title> in docProps/core.xml) on
    a Word file. The title is what Word's File > Info > Properties >
    Title field shows, and what the { TITLE } field code in a header or
    footer pulls.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. Applies to every Word file in the folder.

.PARAMETER Title
    The title string to set. Required.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-NISTDocTitle.ps1 -File "HM-NWIRP-Template.dotx" -Title "NIST NWIRP 2-3, Hurricane Maria Investigation"

.EXAMPLE
    .\Set-NISTDocTitle.ps1 -Folder "C:\templates" -Title "NIST NCSTAR 4, Hurricane Maria Investigation"
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [Parameter(Mandatory=$true)]
    [string]$Title,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

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

$encodedTitle = [System.Web.HttpUtility]::HtmlEncode($Title)

Write-Host "Setting document title to: $Title" -ForegroundColor Cyan
Write-Host "Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "  $_" }
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    # Read core.xml
    $coreXml = $null
    $stream = [System.IO.File]::Open($target, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $entry = $zip.GetEntry('docProps/core.xml')
            if (-not $entry) { Write-Warning "  No docProps/core.xml. Skipping."; continue }
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { $coreXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }

    # Replace or insert <dc:title>
    if ($coreXml -match '<dc:title\b[^>]*>.*?</dc:title>') {
        $coreXml = [regex]::Replace($coreXml, '<dc:title\b[^>]*>.*?</dc:title>', "<dc:title>$encodedTitle</dc:title>", 1)
    } elseif ($coreXml -match '<dc:title\s*/>') {
        $coreXml = [regex]::Replace($coreXml, '<dc:title\s*/>', "<dc:title>$encodedTitle</dc:title>", 1)
    } else {
        # Insert before </cp:coreProperties> (the root element close)
        if ($coreXml -match '</cp:coreProperties>') {
            $coreXml = [regex]::Replace($coreXml, '</cp:coreProperties>', "<dc:title>$encodedTitle</dc:title></cp:coreProperties>", 1)
        } else {
            Write-Warning "  Cannot insert title (no </cp:coreProperties> found). Skipping."
            continue
        }
    }

    # Write the modified core.xml back into the package
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "doctitle-$([guid]::NewGuid()).zip")
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
                        if ($e.FullName -eq 'docProps/core.xml') {
                            $w = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
                            try { $w.Write($coreXml) } finally { $w.Dispose() }
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

    Move-Item -LiteralPath $tempZip -Destination $target -Force
    Write-Host "  Title set." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done. Verify in Word: File > Info > Properties > Title." -ForegroundColor Green
