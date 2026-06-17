<#
.SYNOPSIS
    Sets page margins and header/footer distance on every section in a
    Word document.

.DESCRIPTION
    Updates every w:pgMar element (one per section) in word/document.xml
    with the specified margin values. Supports mirror-margin layouts where
    w:left maps to the inner (binding) side and w:right to the outer side.

    Default values are 1.0" for all body margins and 0.3" for header and
    footer distance - a common formal report layout. Override any value
    via parameters.

    Operates directly on the ZIP package via System.IO.Compression - no
    Word installation required. Idempotent: re-running with the same
    values produces the same result.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. The script applies to every supported Word file
    in the folder.

.PARAMETER InnerInches
    Inner (binding-side) margin in inches. Default 1.25. Maps to w:left.

.PARAMETER OuterInches
    Outer (non-binding) margin in inches. Default 0.75. Maps to w:right.

.PARAMETER TopInches
    Top margin in inches. Default 1.0.

.PARAMETER BottomInches
    Bottom margin in inches. Default 1.0.

.PARAMETER HeaderInches
    Distance in inches from the top of the page to the header. Default 0.3.

.PARAMETER FooterInches
    Distance in inches from the bottom of the page to the footer. Default 0.3.

.PARAMETER Force
    Skip prompts.

.EXAMPLE
    .\Set-PageLayout.ps1 -Folder "C:\templates"

.EXAMPLE
    .\Set-PageLayout.ps1 -File "Report-Template.dotx" -InnerInches 1.5 -OuterInches 0.75
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [double]$InnerInches  = 1.0,
    [double]$OuterInches  = 1.0,
    [double]$TopInches    = 1.0,
    [double]$BottomInches = 1.0,
    [double]$HeaderInches = 0.3,
    [double]$FooterInches = 0.3,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$leftTwips   = [int]([math]::Round($InnerInches  * 1440))
$rightTwips  = [int]([math]::Round($OuterInches  * 1440))
$topTwips    = [int]([math]::Round($TopInches    * 1440))
$bottomTwips = [int]([math]::Round($BottomInches * 1440))
$headerTwips = [int]([math]::Round($HeaderInches * 1440))
$footerTwips = [int]([math]::Round($FooterInches * 1440))

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

Write-Host "Page-layout values (mirror margins assumed; w:left = inner, w:right = outer):" -ForegroundColor Cyan
Write-Host ("  Inner (binding):  {0,5} in ({1} twips)" -f $InnerInches, $leftTwips)   -ForegroundColor Cyan
Write-Host ("  Outer:            {0,5} in ({1} twips)" -f $OuterInches, $rightTwips)  -ForegroundColor Cyan
Write-Host ("  Top:              {0,5} in ({1} twips)" -f $TopInches, $topTwips)      -ForegroundColor Cyan
Write-Host ("  Bottom:           {0,5} in ({1} twips)" -f $BottomInches, $bottomTwips) -ForegroundColor Cyan
Write-Host ("  Header from top:  {0,5} in ({1} twips)" -f $HeaderInches, $headerTwips) -ForegroundColor Cyan
Write-Host ("  Footer from bot:  {0,5} in ({1} twips)" -f $FooterInches, $footerTwips) -ForegroundColor Cyan
Write-Host ""
Write-Host "Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "  $_" }
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Helper: set or insert an attribute on a self-closing tag
# ---------------------------------------------------------------------------
function Set-PgMarAttr {
    param([string]$Tag, [string]$AttrName, [int]$Value)
    if ($Tag -match "$AttrName=`"[^`"]*`"") {
        return [regex]::Replace($Tag, "$AttrName=`"[^`"]*`"", "$AttrName=`"$Value`"")
    } else {
        return [regex]::Replace($Tag, '/>$', " $AttrName=`"$Value`"/>")
    }
}

# ---------------------------------------------------------------------------
# Process each file
# ---------------------------------------------------------------------------

$pgMarEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $tag = $m.Value
    $tag = Set-PgMarAttr -Tag $tag -AttrName 'w:left'   -Value $leftTwips
    $tag = Set-PgMarAttr -Tag $tag -AttrName 'w:right'  -Value $rightTwips
    $tag = Set-PgMarAttr -Tag $tag -AttrName 'w:top'    -Value $topTwips
    $tag = Set-PgMarAttr -Tag $tag -AttrName 'w:bottom' -Value $bottomTwips
    $tag = Set-PgMarAttr -Tag $tag -AttrName 'w:header' -Value $headerTwips
    $tag = Set-PgMarAttr -Tag $tag -AttrName 'w:footer' -Value $footerTwips
    return $tag
}

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

    $documentXml = [System.Text.RegularExpressions.Regex]::Replace($documentXml, '<w:pgMar\b[^/]*/>', $pgMarEvaluator)
    $sectionCount = [regex]::Matches($documentXml, "<w:pgMar\b").Count
    Write-Host "  Updated pgMar in $sectionCount section(s)" -ForegroundColor DarkGray

    # Write document.xml back into the package
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "pglayout-$([guid]::NewGuid()).zip")
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
                            } finally { $es.Dispose() }
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
Write-Host "Done. Open in Word and verify under Layout > Margins > Custom Margins." -ForegroundColor Green
Write-Host "Note: mirror margins are enabled separately (e.g., via a header/footer content script), not this script." -ForegroundColor Cyan
