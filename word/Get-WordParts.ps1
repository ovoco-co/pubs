<#
.SYNOPSIS
    Lists every Building Block (Quick Part, AutoText, cover page, equation,
    header, footer, etc.) in a Word template's glossary, as a Markdown
    document.

.DESCRIPTION
    Reads word/glossary/document.xml directly from the package and emits
    one Markdown section per <w:docPart>. Each section shows the part's
    name, gallery (Quick Parts, AutoText, etc.), category, type,
    description, behaviors, and a content preview.

    Plain .docx files (not templates) typically have no glossary part;
    in that case the script reports "no building blocks found" and exits.

.PARAMETER File
    Path to the .docx, .dotx, .docm, or .dotm file. Required.

.PARAMETER OutputDir
    Where to write the output. Defaults to a sibling "_parts" folder.

.PARAMETER Html
    Also produce an HTML version for browser viewing.

.PARAMETER FilterGallery
    Restrict output to a single gallery (case-insensitive substring match):
    "QuickParts", "AutoText", "Headers", "Footers", "CoverPage", etc.
    Default is all galleries.

.EXAMPLE
    .\Get-WordParts.ps1 -File "master.dotx"

.EXAMPLE
    .\Get-WordParts.ps1 -File "master.dotx" -FilterGallery AutoText
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [switch]$Html,

    [Parameter(Mandatory = $false)]
    [string]$FilterGallery
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
$resolvedPath = (Resolve-Path -LiteralPath $File).ProviderPath
if ($resolvedPath -notmatch '\.(docx|dotx|docm|dotm)$') {
    throw "Unsupported extension. Path resolved to: $resolvedPath. Expected .docx, .dotx, .docm, or .dotm."
}
$dirName  = [System.IO.Path]::GetDirectoryName($resolvedPath)
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
$fileName = [System.IO.Path]::GetFileName($resolvedPath)

if (-not $OutputDir) { $OutputDir = Join-Path $dirName '_parts' }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function W-Val { param([System.Xml.XmlElement]$N)
    if (-not $N) { return $null }
    $v = $N.GetAttribute('val', $wNs); if ($v) { return $v }
    return $N.GetAttribute('val')
}

function Get-XmlText {
    param([System.Xml.XmlNode]$Node, [System.Xml.XmlNamespaceManager]$Nsm)
    if (-not $Node) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $Node.SelectNodes('.//w:t', $Nsm)) { [void]$sb.Append($t.InnerText) }
    foreach ($br in $Node.SelectNodes('.//w:br', $Nsm)) { [void]$sb.Append(' ') }
    foreach ($tab in $Node.SelectNodes('.//w:tab', $Nsm)) { [void]$sb.Append("`t") }
    return $sb.ToString()
}

function Trunc {
    param([string]$Text, [int]$Max = 400)
    if ($null -eq $Text) { return '' }
    $clean = $Text -replace "`r`n", ' ' -replace "`r", ' ' -replace "`n", ' '
    $clean = $clean.Trim()
    if ($clean.Length -le $Max) { return $clean }
    return $clean.Substring(0, $Max) + '...'
}

# ---------------------------------------------------------------------------
# Open package, read glossary
# ---------------------------------------------------------------------------

Write-Host "Reading building blocks from $fileName..." -ForegroundColor Cyan

$glossaryXmlText = $null
$stream = [System.IO.File]::Open($resolvedPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $entry = $zip.GetEntry('word/glossary/document.xml')
        if (-not $entry) {
            Write-Host "No word/glossary/document.xml in this file. Plain .docx files typically have no building blocks; only templates do." -ForegroundColor Yellow
            return
        }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { $glossaryXmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
} finally { $stream.Dispose() }

if (-not $glossaryXmlText) { return }

$xml = New-Object System.Xml.XmlDocument
$xml.LoadXml($glossaryXmlText)
$nsm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsm.AddNamespace('w', $wNs)

# ---------------------------------------------------------------------------
# Extract every <w:docPart>
# ---------------------------------------------------------------------------

$parts = @()
foreach ($dp in $xml.SelectNodes('//w:docPart', $nsm)) {
    $pr = $dp.SelectSingleNode('w:docPartPr', $nsm)
    if (-not $pr) { continue }

    $name = $pr.SelectSingleNode('w:name', $nsm)
    $catName = $pr.SelectSingleNode('w:category/w:name', $nsm)
    $catGallery = $pr.SelectSingleNode('w:category/w:gallery', $nsm)
    $type = $pr.SelectSingleNode('w:types/w:type', $nsm)
    $description = $pr.SelectSingleNode('w:description', $nsm)
    $behaviors = @()
    foreach ($b in $pr.SelectNodes('w:behaviors/w:behavior', $nsm)) {
        $v = W-Val $b; if ($v) { $behaviors += $v }
    }
    $guid = $pr.SelectSingleNode('w:guid', $nsm)
    $contentText = Get-XmlText -Node ($dp.SelectSingleNode('w:docPartBody', $nsm)) -Nsm $nsm

    $parts += [pscustomobject]@{
        Name        = if ($name) { W-Val $name } else { '(unnamed)' }
        Gallery     = if ($catGallery) { W-Val $catGallery } else { $null }
        Category    = if ($catName) { W-Val $catName } else { $null }
        Type        = if ($type) { W-Val $type } else { $null }
        Description = if ($description) { W-Val $description } else { $null }
        Behaviors   = $behaviors
        Guid        = if ($guid) { W-Val $guid } else { $null }
        Content     = $contentText
    }
}

if ($FilterGallery) {
    $parts = $parts | Where-Object { $_.Gallery -and $_.Gallery -like "*$FilterGallery*" }
}

if ($parts.Count -eq 0) {
    Write-Host "No building blocks found." -ForegroundColor Yellow
    return
}

$parts = $parts | Sort-Object Gallery, Category, Name

# ---------------------------------------------------------------------------
# Render Markdown
# ---------------------------------------------------------------------------

$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine("# Word parts (building blocks) in $fileName")
[void]$md.AppendLine('')
[void]$md.AppendLine("Generated $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))  ")
[void]$md.AppendLine("$($parts.Count) part(s)")
[void]$md.AppendLine('')

# Summary by gallery
[void]$md.AppendLine('## Summary')
[void]$md.AppendLine('')
$grouped = $parts | Group-Object Gallery | Sort-Object Name
foreach ($g in $grouped) {
    $galleryName = if ($g.Name) { $g.Name } else { '(no gallery)' }
    [void]$md.AppendLine("- $galleryName`: $($g.Count)")
}
[void]$md.AppendLine('')

# Detail
$lastGallery = $null
foreach ($p in $parts) {
    $g = if ($p.Gallery) { $p.Gallery } else { '(no gallery)' }
    if ($g -ne $lastGallery) {
        [void]$md.AppendLine('')
        [void]$md.AppendLine("## $g")
        [void]$md.AppendLine('')
        $lastGallery = $g
    }

    [void]$md.AppendLine("### $($p.Name)")
    [void]$md.AppendLine('')
    [void]$md.AppendLine('```')
    foreach ($k in @('Gallery','Category','Type','Description','Guid')) {
        $v = $p.$k
        if ($null -eq $v -or $v -eq '') { $v = '(not set)' }
        [void]$md.AppendLine(("{0,-14} {1}" -f ($k + ':'), $v))
    }
    $bv = if ($p.Behaviors.Count -gt 0) { $p.Behaviors -join ', ' } else { '(none)' }
    [void]$md.AppendLine(("{0,-14} {1}" -f 'Behaviors:', $bv))
    [void]$md.AppendLine('```')
    [void]$md.AppendLine('')

    if ($p.Content) {
        [void]$md.AppendLine('**Content preview**')
        [void]$md.AppendLine('')
        [void]$md.AppendLine('```')
        [void]$md.AppendLine((Trunc $p.Content 600))
        [void]$md.AppendLine('```')
        [void]$md.AppendLine('')
    }
}

$mdPath = Join-Path $OutputDir ($baseName + '.parts.md')
$md.ToString() | Set-Content -LiteralPath $mdPath -Encoding UTF8
Write-Host ""
Write-Host "Wrote $($parts.Count) building blocks to: $mdPath" -ForegroundColor Green

# Console summary
Write-Host ""
Write-Host "By gallery:" -ForegroundColor Cyan
foreach ($g in $grouped) {
    $galleryName = if ($g.Name) { $g.Name } else { '(no gallery)' }
    Write-Host ("  {0,-15} {1}" -f $galleryName, $g.Count)
}
