<#
.SYNOPSIS
    Audits a Word file's theme colors and reports every style or run that
    bypasses the theme by hardcoding an explicit hex color.

.DESCRIPTION
    The classic reason a Word template's theme colors "do not stay" is
    that styles or run-level overrides set an explicit RGB value
    (`<w:color w:val="FF6B35"/>`) instead of referencing a theme color
    (`<w:color w:themeColor="accent1"/>`). The explicit value always wins
    over the theme. This script finds every such case so you can decide
    which ones to convert to theme references and bake your colors in
    permanently.

    Reports:
      - The theme palette currently set in word/theme/theme1.xml
      - Every style that sets an explicit color (not a theme reference)
      - Every run override in document.xml, headers, footers, and the
        glossary that sets an explicit color

    Output is a CSV file plus a console summary.

.PARAMETER File
    Path to the .docx, .dotx, .docm, or .dotm file to audit.

.PARAMETER OutputDir
    Where to write the audit CSV. Defaults to a sibling _theme-audit folder.

.EXAMPLE
    .\Audit-WordTheme.ps1 -File "C:\templates\master.dotx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $File -PathType Leaf)) {
    throw "File not found: $File"
}

$resolvedPath = (Resolve-Path -LiteralPath $File).ProviderPath

# Validate by string match rather than FileInfo.Extension; the latter has
# shown empty results on certain UNC paths.
if ($resolvedPath -notmatch '\.(docx|dotx|docm|dotm)$') {
    throw "Unsupported extension. Path resolved to: $resolvedPath  (length $($resolvedPath.Length)). Expected .docx, .dotx, .docm, or .dotm."
}

# Compute path components via Path static methods (avoid FileInfo, which
# has shown null DirectoryName/BaseName on some UNC configurations).
$dirName  = [System.IO.Path]::GetDirectoryName($resolvedPath)
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
$fileName = [System.IO.Path]::GetFileName($resolvedPath)

if (-not $OutputDir) {
    $OutputDir = Join-Path $dirName '_theme-audit'
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$aNs = 'http://schemas.openxmlformats.org/drawingml/2006/main'

function Get-PackagedXml {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName); if (-not $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        $text = $reader.ReadToEnd()
        $reader.Dispose()
    } finally { $stream.Dispose() }
    $xml = New-Object System.Xml.XmlDocument
    $xml.LoadXml($text)
    return $xml
}

function New-NsManager {
    param([System.Xml.XmlDocument]$Xml, [hashtable]$Prefixes)
    $nsm = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    foreach ($k in $Prefixes.Keys) { $nsm.AddNamespace($k, $Prefixes[$k]) }
    # Comma-wrap so PowerShell does not enumerate the manager (IEnumerable)
    # and hand callers an array of namespace prefixes.
    return ,$nsm
}

function Get-ThemePalette {
    param([System.IO.Compression.ZipArchive]$Zip)
    $xml = Get-PackagedXml -Zip $Zip -EntryName 'word/theme/theme1.xml'
    if (-not $xml) { return $null }
    $nsm = New-NsManager -Xml $xml -Prefixes @{ a = $aNs }

    $palette = [ordered]@{}
    $clrScheme = $xml.SelectSingleNode('//a:clrScheme', $nsm)
    if (-not $clrScheme) { return $palette }

    foreach ($child in $clrScheme.ChildNodes) {
        $name = $child.LocalName
        # Each color slot has either a:srgbClr (hex) or a:sysClr (system color, e.g., windowText)
        $srgb = $child.SelectSingleNode('a:srgbClr', $nsm)
        $sys  = $child.SelectSingleNode('a:sysClr', $nsm)
        if ($srgb) {
            $palette[$name] = @{ Type = 'srgb'; Value = $srgb.GetAttribute('val') }
        } elseif ($sys) {
            $palette[$name] = @{ Type = 'sys'; Value = $sys.GetAttribute('val'); LastClr = $sys.GetAttribute('lastClr') }
        }
    }
    return $palette
}

function Find-ExplicitColors {
    param([System.IO.Compression.ZipArchive]$Zip)
    $hits = @()

    $xmlEntries = $Zip.Entries | Where-Object {
        $_.FullName -eq 'word/styles.xml' -or
        $_.FullName -eq 'word/document.xml' -or
        $_.FullName -match '^word/(header|footer)\d+\.xml$' -or
        $_.FullName -eq 'word/glossary/document.xml' -or
        $_.FullName -eq 'word/numbering.xml'
    }

    foreach ($e in $xmlEntries) {
        $xml = Get-PackagedXml -Zip $Zip -EntryName $e.FullName
        if (-not $xml) { continue }
        $nsm = New-NsManager -Xml $xml -Prefixes @{ w = $wNs }

        # Find every w:color element that has w:val and no w:themeColor
        foreach ($colorNode in $xml.SelectNodes('//w:color', $nsm)) {
            $val = $colorNode.GetAttribute('val', $wNs)
            $theme = $colorNode.GetAttribute('themeColor', $wNs)
            if ($val -and -not $theme -and $val -ne 'auto') {
                # Walk up to find context (style name, paragraph, run)
                $context = ''
                $node = $colorNode.ParentNode
                while ($node -and $context -eq '') {
                    if ($node.LocalName -eq 'style') {
                        $styleId = $node.GetAttribute('styleId', $wNs)
                        $nameNode = $node.SelectSingleNode('w:name', $nsm)
                        $styleName = if ($nameNode) { $nameNode.GetAttribute('val', $wNs) } else { $styleId }
                        $type = $node.GetAttribute('type', $wNs)
                        $context = "style[$type]::$styleName"
                    } elseif ($node.LocalName -eq 'tblStylePr') {
                        # Look up parent style
                        $parent = $node.ParentNode
                        $styleId = $parent.GetAttribute('styleId', $wNs)
                        $cType = $node.GetAttribute('type', $wNs)
                        $context = "tableConditional[$styleId.$cType]"
                    }
                    $node = $node.ParentNode
                }
                if (-not $context) { $context = "inline-run" }

                # Find what kind of color this is (rPr, rPrChange, etc.)
                $kind = $colorNode.ParentNode.LocalName

                $hits += [pscustomobject]@{
                    Part        = $e.FullName
                    Context     = $context
                    Container   = $kind
                    ExplicitVal = $val
                    HasThemeAttr = $false
                }
            }
        }
    }

    return $hits
}

# ---------------------------------------------------------------------------
# Run audit
# ---------------------------------------------------------------------------

Write-Host "Auditing $fileName..." -ForegroundColor Cyan

$stream = [System.IO.File]::Open($resolvedPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $palette = Get-ThemePalette -Zip $zip
        $hits = Find-ExplicitColors -Zip $zip
    } finally {
        $zip.Dispose()
    }
} finally {
    $stream.Dispose()
}

# Print theme palette
Write-Host ""
Write-Host "Theme palette (word/theme/theme1.xml):" -ForegroundColor Yellow
if (-not $palette -or $palette.Count -eq 0) {
    Write-Host "  No theme palette found."
} else {
    foreach ($name in $palette.Keys) {
        $entry = $palette[$name]
        if ($entry.Type -eq 'srgb') {
            Write-Host ("  {0,-12} srgb #{1}" -f $name, $entry.Value)
        } else {
            Write-Host ("  {0,-12} sys  {1}  (last: #{2})" -f $name, $entry.Value, $entry.LastClr)
        }
    }
}

# Print explicit color hits
Write-Host ""
Write-Host "Explicit colors that bypass the theme:" -ForegroundColor Yellow
if (-not $hits -or $hits.Count -eq 0) {
    Write-Host "  None. Every color reference is themed (or 'auto')." -ForegroundColor Green
} else {
    Write-Host "  Found $($hits.Count) explicit color reference(s)."
    Write-Host "  Top contexts:"
    $grouped = $hits | Group-Object Context | Sort-Object Count -Descending | Select-Object -First 15
    foreach ($g in $grouped) {
        Write-Host ("  {0,4}x  {1}" -f $g.Count, $g.Name)
    }
}

# Write CSV
$csvPath = Join-Path $OutputDir ($baseName + '.theme-audit.csv')
$hits | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "Theme audit CSV: $csvPath" -ForegroundColor Green

# Write palette CSV
$paletteCsvPath = Join-Path $OutputDir ($baseName + '.theme-palette.csv')
$paletteRows = @()
foreach ($name in $palette.Keys) {
    $entry = $palette[$name]
    $paletteRows += [pscustomobject]@{
        Slot     = $name
        Type     = $entry.Type
        Value    = $entry.Value
        LastClr  = if ($entry.LastClr) { $entry.LastClr } else { '' }
    }
}
$paletteRows | Export-Csv -LiteralPath $paletteCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "Theme palette CSV: $paletteCsvPath" -ForegroundColor Green
