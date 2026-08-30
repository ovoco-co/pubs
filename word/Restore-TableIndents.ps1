<#
.SYNOPSIS
    Restores the indentation pandoc drops from table cells, by reading the
    leading tabs out of the original Word file and writing real paragraph
    indents into the converted one.

.DESCRIPTION
    Statistical tables built by R flextable carry their row hierarchy as a
    leading tab in the label cell: the group name sits flush left, its members
    one tab in. Pandoc's docx reader discards those tabs, and its paragraph
    indents with them, so a converted table comes out visually flat and the
    grouping is lost.

    This walks the source and the converted file in parallel and, wherever a
    source cell paragraph starts with one or more tabs, writes an equivalent
    <w:ind w:left> on the matching converted paragraph. A real indent is better
    than the tab it replaces: it survives restyling, it wraps correctly on the
    second line, and Set-TableParagraphStyle.ps1 explicitly preserves it.

    Mapping is positional, table by table, row by row, cell by cell, paragraph
    by paragraph. Pandoc preserves table structure exactly even when it drops
    formatting, so the two line up. The script verifies that before writing and
    stops if the shapes differ at all, rather than indenting the wrong rows.

    Only leading tabs count. A tab used inside a line to separate values is
    left alone, since it is not expressing hierarchy.

    Idempotent: an existing w:ind on the target paragraph is overwritten with
    the same computed value, so re-running is a no-op.

    Backs up to <name>.indents.bak before writing.

.PARAMETER Source
    The original Word file, before conversion, that still has the tabs.

.PARAMETER Target
    The converted Word file to write indents into.

.PARAMETER Inches
    Indent per tab level. Defaults to 0.125", which reads at the 9pt of the
    NIST Table Text style about as a tab did at flextable's 7pt.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Restore-TableIndents.ps1 -Source "analysis.docx" -Target "C:\out\analysis.docx"

.EXAMPLE
    .\Restore-TableIndents.ps1 -Source "a.docx" -Target "out\a.docx" -Inches 0.2 -Force
#>

[CmdletBinding()]
param(
    [string]$Source,

    [Parameter(Mandatory=$true)]
    [string]$Target,

    [double]$Inches = 0.125,

    [int]$SpacesPerLevel = 1,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

if (-not (Test-Path -LiteralPath $Target -PathType Leaf)) { throw "Target not found: $Target" }
$Target = (Resolve-Path -LiteralPath $Target).ProviderPath
$selfMode = -not $Source
if (-not $selfMode) {
    if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) { throw "Source not found: $Source" }
    $Source = (Resolve-Path -LiteralPath $Source).ProviderPath
    if ([string]::Equals($Source, $Target, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Source and Target are the same file."
    }
}

$twipsPerLevel = [int][Math]::Round($Inches * 1440)

Write-Host "Restore Table Indents" -ForegroundColor Cyan
if ($selfMode) {
    Write-Host "  Source: none, reading the indent signal out of the target itself" -ForegroundColor Cyan
} else {
    Write-Host "  Source (has the tabs): $Source" -ForegroundColor Cyan
}
Write-Host "  Target (gets indents): $Target" -ForegroundColor Cyan
Write-Host "  Indent per tab level : $Inches`" ($twipsPerLevel twips)" -ForegroundColor Cyan
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

function Get-Xml {
    param([byte[]]$Bytes)
    # Word writes document.xml with a UTF-8 BOM. GetString keeps it as a leading
    # U+FEFF, and XmlDocument.LoadXml rejects that with "Data at the root level
    # is invalid" - the byte order mark is only legal in a stream, not a string.
    $text = $utf8NoBom.GetString($Bytes)
    return $text.TrimStart([char]0xFEFF)
}

function Read-Entries {
    param([string]$Path)
    $entries = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $stream  = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
    return ,$entries
}

function Get-TablePlan {
    param([System.Xml.XmlDocument]$Doc, [System.Xml.XmlNamespaceManager]$Ns)
    # Every cell paragraph, in document order, as a flat list. Depth does not
    # matter: a nested table's paragraphs appear in the same relative place in
    # both files, so the flat order still lines up.
    $plan = New-Object System.Collections.ArrayList
    foreach ($tbl in $Doc.SelectNodes('//w:tbl', $Ns)) {
        foreach ($tr in $tbl.SelectNodes('./w:tr', $Ns)) {
            foreach ($tc in $tr.SelectNodes('./w:tc', $Ns)) {
                foreach ($p in $tc.SelectNodes('./w:p', $Ns)) { [void]$plan.Add($p) }
            }
        }
    }
    return ,$plan
}

function Get-LeadingTabCount {
    param([System.Xml.XmlElement]$Para, [System.Xml.XmlNamespaceManager]$Ns)
    # Count <w:tab/> elements that appear before any run text with content.
    $count = 0
    foreach ($r in $Para.SelectNodes('./w:r', $Ns)) {
        foreach ($child in $r.ChildNodes) {
            if ($child.LocalName -eq 'tab')  { $count++; continue }
            if ($child.LocalName -eq 't') {
                if ($child.InnerText.Trim().Length -gt 0) { return $count }
                continue
            }
        }
    }
    return $count
}

# Leading spaces are the other way a converter carries hierarchy: flextable's
# HTML writes "  Yes" under an unindented group label, and pandoc keeps the
# space. Turn it into a real indent and drop the space, since the two together
# would double the offset.
function Get-LeadingSpaceLevel {
    param([System.Xml.XmlElement]$Para, [System.Xml.XmlNamespaceManager]$Ns, [int]$PerLevel)
    # The indent space is not always in the same run as the label. Converters
    # split runs on formatting changes, so " " and "No Response" can be two
    # runs. Walk w:t nodes in order and count until real text appears.
    $spaces = 0
    foreach ($t in $Para.SelectNodes('.//w:t', $Ns)) {
        $text = $t.InnerText
        $lead = $text.Length - $text.TrimStart(' ', [char]0x00A0).Length
        $spaces += $lead
        if ($lead -lt $text.Length) { break }   # this run has real text after the spaces
    }
    if ($spaces -lt $PerLevel) { return 0 }
    return [int][Math]::Floor($spaces / $PerLevel)
}

function Remove-LeadingSpaces {
    param([System.Xml.XmlElement]$Para, [System.Xml.XmlNamespaceManager]$Ns)
    foreach ($t in $Para.SelectNodes('.//w:t', $Ns)) {
        $text    = $t.InnerText
        $trimmed = $text.TrimStart(' ', [char]0x00A0)
        if ($trimmed -ne $text) {
            $t.InnerText = $trimmed
            # Word is strict about w:t: without xml:space, remaining meaningful
            # spacing can be dropped on the next round trip.
            [void]$t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        }
        if ($trimmed.Length -gt 0) { break }   # reached the label itself
    }
}

$tgtEntries = Read-Entries -Path $Target
if (-not $tgtEntries.ContainsKey('word/document.xml')) { throw "No word/document.xml in the Target file." }

if (-not $selfMode) {
    $srcEntries = Read-Entries -Path $Source
    if (-not $srcEntries.ContainsKey('word/document.xml')) { throw "No word/document.xml in the Source file." }
    $srcDoc = New-Object System.Xml.XmlDocument
    $srcDoc.PreserveWhitespace = $true
    $srcDoc.LoadXml((Get-Xml -Bytes $srcEntries['word/document.xml']))
    $srcNs = New-Object System.Xml.XmlNamespaceManager($srcDoc.NameTable)
    $srcNs.AddNamespace('w', $W)
}

$tgtDoc = New-Object System.Xml.XmlDocument
$tgtDoc.PreserveWhitespace = $true
$tgtDoc.LoadXml((Get-Xml -Bytes $tgtEntries['word/document.xml']))
$tgtNs = New-Object System.Xml.XmlNamespaceManager($tgtDoc.NameTable)
$tgtNs.AddNamespace('w', $W)

$tgtPlan = Get-TablePlan -Doc $tgtDoc -Ns $tgtNs
Write-Host "  Target table paragraphs: $($tgtPlan.Count)" -ForegroundColor Gray

if (-not $selfMode) {
    $srcPlan = Get-TablePlan -Doc $srcDoc -Ns $srcNs
    Write-Host "  Source table paragraphs: $($srcPlan.Count)" -ForegroundColor Gray
    if ($srcPlan.Count -ne $tgtPlan.Count) {
        throw "Table structure differs: $($srcPlan.Count) cell paragraphs in the source, $($tgtPlan.Count) in the target. Positional mapping is not safe. Check that -Target really was converted from -Source."
    }
}
if ($tgtPlan.Count -eq 0) { Write-Host "  No table paragraphs. Nothing to do." -ForegroundColor DarkGray; return }

$applied = 0
$levels  = @{}
$stripped = 0
for ($i = 0; $i -lt $tgtPlan.Count; $i++) {
    $para = $tgtPlan[$i]
    $strip = $false
    if ($selfMode) {
        $tabs = Get-LeadingTabCount -Para $para -Ns $tgtNs
        if ($tabs -le 0) {
            $tabs = Get-LeadingSpaceLevel -Para $para -Ns $tgtNs -PerLevel $SpacesPerLevel
            if ($tabs -gt 0) { $strip = $true }
        }
    } else {
        $tabs = Get-LeadingTabCount -Para $srcPlan[$i] -Ns $srcNs
    }
    if ($tabs -le 0) { continue }
    if ($strip) { Remove-LeadingSpaces -Para $para -Ns $tgtNs; $stripped++ }
    $pPr  = $para.SelectSingleNode('./w:pPr', $tgtNs)
    if (-not $pPr) {
        # pPr has to be the first child of w:p.
        $pPr = $tgtDoc.CreateElement('w', 'pPr', $W)
        [void]$para.PrependChild($pPr)
    }
    $ind = $pPr.SelectSingleNode('./w:ind', $tgtNs)
    if (-not $ind) {
        $ind = $tgtDoc.CreateElement('w', 'ind', $W)
        # w:ind sits after pStyle and numPr in the schema's ordering, and
        # before spacing/jc. Appending is valid for the elements in play here.
        [void]$pPr.AppendChild($ind)
    }
    [void]$ind.SetAttribute('left', $W, ($tabs * $twipsPerLevel).ToString())
    $applied++
    $levels[$tabs] = 1 + $(if ($levels.ContainsKey($tabs)) { $levels[$tabs] } else { 0 })
}

Write-Host "  Paragraphs indented: $applied" -ForegroundColor Gray
if ($stripped -gt 0) { Write-Host "  Leading spaces removed (now carried by the indent): $stripped" -ForegroundColor Gray }
foreach ($k in ($levels.Keys | Sort-Object)) {
    Write-Host ("    $k tab(s) -> {0}`" : {1} paragraph(s)" -f ($k * $Inches), $levels[$k]) -ForegroundColor DarkGray
}

if ($applied -eq 0) {
    Write-Host "  No indent signal found. File unchanged." -ForegroundColor DarkGray
    return
}

if (-not $Force -and (Test-Path -LiteralPath "$Target.indents.bak")) {
    Write-Warning "  Backup $Target.indents.bak already exists; leaving it as-is."
} else {
    Copy-Item -LiteralPath $Target -Destination "$Target.indents.bak" -Force
}

$tgtEntries['word/document.xml'] = $utf8NoBom.GetBytes($tgtDoc.OuterXml)

$tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "indents-$([guid]::NewGuid()).zip")
$outStream = [System.IO.File]::Create($tempZip)
try {
    $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $tgtEntries.Keys) {
            $ne = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $os = $ne.Open()
            try { $os.Write($tgtEntries[$name], 0, $tgtEntries[$name].Length) } finally { $os.Dispose() }
        }
    } finally { $outZip.Dispose() }
} finally { $outStream.Dispose() }

Move-Item -LiteralPath $tempZip -Destination $Target -Force
Write-Host "  Saved." -ForegroundColor Green
Write-Host ""
Write-Host "Done." -ForegroundColor Green
