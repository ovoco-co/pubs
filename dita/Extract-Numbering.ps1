<#
.SYNOPSIS
    Extracts numbering definitions from a Word document into JSON.

.DESCRIPTION
    Reads word/numbering.xml from a .docx package, extracts abstractNum
    definitions and num-to-abstractNum mappings, and writes them as a
    JSON file for downstream processing.

.PARAMETER InputDocx
    Path to the source .docx file.

.PARAMETER OutputJson
    Path where the numbering JSON will be written.

.EXAMPLE
    .\Extract-Numbering.ps1 -InputDocx "chapter.docx" -OutputJson "numbering.json"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$numXmlText = $null
$stream = [System.IO.File]::Open($InputDocx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $entry = $zip.GetEntry('word/numbering.xml')
        if (-not $entry) {
            Write-Host "No numbering.xml found. Writing empty JSON."
            @{} | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputJson -Encoding UTF8
            return
        }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        try { $numXmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
} finally { $stream.Dispose() }

[xml]$numXml = [xml]$numXmlText
$nsMgr = New-Object System.Xml.XmlNamespaceManager($numXml.NameTable)
$nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

$abstractNums = @{}
foreach ($abs in $numXml.SelectNodes("//w:abstractNum", $nsMgr)) {
    $absId = $abs.GetAttribute("abstractNumId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $levels = @()

    foreach ($lvl in $abs.SelectNodes("./w:lvl", $nsMgr)) {
        $ilvl = $lvl.GetAttribute("ilvl", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        $fmtNode = $lvl.SelectSingleNode("./w:numFmt", $nsMgr)
        $txtNode = $lvl.SelectSingleNode("./w:lvlText", $nsMgr)

        $levels += [pscustomobject]@{
            ilvl = $ilvl
            format = if ($fmtNode) { $fmtNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main") } else { $null }
            text = if ($txtNode) { $txtNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main") } else { $null }
        }
    }

    $abstractNums[$absId] = $levels
}

$nums = @()
foreach ($num in $numXml.SelectNodes("//w:num", $nsMgr)) {
    $numId = $num.GetAttribute("numId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $absNode = $num.SelectSingleNode("./w:abstractNumId", $nsMgr)
    $absId = if ($absNode) { $absNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main") } else { $null }

    $nums += [pscustomobject]@{
        numId = $numId
        abstractNumId = $absId
        levels = if ($absId -and $abstractNums.ContainsKey($absId)) { $abstractNums[$absId] } else { @() }
    }
}

$nums | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputJson -Encoding UTF8
Write-Host "Numbering metadata written to: $OutputJson"