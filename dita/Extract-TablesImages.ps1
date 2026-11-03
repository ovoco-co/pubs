<#
.SYNOPSIS
    Extracts tables and images from a Word document.

.DESCRIPTION
    Reads word/document.xml from a .docx package, extracts table content
    as JSON (one array per table with rows and cells), and copies all
    images from word/media/ to an output directory.

.PARAMETER InputDocx
    Path to the source .docx file.

.PARAMETER OutputDir
    Directory where tables.json and images/ will be written.
    Created if it does not exist.

.EXAMPLE
    .\Extract-TablesImages.ps1 -InputDocx "chapter.docx" -OutputDir "output\assets"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-NodeText {
    param(
        [System.Xml.XmlNode]$Node,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )
    $texts = @()
    foreach ($t in $Node.SelectNodes(".//w:t", $NsMgr)) {
        if ($null -ne $t.InnerText) { $texts += $t.InnerText }
    }
    return ($texts -join " ").Trim()
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$imagesOut = Join-Path $OutputDir "images"
if (-not (Test-Path $imagesOut)) {
    New-Item -ItemType Directory -Path $imagesOut | Out-Null
}

$docXmlText = $null
$stream = [System.IO.File]::Open($InputDocx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        # Read document.xml
        $docEntry = $zip.GetEntry('word/document.xml')
        if ($docEntry) {
            $reader = New-Object System.IO.StreamReader($docEntry.Open(), [System.Text.Encoding]::UTF8)
            try { $docXmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }

        # Extract images from word/media/
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -match '^word/media/(.+)$') {
                $fileName = $Matches[1]
                $destPath = Join-Path $imagesOut $fileName
                $entryStream = $entry.Open()
                try {
                    $fileStream = [System.IO.File]::Create($destPath)
                    try { $entryStream.CopyTo($fileStream) } finally { $fileStream.Dispose() }
                } finally { $entryStream.Dispose() }
            }
        }
    } finally { $zip.Dispose() }
} finally { $stream.Dispose() }

$tables = @()
if ($docXmlText) {
    [xml]$docXml = [xml]$docXmlText
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
    $nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

    $tableIndex = 1
    foreach ($tbl in $docXml.SelectNodes("//w:tbl", $nsMgr)) {
        $rows = @()
        foreach ($tr in $tbl.SelectNodes("./w:tr", $nsMgr)) {
            $cells = @()
            foreach ($tc in $tr.SelectNodes("./w:tc", $nsMgr)) {
                $cells += (Get-NodeText -Node $tc -NsMgr $nsMgr)
            }
            $rows += ,$cells
        }

        $tables += [pscustomobject]@{
            tableIndex = $tableIndex
            rows = $rows
        }
        $tableIndex++
    }
}

$tablesPath = Join-Path $OutputDir "tables.json"
$tables | ConvertTo-Json -Depth 10 | Set-Content -Path $tablesPath -Encoding UTF8

Write-Host "Tables written to: $tablesPath"
Write-Host "Images extracted to: $imagesOut"
