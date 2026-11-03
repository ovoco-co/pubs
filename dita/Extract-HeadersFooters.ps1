<#
.SYNOPSIS
    Extracts header and footer text from a Word document into JSON.

.DESCRIPTION
    Reads header*.xml and footer*.xml from a .docx package, extracts
    the text content of each, and writes a JSON file listing all headers
    and footers with their filenames and text.

.PARAMETER InputDocx
    Path to the source .docx file.

.PARAMETER OutputJson
    Path where the headers/footers JSON will be written.

.EXAMPLE
    .\Extract-HeadersFooters.ps1 -InputDocx "chapter.docx" -OutputJson "headers-footers.json"
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

function Get-AllTextFromXmlString {
    param([string]$XmlText)
    [xml]$doc = $XmlText
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $texts = @()
    foreach ($t in $doc.SelectNodes("//w:t", $nsMgr)) {
        if ($null -ne $t.InnerText) { $texts += $t.InnerText }
    }
    return ($texts -join " ").Trim()
}

$result = [ordered]@{
    headers = @()
    footers = @()
}

$stream = [System.IO.File]::Open($InputDocx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -match '^word/(header\d+\.xml)$') {
                $fileName = $Matches[1]
                $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
                try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
                $result.headers += [pscustomobject]@{
                    file = $fileName
                    text = Get-AllTextFromXmlString -XmlText $xmlText
                }
            }
            elseif ($entry.FullName -match '^word/(footer\d+\.xml)$') {
                $fileName = $Matches[1]
                $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
                try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
                $result.footers += [pscustomobject]@{
                    file = $fileName
                    text = Get-AllTextFromXmlString -XmlText $xmlText
                }
            }
        }
    } finally { $zip.Dispose() }
} finally { $stream.Dispose() }

$result | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputJson -Encoding UTF8
Write-Host "Headers/footers written to: $OutputJson"
