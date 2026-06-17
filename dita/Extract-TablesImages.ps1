param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-NodeText {
    param(
        [System.Xml.XmlNode]$Node,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $texts = @()
    foreach ($t in $Node.SelectNodes(".//w:t", $NsMgr)) {
        if ($null -ne $t.InnerText) {
            $texts += $t.InnerText
        }
    }
    return ($texts -join " ").Trim()
}

$tempRoot = Join-Path $env:TEMP ("docx_tables_" + [Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    $tempZip = Join-Path $tempRoot "input.zip"
    Copy-Item -Path $InputDocx -Destination $tempZip -Force
    Expand-Archive -Path $tempZip -DestinationPath $tempRoot -Force

    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir | Out-Null
    }

    $documentPath = Join-Path $tempRoot "word\document.xml"
    [xml]$docXml = Get-Content -Path $documentPath -Raw -Encoding UTF8
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
    $nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

    $tables = @()
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

    $tablesPath = Join-Path $OutputDir "tables.json"
    $tables | ConvertTo-Json -Depth 10 | Set-Content -Path $tablesPath -Encoding UTF8

    $mediaDir = Join-Path $tempRoot "word\media"
    $imagesOut = Join-Path $OutputDir "images"
    if (-not (Test-Path $imagesOut)) {
        New-Item -ItemType Directory -Path $imagesOut | Out-Null
    }

    if (Test-Path $mediaDir) {
        Get-ChildItem -Path $mediaDir -File | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination (Join-Path $imagesOut $_.Name) -Force
        }
    }

    Write-Host "Tables written to: $tablesPath"
    Write-Host "Images extracted to: $imagesOut"
}
finally {
    if (Test-Path $tempRoot) {
        Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}