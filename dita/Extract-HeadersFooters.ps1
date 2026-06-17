param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-AllTextFromXml {
    param([xml]$XmlDoc)

    $nsMgr = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $texts = @()

    foreach ($t in $XmlDoc.SelectNodes("//w:t", $nsMgr)) {
        if ($null -ne $t.InnerText) {
            $texts += $t.InnerText
        }
    }

    return ($texts -join " ").Trim()
}

$tempRoot = Join-Path $env:TEMP ("docx_hf_" + [Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    $tempZip = Join-Path $tempRoot "input.zip"
    Copy-Item -Path $InputDocx -Destination $tempZip -Force
    Expand-Archive -Path $tempZip -DestinationPath $tempRoot -Force

    $headerFiles = Get-ChildItem -Path (Join-Path $tempRoot "word") -Recurse -Filter "header*.xml" -ErrorAction SilentlyContinue
    $footerFiles = Get-ChildItem -Path (Join-Path $tempRoot "word") -Recurse -Filter "footer*.xml" -ErrorAction SilentlyContinue

    $result = [ordered]@{
        headers = @()
        footers = @()
    }

    foreach ($hf in $headerFiles) {
        [xml]$x = Get-Content -Path $hf.FullName -Raw -Encoding UTF8
        $result.headers += [pscustomobject]@{
            file = $hf.Name
            text = Get-AllTextFromXml -XmlDoc $x
        }
    }

    foreach ($ff in $footerFiles) {
        [xml]$x = Get-Content -Path $ff.FullName -Raw -Encoding UTF8
        $result.footers += [pscustomobject]@{
            file = $ff.Name
            text = Get-AllTextFromXml -XmlDoc $x
        }
    }

    $result | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputJson -Encoding UTF8
    Write-Host "Headers/footers written to: $OutputJson"
}
finally {
    if (Test-Path $tempRoot) {
        Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}