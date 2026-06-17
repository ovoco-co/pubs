param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$tempRoot = Join-Path $env:TEMP ("docx_numbering_" + [Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    $tempZip = Join-Path $tempRoot "input.zip"
    Copy-Item -Path $InputDocx -Destination $tempZip -Force
    Expand-Archive -Path $tempZip -DestinationPath $tempRoot -Force

    $numberingPath = Join-Path $tempRoot "word\numbering.xml"
    if (-not (Test-Path $numberingPath)) {
        Write-Host "No numbering.xml found. Writing empty JSON."
        @{} | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputJson -Encoding UTF8
        return
    }

    [xml]$numXml = Get-Content -Path $numberingPath -Raw -Encoding UTF8
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
}
finally {
    if (Test-Path $tempRoot) {
        Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}