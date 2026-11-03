<#
.SYNOPSIS
    Converts a Word document into DITA topic files and a ditamap.

.DESCRIPTION
    Reads word/document.xml from a .docx package, maps Word paragraph styles
    to DITA elements using a JSON style map, splits on Heading 1 boundaries
    into separate topic files, and generates a ditamap referencing all topics.

    Supports paragraphs, lists (bullet and numbered), notes, codeblocks,
    captions, inline formatting (bold, italic, subscript, superscript,
    hyperlinks), images, and tables. Nested headings become DITA sections.

.PARAMETER InputDocx
    Path to the source .docx file.

.PARAMETER OutputDir
    Directory where .dita topic files and the ditamap will be written.
    Created if it does not exist.

.PARAMETER StyleMapJson
    Path to a JSON file mapping Word style names to DITA element types.

.EXAMPLE
    .\Convert-WordToDita.ps1 -InputDocx "chapter.docx" -OutputDir "output\dita" -StyleMapJson "stylemap.json"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [Parameter(Mandatory = $true)]
    [string]$StyleMapJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Escape-XmlText {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Get-RunDitaContent {
    param(
        [System.Xml.XmlNode]$Run,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $textNodes = $Run.SelectNodes("./w:t", $NsMgr)
    $rawText = ($textNodes | ForEach-Object { $_.InnerText }) -join ""
    if ([string]::IsNullOrEmpty($rawText)) { return "" }
    $escaped = Escape-XmlText $rawText

    $rPr = $Run.SelectSingleNode("./w:rPr", $NsMgr)
    if ($null -eq $rPr) { return $escaped }

    $bold = $rPr.SelectSingleNode("./w:b", $NsMgr)
    $italic = $rPr.SelectSingleNode("./w:i", $NsMgr)
    $vertAlign = $rPr.SelectSingleNode("./w:vertAlign", $NsMgr)

    $result = $escaped
    if ($null -ne $vertAlign) {
        $val = $vertAlign.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        if ($val -eq "superscript") { $result = "<sup>$result</sup>" }
        elseif ($val -eq "subscript") { $result = "<sub>$result</sub>" }
    }
    if ($null -ne $italic) { $result = "<i>$result</i>" }
    if ($null -ne $bold) { $result = "<b>$result</b>" }

    return $result
}

function Get-ParagraphDitaContent {
    param(
        [System.Xml.XmlNode]$Paragraph,
        [System.Xml.XmlNamespaceManager]$NsMgr,
        [hashtable]$RelMap
    )

    $parts = @()

    foreach ($child in $Paragraph.ChildNodes) {
        if ($child.LocalName -eq "r") {
            $content = Get-RunDitaContent -Run $child -NsMgr $NsMgr
            if ($content) { $parts += $content }
        }
        elseif ($child.LocalName -eq "hyperlink") {
            $rId = $child.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            $linkRuns = $child.SelectNodes("./w:r", $NsMgr)
            $linkText = ($linkRuns | ForEach-Object { Get-RunDitaContent -Run $_ -NsMgr $NsMgr }) -join ""
            if ($rId -and $RelMap.ContainsKey($rId)) {
                $href = Escape-XmlText $RelMap[$rId]
                $parts += "<xref href=""$href"" scope=""external"" format=""html"">$linkText</xref>"
            } elseif ($linkText) {
                $parts += $linkText
            }
        }
    }

    return ($parts -join "").Trim()
}

function Get-ParagraphPlainText {
    param(
        [System.Xml.XmlNode]$Paragraph,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $runs = $Paragraph.SelectNodes(".//w:t", $NsMgr)
    $texts = @()
    foreach ($run in $runs) {
        if ($null -ne $run.InnerText) { $texts += $run.InnerText }
    }
    return ($texts -join "").Trim()
}

function Get-ParagraphStyle {
    param(
        [System.Xml.XmlNode]$Paragraph,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $styleNode = $Paragraph.SelectSingleNode("./w:pPr/w:pStyle", $NsMgr)
    if ($null -ne $styleNode) {
        return $styleNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    }
    return "Normal"
}

function Get-HeadingLevel {
    param([string]$StyleName)

    switch -Regex ($StyleName) {
        '^H1$' { return 1 }
        '^H2$' { return 2 }
        '^H3$' { return 3 }
        '^H4$' { return 4 }
        '^H5$' { return 5 }
        '^Heading1$' { return 1 }
        '^Heading2$' { return 2 }
        '^Heading3$' { return 3 }
        '^Heading4$' { return 4 }
        '^Heading5$' { return 5 }
        '^Heading 1$' { return 1 }
        '^Heading 2$' { return 2 }
        '^Heading 3$' { return 3 }
        '^Heading 4$' { return 4 }
        '^Heading 5$' { return 5 }
        default { return 0 }
    }
}

function Get-ParagraphImages {
    param(
        [System.Xml.XmlNode]$Paragraph,
        [System.Xml.XmlNamespaceManager]$NsMgr,
        [hashtable]$RelMap
    )

    $images = @()
    $blips = $Paragraph.SelectNodes(".//a:blip", $NsMgr)
    foreach ($blip in $blips) {
        $embed = $blip.GetAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        if ($embed -and $RelMap.ContainsKey($embed)) {
            $target = $RelMap[$embed]
            $fileName = [System.IO.Path]::GetFileName($target)
            $images += $fileName
        }
    }
    return $images
}

function New-SafeId {
    param([string]$Text)
    $id = $Text -replace '[^a-zA-Z0-9_-]', '_' -replace '_+', '_' -replace '^_|_$', ''
    if ([string]::IsNullOrWhiteSpace($id)) { $id = "id" }
    if ($id -match '^\d') { $id = "t_$id" }
    return $id.ToLower()
}

function New-SafeFileName {
    param([string]$Text)
    $safe = $Text -replace '[\\/:*?"<>|]', '_' -replace '\s+', '_' -replace '_+', '_'
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "topic" }
    return $safe.Trim('_')
}

function Write-IndentedLine {
    param(
        [System.Text.StringBuilder]$Builder,
        [int]$Indent,
        [string]$Text
    )
    [void]$Builder.AppendLine(("{0}{1}" -f ("  " * $Indent), $Text))
}

function Convert-TableToDita {
    param(
        [System.Xml.XmlNode]$Table,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $sb = New-Object System.Text.StringBuilder
    $rows = $Table.SelectNodes("./w:tr", $NsMgr)
    if ($rows.Count -eq 0) { return "" }

    $firstRow = $rows[0]
    $colCount = $firstRow.SelectNodes("./w:tc", $NsMgr).Count

    Write-IndentedLine -Builder $sb -Indent 1 -Text "<simpletable>"

    $isFirst = $true
    foreach ($tr in $rows) {
        $tag = if ($isFirst) { "sthead" } else { "strow" }
        Write-IndentedLine -Builder $sb -Indent 2 -Text "<$tag>"

        foreach ($tc in $tr.SelectNodes("./w:tc", $NsMgr)) {
            $cellTexts = @()
            foreach ($cp in $tc.SelectNodes("./w:p", $NsMgr)) {
                $t = ($cp.SelectNodes(".//w:t", $NsMgr) | ForEach-Object { $_.InnerText }) -join ""
                if ($t) { $cellTexts += (Escape-XmlText $t) }
            }
            $cellContent = $cellTexts -join " "
            Write-IndentedLine -Builder $sb -Indent 3 -Text "<stentry>$cellContent</stentry>"
        }

        Write-IndentedLine -Builder $sb -Indent 2 -Text "</$tag>"
        $isFirst = $false
    }

    Write-IndentedLine -Builder $sb -Indent 1 -Text "</simpletable>"
    return $sb.ToString()
}

function Convert-ParagraphsToTopicXml {
    param(
        [Parameter(Mandatory = $true)]
        $ParagraphItems,
        [string]$TopicId = "topic"
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    [void]$sb.AppendLine('<!DOCTYPE topic PUBLIC "-//OASIS//DTD DITA Topic//EN" "topic.dtd">')
    [void]$sb.AppendLine("<topic id=""$(Escape-XmlText $TopicId)"">")

    $topicTitleWritten = $false
    $openSections = New-Object System.Collections.Generic.Stack[int]
    $listOpen = $false
    $listType = $null
    $sectionCounter = 0

    function Close-List {
        param(
            [System.Text.StringBuilder]$Builder,
            [ref]$IsOpen,
            [ref]$Type
        )
        if ($IsOpen.Value) {
            if ($Type.Value -eq "ul") {
                Write-IndentedLine -Builder $Builder -Indent 1 -Text "</ul>"
            }
            elseif ($Type.Value -eq "ol") {
                Write-IndentedLine -Builder $Builder -Indent 1 -Text "</ol>"
            }
            $IsOpen.Value = $false
            $Type.Value = $null
        }
    }

    function Close-SectionsToLevel {
        param(
            [System.Text.StringBuilder]$Builder,
            [System.Collections.Generic.Stack[int]]$Stack,
            [int]$TargetLevel
        )
        while ($Stack.Count -gt 0 -and $Stack.Peek() -ge $TargetLevel) {
            $indent = $Stack.Count
            Write-IndentedLine -Builder $Builder -Indent $indent -Text "</section>"
            [void]$Stack.Pop()
        }
    }

    function Open-Section {
        param(
            [System.Text.StringBuilder]$Builder,
            [System.Collections.Generic.Stack[int]]$Stack,
            [int]$Level,
            [string]$TitleText,
            [string]$SectionId
        )
        $indent = $Level
        Write-IndentedLine -Builder $Builder -Indent $indent -Text "<section id=""$SectionId"">"
        Write-IndentedLine -Builder $Builder -Indent ($indent + 1) -Text "<title>$(Escape-XmlText $TitleText)</title>"
        $Stack.Push($Level) | Out-Null
    }

    foreach ($item in $ParagraphItems) {
        $ditaContent = $item.DitaContent
        $text = $item.Text
        $headingLevel = $item.HeadingLevel
        $mapping = $item.Mapping
        $images = $item.Images
        $tableXml = $item.TableXml

        # Emit table if present
        if ($tableXml) {
            Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
            [void]$sb.Append($tableXml)
            continue
        }

        if ($headingLevel -eq 1) {
            if (-not $topicTitleWritten) {
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<title>$(Escape-XmlText $text)</title>"
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<body>"
                $topicTitleWritten = $true
                continue
            }
            else {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Close-SectionsToLevel -Builder $sb -Stack $openSections -TargetLevel 1
                $sectionCounter++
                $secId = New-SafeId -Text $text
                Open-Section -Builder $sb -Stack $openSections -Level 1 -TitleText $text -SectionId "${secId}_${sectionCounter}"
                continue
            }
        }

        if ($headingLevel -gt 1) {
            Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
            Close-SectionsToLevel -Builder $sb -Stack $openSections -TargetLevel $headingLevel
            $sectionCounter++
            $secId = New-SafeId -Text $text
            Open-Section -Builder $sb -Stack $openSections -Level $headingLevel -TitleText $text -SectionId "${secId}_${sectionCounter}"
            continue
        }

        # Emit images inline
        if ($images -and $images.Count -gt 0) {
            Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
            foreach ($img in $images) {
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<image href=""$img""/>"
            }
        }

        switch ($mapping.type) {
            "p" {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<p>$ditaContent</p>"
            }

            "note" {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                $noteType = if ($mapping.outputType) { $mapping.outputType } else { "note" }
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<note type=""$noteType""><p>$ditaContent</p></note>"
            }

            "codeblock" {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<codeblock>$(Escape-XmlText $text)</codeblock>"
            }

            "bullet" {
                if (-not $listOpen -or $listType -ne "ul") {
                    Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                    Write-IndentedLine -Builder $sb -Indent 1 -Text "<ul>"
                    $listOpen = $true
                    $listType = "ul"
                }
                Write-IndentedLine -Builder $sb -Indent 2 -Text "<li><p>$ditaContent</p></li>"
            }

            "numbered" {
                if (-not $listOpen -or $listType -ne "ol") {
                    Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                    Write-IndentedLine -Builder $sb -Indent 1 -Text "<ol>"
                    $listOpen = $true
                    $listType = "ol"
                }
                Write-IndentedLine -Builder $sb -Indent 2 -Text "<li><p>$ditaContent</p></li>"
            }

            "caption" {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<fig><title>Caption</title><desc>$ditaContent</desc></fig>"
            }

            default {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<p>$ditaContent</p>"
            }
        }
    }

    if ($listOpen) {
        if ($listType -eq "ul") {
            Write-IndentedLine -Builder $sb -Indent 1 -Text "</ul>"
        }
        elseif ($listType -eq "ol") {
            Write-IndentedLine -Builder $sb -Indent 1 -Text "</ol>"
        }
    }

    while ($openSections.Count -gt 0) {
        $indent = $openSections.Count
        Write-IndentedLine -Builder $sb -Indent $indent -Text "</section>"
        [void]$openSections.Pop()
    }

    if ($topicTitleWritten) {
        Write-IndentedLine -Builder $sb -Indent 1 -Text "</body>"
    }
    [void]$sb.AppendLine('</topic>')
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if (-not (Test-Path $StyleMapJson)) {
    throw "Style map JSON not found: $StyleMapJson"
}

$styleMap = Get-Content -Path $StyleMapJson -Raw | ConvertFrom-Json

# Read document.xml and relationship map via ZipArchive
$stream = [System.IO.File]::Open($InputDocx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $docEntry = $zip.GetEntry('word/document.xml')
        if (-not $docEntry) { throw "word/document.xml not found." }
        $reader = New-Object System.IO.StreamReader($docEntry.Open(), [System.Text.Encoding]::UTF8)
        try { $docXmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }

        # Build relationship map for images and hyperlinks
        $relMap = @{}
        $relsEntry = $zip.GetEntry('word/_rels/document.xml.rels')
        if ($relsEntry) {
            $relsReader = New-Object System.IO.StreamReader($relsEntry.Open(), [System.Text.Encoding]::UTF8)
            try { $relsText = $relsReader.ReadToEnd() } finally { $relsReader.Dispose() }
            $relsMatches = [regex]::Matches($relsText, 'Id="([^"]+)"[^>]*Target="([^"]+)"')
            foreach ($rm in $relsMatches) {
                $relMap[$rm.Groups[1].Value] = $rm.Groups[2].Value
            }
        }
    } finally { $zip.Dispose() }
} finally { $stream.Dispose() }

[xml]$docXml = [xml]$docXmlText
$nsMgr = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
$nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
$nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
$nsMgr.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
$nsMgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
$nsMgr.AddNamespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")

# Process body children: paragraphs and tables
$bodyNodes = $docXml.SelectNodes("//w:body/*", $nsMgr)
if ($null -eq $bodyNodes -or $bodyNodes.Count -eq 0) {
    throw "No body content found."
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$topics = New-Object System.Collections.Generic.List[object]
$currentTopicTitle = $null
$currentParagraphs = $null

foreach ($node in $bodyNodes) {
    # Handle tables
    if ($node.LocalName -eq "tbl") {
        if ($null -eq $currentParagraphs) {
            $currentTopicTitle = "topic"
            $currentParagraphs = New-Object System.Collections.Generic.List[object]
        }
        $tableXml = Convert-TableToDita -Table $node -NsMgr $nsMgr
        $currentParagraphs.Add([pscustomobject]@{
            Text = ""
            DitaContent = ""
            StyleName = ""
            HeadingLevel = 0
            Mapping = [pscustomobject]@{ type = "table" }
            Images = @()
            TableXml = $tableXml
        }) | Out-Null
        continue
    }

    if ($node.LocalName -ne "p") { continue }

    $p = $node
    $text = Get-ParagraphPlainText -Paragraph $p -NsMgr $nsMgr
    $ditaContent = Get-ParagraphDitaContent -Paragraph $p -NsMgr $nsMgr -RelMap $relMap
    $images = Get-ParagraphImages -Paragraph $p -NsMgr $nsMgr -RelMap $relMap

    if ([string]::IsNullOrWhiteSpace($text) -and $images.Count -eq 0) {
        continue
    }

    $styleName = Get-ParagraphStyle -Paragraph $p -NsMgr $nsMgr
    $headingLevel = Get-HeadingLevel -StyleName $styleName

    if ($styleMap.PSObject.Properties.Name -contains $styleName) {
        $mapping = $styleMap.$styleName
    }
    else {
        $mapping = [pscustomobject]@{ type = "p" }
    }

    if ($headingLevel -eq 1) {
        if ($null -ne $currentParagraphs) {
            $topics.Add([pscustomobject]@{
                Title = $currentTopicTitle
                Paragraphs = $currentParagraphs
            }) | Out-Null
        }

        $currentTopicTitle = $text
        $currentParagraphs = New-Object System.Collections.Generic.List[object]
        continue
    }

    if ($null -eq $currentParagraphs) {
        $currentTopicTitle = "topic"
        $currentParagraphs = New-Object System.Collections.Generic.List[object]
    }

    $currentParagraphs.Add([pscustomobject]@{
        Text = $text
        DitaContent = $ditaContent
        StyleName = $styleName
        HeadingLevel = $headingLevel
        Mapping = $mapping
        Images = $images
        TableXml = $null
    }) | Out-Null
}

if ($null -ne $currentParagraphs) {
    $topics.Add([pscustomobject]@{
        Title = $currentTopicTitle
        Paragraphs = $currentParagraphs
    }) | Out-Null
}

$mapBuilder = New-Object System.Text.StringBuilder
[void]$mapBuilder.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$mapBuilder.AppendLine('<!DOCTYPE map PUBLIC "-//OASIS//DTD DITA Map//EN" "map.dtd">')
[void]$mapBuilder.AppendLine('<map>')
[void]$mapBuilder.AppendLine('  <title>Converted Topics</title>')

$usedNames = @{}

foreach ($topic in $topics) {
    $safeName = New-SafeFileName -Text $topic.Title
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = "topic"
    }

    $finalName = $safeName
    $i = 1
    while ($usedNames.ContainsKey($finalName.ToLower())) {
        $finalName = "${safeName}_$i"
        $i++
    }
    $usedNames[$finalName.ToLower()] = $true

    $topicFileName = "$finalName.dita"
    $topicPath = Join-Path $OutputDir $topicFileName
    $topicId = New-SafeId -Text $finalName

    $xmlText = Convert-ParagraphsToTopicXml -ParagraphItems $topic.Paragraphs -TopicId $topicId
    $xmlText | Set-Content -Path $topicPath -Encoding UTF8

    [void]$mapBuilder.AppendLine("  <topicref href=""$topicFileName""/>")
}

[void]$mapBuilder.AppendLine('</map>')
$mapPath = Join-Path $OutputDir "map.ditamap"
$mapBuilder.ToString() | Set-Content -Path $mapPath -Encoding UTF8

Write-Host "Done."
Write-Host "Topics written to: $OutputDir"
Write-Host "DITA map written to: $mapPath"
