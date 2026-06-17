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

function Escape-XmlText {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Get-ParagraphText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Paragraph,
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $runs = $Paragraph.SelectNodes(".//w:t", $NsMgr)
    $texts = @()

    foreach ($run in $runs) {
        if ($null -ne $run.InnerText) {
            $texts += $run.InnerText
        }
    }

    return ($texts -join "").Trim()
}

function Get-ParagraphStyle {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Paragraph,
        [Parameter(Mandatory = $true)]
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
        '^Heading 1$' { return 1 }
        '^Heading 2$' { return 2 }
        '^Heading 3$' { return 3 }
        '^Heading 4$' { return 4 }
        '^Heading 5$' { return 5 }
        default { return 0 }
    }
}

function New-SafeFileName {
    param([string]$Text)

    $safe = $Text -replace '[\\/:*?"<>|]', '_' `
                   -replace '\s+', '_' `
                   -replace '_+', '_'

    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = "topic"
    }

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

function Convert-ParagraphsToTopicXml {
    param(
        [Parameter(Mandatory = $true)]
        $ParagraphItems
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    [void]$sb.AppendLine('<topic xmlns="http://dita.oasis-open.org/architecture/2005/">')

    $topicTitleWritten = $false
    $openSections = New-Object System.Collections.Generic.Stack[int]
    $listOpen = $false
    $listType = $null

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
            [string]$TitleText
        )

        $indent = $Level
        Write-IndentedLine -Builder $Builder -Indent $indent -Text "<section>"
        Write-IndentedLine -Builder $Builder -Indent ($indent + 1) -Text "<title>$(Escape-XmlText $TitleText)</title>"
        $Stack.Push($Level) | Out-Null
    }

    foreach ($item in $ParagraphItems) {
        $text = $item.Text
        $headingLevel = $item.HeadingLevel
        $mapping = $item.Mapping

        if ($headingLevel -eq 1) {
            if (-not $topicTitleWritten) {
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<title>$(Escape-XmlText $text)</title>"
                $topicTitleWritten = $true
                continue
            }
            else {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Close-SectionsToLevel -Builder $sb -Stack $openSections -TargetLevel 1
                Open-Section -Builder $sb -Stack $openSections -Level 1 -TitleText $text
                continue
            }
        }

        if ($headingLevel -gt 1) {
            Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
            Close-SectionsToLevel -Builder $sb -Stack $openSections -TargetLevel $headingLevel
            Open-Section -Builder $sb -Stack $openSections -Level $headingLevel -TitleText $text
            continue
        }

        switch ($mapping.type) {
            "p" {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<p>$(Escape-XmlText $text)</p>"
            }

            "note" {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                $noteType = if ($mapping.outputType) { $mapping.outputType } else { "note" }
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<note type=""$noteType""><p>$(Escape-XmlText $text)</p></note>"
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
                Write-IndentedLine -Builder $sb -Indent 2 -Text "<li><p>$(Escape-XmlText $text)</p></li>"
            }

            "numbered" {
                if (-not $listOpen -or $listType -ne "ol") {
                    Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                    Write-IndentedLine -Builder $sb -Indent 1 -Text "<ol>"
                    $listOpen = $true
                    $listType = "ol"
                }
                Write-IndentedLine -Builder $sb -Indent 2 -Text "<li><p>$(Escape-XmlText $text)</p></li>"
            }

            "caption" {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<fig><title>Caption</title><desc>$(Escape-XmlText $text)</desc></fig>"
            }

            default {
                Close-List -Builder $sb -IsOpen ([ref]$listOpen) -Type ([ref]$listType)
                Write-IndentedLine -Builder $sb -Indent 1 -Text "<p>$(Escape-XmlText $text)</p>"
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

    [void]$sb.AppendLine('</topic>')
    return $sb.ToString()
}

if (-not (Test-Path $StyleMapJson)) {
    throw "Style map JSON not found: $StyleMapJson"
}

$styleMap = Get-Content -Path $StyleMapJson -Raw | ConvertFrom-Json

$tempRoot = Join-Path $env:TEMP ("docx_extract_" + [Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    $tempZip = Join-Path $tempRoot "input.zip"
    Copy-Item -Path $InputDocx -Destination $tempZip -Force
    Expand-Archive -Path $tempZip -DestinationPath $tempRoot -Force

    $documentPath = Join-Path $tempRoot "word\document.xml"
    if (-not (Test-Path $documentPath)) {
        throw "word/document.xml not found."
    }

    [xml]$docXml = Get-Content -Path $documentPath -Raw -Encoding UTF8
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
    $nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

    $paragraphs = $docXml.SelectNodes("//w:body/w:p", $nsMgr)
    if ($null -eq $paragraphs -or $paragraphs.Count -eq 0) {
        throw "No paragraphs found."
    }

    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir | Out-Null
    }

    $topics = New-Object System.Collections.Generic.List[object]
    $currentTopicTitle = $null
    $currentParagraphs = $null

    foreach ($p in $paragraphs) {
        $text = Get-ParagraphText -Paragraph $p -NsMgr $nsMgr
        if ([string]::IsNullOrWhiteSpace($text)) {
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
            StyleName = $styleName
            HeadingLevel = $headingLevel
            Mapping = $mapping
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
    [void]$mapBuilder.AppendLine('<map xmlns="http://dita.oasis-open.org/architecture/2005/">')
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

        $xmlText = Convert-ParagraphsToTopicXml -ParagraphItems $topic.Paragraphs
        $xmlText | Set-Content -Path $topicPath -Encoding UTF8

        [void]$mapBuilder.AppendLine("  <topicref href=""$topicFileName""/>")
    }

    [void]$mapBuilder.AppendLine('</map>')
    $mapPath = Join-Path $OutputDir "map.ditamap"
    $mapBuilder.ToString() | Set-Content -Path $mapPath -Encoding UTF8

    Write-Host "Done."
    Write-Host "Topics written to: $OutputDir"
    Write-Host "DITA map written to: $mapPath"
}
finally {
    if (Test-Path $tempRoot) {
        Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}