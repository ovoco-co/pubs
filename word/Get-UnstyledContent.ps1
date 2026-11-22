<#
.SYNOPSIS
    Inventories every paragraph that has no style, or is on Normal, and
    classifies it by shape so each class can be dealt with deliberately.
    Read-only: writes nothing.

.DESCRIPTION
    Phase A of bringing an author chapter set into the template. Run this
    BEFORE Set-StyleRemap.ps1 and read the output.

    Unstyled content is not one thing. In one chapter set 77 percent of all
    paragraphs carried no pStyle at all, and sweeping them onto BodyText is
    what destroyed the caption signal: caption paragraphs became BodyText, a
    later caption pass looked for a caption STYLE, found none, and inserted 50
    placeholder captions beside captions that were already there.

    So classify first, decide per class, then apply.

    The class -> style mapping is not invented here. It was measured against
    the last certified chapter set, which is unambiguous for most classes:

        table cell        TableText                  13849 of 14017
        image only        graphic                       69 of 72
        caption           CaptionTable/CaptionFigure   195 of 195
        xref prose        BodyText                      74 of 74
        reference entry   References                   308 of 308
        empty             left on Normal               128 of 149
        body prose        BodyText                     majority

    Two classes are deliberately NOT given a target:

      heading-shaped  In the certified set these are spread across Heading3
                      (92), Heading4 (81), ListBullet (68) and Heading2 (33).
                      The shape alone does not say which, so this script
                      reports them and the level is a human decision.

      panel label     A short caption-like line sharing a paragraph with an
                      image, as in a multi-panel figure. The certified set has
                      none, so there is no evidence to copy.

    Classifier notes:

      A CAPTION is label + number + period or colon + title ("Figure 10.
      Example of cable damage"). XREF PROSE mentions one without terminating
      the number ("Figure 5 shows the outages"). 25 of 55 candidates in one
      chapter were prose, and styling those as captions is wrong.

      A LIST ITEM is recognised by a numPr only on an UNSTYLED paragraph.
      Heading2/3/4 carry a numPr from the HMList chain, so on a styled
      paragraph a numPr says nothing about lists.

.PARAMETER IncludeStyled
    Also report paragraphs that already carry a style, so the classifier can be
    checked against a document known to be correct.

.PARAMETER Samples
    Sample lines to show per class. Default 3, 0 for counts only.

.EXAMPLE
    .\Get-UnstyledContent.ps1 -Folder "C:\...\HM-261109"

.EXAMPLE
    # check the classifier against a document set that is already right
    .\Get-UnstyledContent.ps1 -Folder "C:\...\HM-261102" -IncludeStyled
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$IncludeStyled,
    [int]$Samples = 3
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $resolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($resolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|docm|dotx)$' -and $_.Name -notlike '~$*' } |
        ForEach-Object { $_.FullName } | Sort-Object
}
if ($targets.Count -eq 0) { throw "No Word files found." }

# class -> what to put it on. $null means report only, no target.
$targetStyle = [ordered]@{
    'text box'        = '(leave alone)'
    'table cell'      = 'TableText'
    'image only'      = 'graphic'
    'caption'         = 'CaptionFigure / CaptionTable (by label)'
    'xref prose'      = 'BodyText'
    'reference entry' = 'References'
    'list item'       = 'ListBullet / ListNumber (by numFmt)'
    'body prose'      = 'BodyText'
    'empty'           = '(leave on Normal)'
    'panel label'     = $null
    'heading-shaped'  = $null
}

function Read-Part {
    param([string]$Path, [string]$Part)
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open,
              [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $e = $zip.GetEntry($Part)
            if (-not $e) { return $null }
            $r = New-Object System.IO.StreamReader($e.Open())
            try { $text = $r.ReadToEnd() } finally { $r.Dispose() }
            if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
            return $text
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
}

function Get-Text {
    param([string]$Block)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in [regex]::Matches($Block, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { [void]$sb.Append($t.Groups[1].Value) }
    return $sb.ToString()
}

# label + number + terminator + title
$capRx  = '^\s*(?:Figure|Table)\s*\u00a0?\s*(?:[A-Z]?\d+(?:[.\-\u2013]\d+)?|[A-Z][.\-\u2013]\d+)\s*[.:]\s+\S'
# label + number then lower-case prose: a mention, not a caption
$xrefRx = '^\s*(?:Figure|Table)\s*\u00a0?\s*[A-Z]?\d+(?:[.\-\u2013]\d+)?\s+[a-z]'
# Author, A. B. (2019).
$refRx  = '^[A-Z][A-Za-z\u00c0-\u024f''\- ]+,\s+[A-Z]\.'

$grand = [ordered]@{}
foreach ($k in $targetStyle.Keys) { $grand[$k] = 0 }

foreach ($target in $targets) {
    $name = [System.IO.Path]::GetFileName($target)
    $doc = Read-Part -Path $target -Part 'word/document.xml'
    if (-not $doc) { continue }

    $tbl = New-Object System.Collections.ArrayList
    # @($a, $b + $c) is THREE elements: inside an array literal the comma
    # binds tighter than +, so this parsed as @(Index, Index, +Length) and
    # $s[1] was the START offset again - every containment test came out
    # false and the whole classification was wrong. Parenthesise the sum.
    foreach ($m in [regex]::Matches($doc, '(?s)<w:tbl>.*?</w:tbl>')) { [void]$tbl.Add(@($m.Index, ($m.Index + $m.Length))) }

    # Text-box content is figure furniture: chart axis labels, date ticks,
    # callouts, panel letters. It is not body text and it is not a heading, and
    # it does not sit inside w:tbl so it falls through every other test. Ch04
    # has 34 text boxes, and their contents were what made "heading-shaped"
    # fire 884 times across the set - '9/7/17', 'Company',
    # '% of cell sites out of service by municipio'. Leave them alone.
    $txbx = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($doc, '(?s)<w:txbxContent>.*?</w:txbxContent>')) {
        [void]$txbx.Add(@($m.Index, ($m.Index + $m.Length)))
    }

    $counts = [ordered]@{}
    foreach ($k in $targetStyle.Keys) { $counts[$k] = 0 }
    $sampleLines = @{}
    $styledSeen = @{}
    $total = 0

    foreach ($pm in [regex]::Matches($doc, '(?s)<w:p\b[^>]*?>.*?</w:p>|<w:p\b[^>]*?/>')) {
        $p = $pm.Value
        $sm = [regex]::Match($p, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        $sid = if ($sm.Success) { $sm.Groups[1].Value } else { $null }
        # Normal counts as unstyled: it is the fallback, not a choice
        $isUnstyled = (-not $sid) -or ($sid -eq 'Normal')
        if (-not $isUnstyled -and -not $IncludeStyled) { continue }
        $total++

        $t = (Get-Text -Block $p).Trim()
        $inTbl = $false
        foreach ($s in $tbl) { if ($pm.Index -ge $s[0] -and $pm.Index -lt $s[1]) { $inTbl = $true; break } }
        $hasImg = $p -match '<a:blip\b|<w:pict\b'
        $hasNum = $p -match '<w:numPr>'

        $inTxbx = $false
        foreach ($s2 in $txbx) { if ($pm.Index -ge $s2[0] -and $pm.Index -lt $s2[1]) { $inTxbx = $true; break } }

        $class =
            if ($inTxbx)                                 { 'text box' }
            elseif ($inTbl)                              { 'table cell' }
            elseif (-not $t -and $hasImg)                { 'image only' }
            elseif (-not $t)                             { 'empty' }
            elseif ($t -match $capRx)                    { 'caption' }
            elseif ($t -match $xrefRx)                   { 'xref prose' }
            elseif ($t -match $refRx)                    { 'reference entry' }
            elseif ($hasImg -and $t.Length -lt 40)       { 'panel label' }
            elseif ($hasNum -and $isUnstyled)            { 'list item' }
            elseif ($t.Length -lt 60 -and $t -notmatch '[.!?:;]\s*$') { 'heading-shaped' }
            else                                         { 'body prose' }

        $counts[$class]++
        $grand[$class]++
        if ($IncludeStyled -and $sid) {
            if (-not $styledSeen.ContainsKey($class)) { $styledSeen[$class] = @{} }
            $styledSeen[$class][$sid] = 1 + [int]$styledSeen[$class][$sid]
        }
        if ($Samples -gt 0 -and $t) {
            if (-not $sampleLines.ContainsKey($class)) { $sampleLines[$class] = New-Object System.Collections.ArrayList }
            if ($sampleLines[$class].Count -lt $Samples) {
                [void]$sampleLines[$class].Add($t.Substring(0, [Math]::Min(68, $t.Length)))
            }
        }
    }

    Write-Host ""
    Write-Host $name -ForegroundColor White
    if ($total -eq 0) { Write-Host "  nothing unstyled" -ForegroundColor Green; continue }
    foreach ($k in $counts.Keys) {
        if ($counts[$k] -eq 0) { continue }
        $tgt = $targetStyle[$k]
        if ($null -eq $tgt) {
            Write-Host ("  {0,-17}{1,6}   -> DECIDE (no safe default)" -f $k, $counts[$k]) -ForegroundColor Yellow
        } else {
            Write-Host ("  {0,-17}{1,6}   -> {2}" -f $k, $counts[$k], $tgt) -ForegroundColor DarkGray
        }
        if ($IncludeStyled -and $styledSeen.ContainsKey($k)) {
            $obs = ($styledSeen[$k].GetEnumerator() | Sort-Object Value -Descending |
                    Select-Object -First 4 | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '  '
            Write-Host ("  {0,-17}{1,6}      observed: {2}" -f '', '', $obs) -ForegroundColor DarkCyan
        }
        if ($sampleLines.ContainsKey($k)) {
            foreach ($s in $sampleLines[$k]) { Write-Host ("        {0}" -f $s) -ForegroundColor DarkGray }
        }
    }
}

Write-Host ""
Write-Host "TOTAL across all files" -ForegroundColor Cyan
foreach ($k in $grand.Keys) {
    if ($grand[$k] -eq 0) { continue }
    $tgt = $targetStyle[$k]
    $color = if ($null -eq $tgt) { 'Yellow' } else { 'Gray' }
    Write-Host ("  {0,-17}{1,6}   -> {2}" -f $k, $grand[$k], $(if ($null -eq $tgt) { 'DECIDE' } else { $tgt })) -ForegroundColor $color
}
Write-Host ""
Write-Host "Nothing was written. Decide the DECIDE classes, then run Set-StyleRemap.ps1." -ForegroundColor Cyan
