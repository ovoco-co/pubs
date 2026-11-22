<#
.SYNOPSIS
    Repoints paragraphs from author styles onto NIST styles, and gives unstyled
    paragraphs a style. Edits word/document.xml only; never touches styles.xml.

.DESCRIPTION
    This is the step that has to run BEFORE Set-NISTStyleArchitecture.ps1.
    The architecture never edits document.xml, and it SKIPS a rebase whose
    target style is absent (it warns "Rebase SKIPPED"), so any paragraph still
    pointing at an author style the architecture deletes falls back to Normal.

    Three kinds of change, each reported separately:

      1. named remaps - an author style to its NIST equivalent
      2. unstyled paragraphs with text, outside tables, to -BodyStyle
      3. unstyled paragraphs inside a table, to -TableStyle

    Empty unstyled paragraphs are left alone by default. They are structural
    spacers, and styling them changes the gap they were inserted to make. Pass
    -StyleEmpty to include them.

    CAPTIONS ARE EXCLUDED BY DEFAULT, and this matters more than it looks. A
    caption paragraph with no style is still a caption. Sweeping it into
    BodyStyle destroys the only signal a later caption pass has: it then looks
    for a caption STYLE, finds none, and inserts a placeholder caption beside a
    caption that was already there. -IncludeCaptionText turns the exclusion off
    if you are deliberately handling captions some other way.

    The default remap table is the set the architecture does not define:

      Bibliography  -> References    hanging-indent reference list
      AppendixText  -> BodyText      defined as basedOn BodyText, no overrides
      NormalWeb     -> BodyText      HTML paste residue, semiHidden
      IntenseQuote  -> Quote         the certified set uses Quote

    Add to it with -Remap 'Old=New','Other=New'. A remap whose target style is
    not in the file is skipped and reported rather than written, because that
    would strand the paragraph on a missing style.

    Table styles outside the architecture's allowlist are repointed to
    TableGrid, since the architecture deletes every table style except
    TableNormal, TableGrid, HM-Cover, HM-Data, NIST-Data and SidebarTable.

    Guards: visible text must be byte-identical, the paragraph count must not
    change, and the result must parse. Any failure and nothing is written.

.PARAMETER Remap
    Additional 'OldStyleId=NewStyleId' pairs, appended to the defaults.

.PARAMETER NoDefaultRemap
    Use only the pairs given in -Remap.

.PARAMETER BodyStyle
    Style for unstyled body paragraphs. Default BodyText.

.PARAMETER TableStyle
    Style for unstyled paragraphs inside tables. Default TableText.

.PARAMETER StyleEmpty
    Also style empty unstyled paragraphs.

.PARAMETER IncludeCaptionText
    Do not exclude caption-shaped paragraphs from the unstyled sweep.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Set-StyleRemap.ps1 -Folder "C:\...\HM-261109" -ListOnly

.EXAMPLE
    .\Set-StyleRemap.ps1 -File "C:\...\Chapter-05.docx" -Remap 'OldBody=BodyText' -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string[]]$Remap = @(),
    [switch]$NoDefaultRemap,
    [string]$BodyStyle  = 'BodyText',
    [string]$TableStyle = 'TableText',
    [switch]$StyleEmpty,
    [switch]$IncludeCaptionText,
    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$defaultRemap = [ordered]@{
    'Bibliography' = 'References'
    'AppendixText' = 'BodyText'
    'NormalWeb'    = 'BodyText'
    'IntenseQuote' = 'Quote'
}

# ListParagraph is Word's generic "indented paragraph that might be a list". It
# carries no glyph of its own, so an item on it shows nothing, and "every other
# list has its own style - bullet 1, 2, 3, numbered 1, 2, 3".
#
# It is remapped only when the paragraph has NO w:numPr. One that still has a
# numPr is a real list item, and Set-ContentStyles.ps1 moves it to List Bullet
# or List Number by the numFmt its numId resolves to - sending it to Body Text
# here would destroy the only evidence of which kind of list it was.
$remapIfNoNumbering = @{ 'ListParagraph' = 'BodyText' }
$map = [ordered]@{}
if (-not $NoDefaultRemap) { foreach ($k in $defaultRemap.Keys) { $map[$k] = $defaultRemap[$k] } }
foreach ($pair in $Remap) {
    if ($pair -notmatch '^\s*([^=]+?)\s*=\s*(.+?)\s*$') { throw "Bad -Remap entry (want 'Old=New'): $pair" }
    $map[$Matches[1]] = $Matches[2]
}

# Every table style the architecture keeps. Anything else it deletes.
$tableKeep = @('TableNormal','TableGrid','HM-Cover','HM-Data','NIST-Data','SidebarTable')

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $resolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($resolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|docm)$' -and $_.Name -notlike '~$*' } |
        ForEach-Object { $_.FullName } | Sort-Object
}
if ($targets.Count -eq 0) { throw "No Word files found." }

Write-Host "Set Style Remap" -ForegroundColor Cyan
Write-Host "  Edits word/document.xml only. Run BEFORE Set-NISTStyleArchitecture.ps1." -ForegroundColor Cyan
foreach ($k in $map.Keys) { Write-Host ("    {0,-16} -> {1}" -f $k, $map[$k]) -ForegroundColor DarkGray }
Write-Host ("    unstyled body    -> {0}" -f $BodyStyle) -ForegroundColor DarkGray
Write-Host ("    unstyled in table-> {0}" -f $TableStyle) -ForegroundColor DarkGray
if (-not $IncludeCaptionText) { Write-Host "    caption-shaped paragraphs are EXCLUDED from the unstyled sweep" -ForegroundColor DarkGray }
Write-Host ""
if ($ListOnly) { Write-Host "  -ListOnly: reporting only, writing nothing." -ForegroundColor Yellow }
if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

function Read-Package {
    param([string]$Path)
    $names = New-Object System.Collections.ArrayList
    $entries = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open,
              [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open(); try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
                [void]$names.Add($e.FullName)
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
    return ,[pscustomobject]@{ Names = $names; Data = $entries }
}

function Get-VisibleText {
    param([string]$Xml)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in [regex]::Matches($Xml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { [void]$sb.Append($t.Groups[1].Value) }
    return $sb.ToString()
}

# <w:pStyle w:val="X"/> and <w:pStyle w:val="X" /> are the same element and one
# document set contains both, so every pattern here tolerates the space and
# every replacement spans the WHOLE element.
$pStyleRx = '<w:pStyle(?:\s[^>]*?)?\s*/>'
function Get-StyleId {
    param([string]$Block)
    $m = [regex]::Match($Block, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

# A caption is label + number + period or colon + title. Prose that merely
# mentions one ("Figure 5 shows the outages") is not a caption.
$captionRx = '^\s*(?:Figure|Table)\s*\u00a0?\s*(?:[A-Z]?\d+(?:[.\-\u2013]\d+)?|[A-Z][.\-\u2013]\d+)\s*[.:]\s+\S'

$filesChanged = 0
foreach ($target in $targets) {
    $name = [System.IO.Path]::GetFileName($target)
    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    if (-not $pkg.Data.ContainsKey('word/document.xml')) { Write-Host "  No document.xml." -ForegroundColor DarkGray; continue }
    $raw = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/document.xml'])
    $hadBom = ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF)
    if ($hadBom) { $raw = $raw.Substring(1) }
    $stylesXml = if ($pkg.Data.ContainsKey('word/styles.xml')) {
        [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/styles.xml']) } else { '' }

    $present = @{}
    foreach ($m in [regex]::Matches($stylesXml, 'w:styleId="([^"]+)"')) { $present[$m.Groups[1].Value] = $true }

    $beforeText = Get-VisibleText -Xml $raw
    $beforeParas = ([regex]::Matches($raw, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>')).Count

    # table spans, so an unstyled paragraph can be told apart from a cell
    $tblSpans = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($raw, '(?s)<w:tbl>.*?</w:tbl>')) {
        # @($a, $b + $c) is THREE elements: in an array literal the comma binds
        # tighter than +, so this parsed as @(Index, Index, +Length) and every
        # containment test came out false - 1,017 table cells were classified as
        # body text. Parenthesise the sum.
        [void]$tblSpans.Add(@($m.Index, ($m.Index + $m.Length)))
    }
    function Test-InTable { param([int]$i) foreach ($s in $tblSpans) { if ($i -ge $s[0] -and $i -lt $s[1]) { return $true } } return $false }

    $tally = [ordered]@{}
    $skipped = @{}
    $sb = New-Object System.Text.StringBuilder
    $pos = 0

    foreach ($pm in [regex]::Matches($raw, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>')) {
        $p = $pm.Value
        $sid = Get-StyleId -Block $p
        $new = $p

        if ($sid) {
            # ListParagraph goes to Body Text ONLY when the paragraph has lost
            # its numbering. With a numPr still on it, it is a live list item
            # and Set-ContentStyles.ps1 decides between List Bullet and List
            # Number from the numFmt its numId resolves to.
            if ($remapIfNoNumbering.ContainsKey($sid) -and $p -notmatch '<w:numPr>') {
                $tgt = $remapIfNoNumbering[$sid]
                if (-not $present.ContainsKey($tgt)) {
                    $skipped["$sid -> $tgt (target style absent)"] = 1 + [int]$skipped["$sid -> $tgt (target style absent)"]
                } else {
                    $sm = [regex]::Match($p, $pStyleRx)
                    $new = $p.Remove($sm.Index, $sm.Length).Insert($sm.Index, "<w:pStyle w:val=`"$tgt`"/>")
                    $k = "$sid -> $tgt (no numbering)"; $tally[$k] = 1 + [int]$tally[$k]
                }
            } elseif ($remapIfNoNumbering.ContainsKey($sid)) {
                $skipped["$sid kept: still has a numPr, for the list pass"] =
                    1 + [int]$skipped["$sid kept: still has a numPr, for the list pass"]
            } elseif ($map.Contains($sid)) {
                $tgt = $map[$sid]
                if (-not $present.ContainsKey($tgt)) {
                    $skipped["$sid -> $tgt (target style absent)"] = 1 + [int]$skipped["$sid -> $tgt (target style absent)"]
                } else {
                    $sm = [regex]::Match($p, $pStyleRx)
                    $new = $p.Remove($sm.Index, $sm.Length).Insert($sm.Index, "<w:pStyle w:val=`"$tgt`"/>")
                    $k = "$sid -> $tgt"; $tally[$k] = 1 + [int]$tally[$k]
                }
            }
        } else {
            $txt = (Get-VisibleText -Xml $p).Trim()
            $inTbl = Test-InTable -i $pm.Index
            $tgt = $null
            if ($inTbl) {
                if ($txt -or $StyleEmpty) { $tgt = $TableStyle }
            } elseif ($txt) {
                if (-not $IncludeCaptionText -and $txt -match $captionRx) {
                    $skipped['caption-shaped, left unstyled for the caption pass'] =
                        1 + [int]$skipped['caption-shaped, left unstyled for the caption pass']
                } else { $tgt = $BodyStyle }
            } elseif ($StyleEmpty) { $tgt = $BodyStyle }

            if ($tgt) {
                if (-not $present.ContainsKey($tgt)) {
                    $skipped["unstyled -> $tgt (target style absent)"] = 1 + [int]$skipped["unstyled -> $tgt (target style absent)"]
                } else {
                    # CT_PPr order: pStyle is the FIRST child of pPr
                    $pprOpen = [regex]::Match($p, '(?s)^(<w:p\b[^>]*?>)(\s*<w:pPr\b[^>]*?>)')
                    if ($pprOpen.Success) {
                        $new = $p.Insert($pprOpen.Groups[2].Index + $pprOpen.Groups[2].Length, "<w:pStyle w:val=`"$tgt`"/>")
                    } else {
                        # A SELF-CLOSING paragraph has no body to insert into:
                        # '<w:p/>' + '<w:pPr>' is malformed, because the element
                        # is already closed. Reopen it and give it a real body.
                        $sc = [regex]::Match($p, '(?s)^<w:p\b([^>]*?)/>$')
                        if ($sc.Success) {
                            $new = '<w:p' + $sc.Groups[1].Value + '>' +
                                   "<w:pPr><w:pStyle w:val=`"$tgt`"/></w:pPr></w:p>"
                        } else {
                            $open = [regex]::Match($p, '(?s)^<w:p\b[^>]*?>')
                            $new = $p.Insert($open.Length, "<w:pPr><w:pStyle w:val=`"$tgt`"/></w:pPr>")
                        }
                    }
                    $k = if ($inTbl) { "unstyled in table -> $tgt" } else { "unstyled -> $tgt" }
                    $tally[$k] = 1 + [int]$tally[$k]
                }
            }
        }

        if ($new -ne $p) {
            [void]$sb.Append($raw.Substring($pos, $pm.Index - $pos))
            [void]$sb.Append($new)
            $pos = $pm.Index + $pm.Length
        }
    }
    [void]$sb.Append($raw.Substring($pos))
    $newDoc = $sb.ToString()

    # table styles the architecture would delete
    $tblFixed = 0
    foreach ($m in [regex]::Matches($newDoc, '<w:tblStyle\b[^>]*\bw:val="([^"]+)"[^>]*/>')) {
        if ($tableKeep -notcontains $m.Groups[1].Value) { $tblFixed++ }
    }
    if ($tblFixed -gt 0) {
        $newDoc = [regex]::Replace($newDoc, '<w:tblStyle\b[^>]*\bw:val="([^"]+)"[^>]*/>', {
            param($m)
            if ($tableKeep -contains $m.Groups[1].Value) { return $m.Value }
            return '<w:tblStyle w:val="TableGrid"/>'
        })
        $tally['table style -> TableGrid'] = $tblFixed
    }

    foreach ($k in $skipped.Keys) { Write-Host ("  SKIPPED x{0}: {1}" -f $skipped[$k], $k) -ForegroundColor DarkYellow }

    if ($tally.Count -eq 0) { Write-Host "  Nothing to remap. File unchanged." -ForegroundColor Green; continue }
    foreach ($k in $tally.Keys) { Write-Host ("  {0,-44} {1}" -f $k, $tally[$k]) -ForegroundColor DarkGray }

    if ((Get-VisibleText -Xml $newDoc) -ne $beforeText) { throw "Rewrite changed visible text in ${name}; nothing written." }
    $afterParas = ([regex]::Matches($newDoc, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>')).Count
    if ($afterParas -ne $beforeParas) { throw "Paragraph count changed in ${name} ($beforeParas -> $afterParas); nothing written." }
    try {
        $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml)
    } catch { throw "Rewrite produced malformed document.xml for ${name}: $($_.Exception.Message)" }

    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        $bak = "$target.remap.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
        else { Write-Warning "  Backup exists; leaving it as-is." }
    }

    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "remap-$([guid]::NewGuid()).zip")
    $out = [System.IO.File]::Create($tempZip)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($out, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($n in $pkg.Names) {
                $e = $zip.CreateEntry($n, [System.IO.Compression.CompressionLevel]::Optimal)
                $s = $e.Open(); try { $s.Write($pkg.Data[$n], 0, $pkg.Data[$n].Length) } finally { $s.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $out.Dispose() }

    try {
        Move-Item -LiteralPath $tempZip -Destination $target -Force
        $filesChanged++
        Write-Host "  Saved." -ForegroundColor Green
    } catch {
        Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue
        Write-Host "  LOCKED (open in Word) - not written." -ForegroundColor Red
    }
}

Write-Host ""
Write-Host ("Done. {0} file(s) changed." -f $filesChanged) -ForegroundColor Green
