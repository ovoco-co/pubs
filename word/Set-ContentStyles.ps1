<#
.SYNOPSIS
    Puts figures, captions, reference entries and list items on their template
    styles. Step 3 of the author chapter set process, after the styles are in
    and Reapply-Styles.ps1 has cleared the direct formatting.

.DESCRIPTION
    Each rule below was measured against the last certified chapter set, not
    chosen. The counts are how many paragraphs of that shape carried that style
    there:

      image-only paragraph  -> graphic            69 of 72
      caption               -> CaptionFigure /    195 of 195
                               CaptionTable
      reference entry       -> References         308 of 308
      block quotation       -> Quote              wholly italic, opens
                                                  with a quotation mark
      list item             -> ListBullet /
                               ListNumber         by the numFmt it points at

    graphic carries keepNext, which is what holds a figure to the caption
    below it. CaptionFigure and CaptionTable are chosen by the caption's own
    label, not by position, because a caption separated from its object by a
    page-break paragraph is still that kind of caption.

    NOT TOUCHED, and reported instead:

      shape text      Paragraphs inside w:txbxContent are figure furniture -
                      axis labels, date ticks, callouts, panel letters. They
                      are made of w:p elements, so a flat scan lists them as
                      body paragraphs; they must be excluded by span.
      panel label     A short line sharing a paragraph with an image. The
                      certified set has none, so there is nothing to copy.
      heading-shaped  Short lines with no terminal punctuation. In the
                      certified set these are spread across Heading3, Heading4,
                      ListBullet and Heading2, so the shape does not say which.

    A caption is label + number + period or colon + title ("Figure 10. Example
    of cable damage"). Prose that merely mentions one ("Figure 5 shows the
    outages") is NOT a caption - 25 of 55 candidates in one chapter were prose,
    and styling those as captions puts a figure caption in the middle of a
    paragraph.

    A reference entry is only treated as one when it sits after a "References"
    heading. The author-date shape alone appears in body prose too.

    Two spaces follow the caption number, per the NIST form: label + NBSP +
    chapter + en dash + number + period + TWO spaces + title. The spaces are
    normalised inside the existing runs so italics and symbols in the title
    survive, and the run that ends up holding them gets xml:space="preserve" -
    without it Word collapses them and the fix silently undoes itself.

.PARAMETER Only
    Run just these rules: Graphic, Caption, Reference, List, CaptionSpacing.
    Default is all of them.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Set-ContentStyles.ps1 -Folder "C:\...\HM-261109-v2" -ListOnly

.EXAMPLE
    .\Set-ContentStyles.ps1 -Folder "C:\...\HM-261109-v2" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [ValidateSet('Graphic','Caption','Reference','List','CaptionSpacing','Quote')]
    [string[]]$Only,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$rules = if ($Only) { $Only } else { @('Graphic','Caption','Reference','List','CaptionSpacing','Quote') }

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

# A ~$ lock file means an author has that chapter open in Word. Writing to it
# raises an IOException, and with no check that exception aborted the whole run:
# one locked chapter left the other seven half-processed and inconsistent. SKIP
# it, say which, and carry on - a file the author is working in is theirs.
$lockedTargets = @()
foreach ($t in $targets) {
    $leaf = Split-Path -Leaf $t
    $lock = Join-Path (Split-Path -Parent $t) ('~$' + $leaf.Substring(2))
    if (Test-Path -LiteralPath $lock) { $lockedTargets += $t }
}
if ($lockedTargets.Count -gt 0) {
    foreach ($l in $lockedTargets) { Write-Warning "  OPEN IN WORD, skipped: $(Split-Path -Leaf $l)" }
    $targets = @($targets | Where-Object { $lockedTargets -notcontains $_ })
    if ($targets.Count -eq 0) { throw "Every target is open in Word. Nothing done." }
}


Write-Host "Set Content Styles" -ForegroundColor Cyan
Write-Host ("  Rules: " + ($rules -join ', ')) -ForegroundColor Cyan
Write-Host "  Shape text, panel labels and heading-shaped lines are reported, not changed." -ForegroundColor Cyan
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

# Text-producing nodes, in document order, each with its XML span and whether
# this script may edit it.
#
# w:t is NOT the only thing that puts characters on the page. A caption in
# correct NIST form reads
#
#     Table <STYLEREF 1 \s> <w:noBreakHyphen/> <SEQ Table> .  Title
#
# and reading only w:t gives "Table 32.  Title" - the hyphen is missing, so
# every text offset past it is one short. Computing the punctuation position
# from that put a period INSIDE the number and turned "Table 3-2." into
# "Table 3-.  2.". Any element that contributes a character has to be in the
# model, even though none of them is ever edited.
function Get-TextNodes {
    param([string]$Block)
    $nodes = New-Object System.Collections.ArrayList

    # Text inside a drawing belongs to the PICTURE, not to the paragraph: a
    # chart's axis labels, a callout, a panel letter. Counting it as paragraph
    # text is why an image paragraph carrying a text box looked like a text
    # paragraph - so the figure rule's "text under 40 characters" test failed
    # and the image never reached the graphic style. In NWIRP 2-1 Ch02 that left
    # a figure sitting on Body Text with its caption reported as misplaced.
    #
    # The spans are EXCLUDED rather than stripped, because the Index of every
    # node is used to splice the original paragraph when a caption prefix is
    # rewritten. Stripping first would shift every offset.
    $skip = New-Object System.Collections.ArrayList
    foreach ($d in [regex]::Matches($Block, '(?s)<w:drawing\b.*?</w:drawing>|<w:pict\b.*?</w:pict>')) {
        [void]$skip.Add(@($d.Index, ($d.Index + $d.Length)))
    }

    $rx = '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>|<w:noBreakHyphen\s*/>|<w:softHyphen\s*/>|<w:tab\s*/>|<w:br(?:\s[^>]*)?/>|<w:sym\b[^>]*/>'
    foreach ($m in [regex]::Matches($Block, $rx)) {
        $inPic = $false
        foreach ($sp in $skip) { if ($m.Index -ge $sp[0] -and $m.Index -lt $sp[1]) { $inPic = $true; break } }
        if ($inPic) { continue }
        $v = $m.Value
        # '<w:tab/>'.StartsWith('<w:t') is TRUE, and so is '<w:tbl>'. Test for
        # the w:t element itself - '<w:t>' or '<w:t ' - or a tab is taken for a
        # text run, gets Groups[1] (empty) as its text, and an edit lands on
        # top of the tab element.
        if ($v -match '^<w:t[ >]') {
            [void]$nodes.Add([pscustomobject]@{ Index = $m.Index; Length = $m.Length; Text = $m.Groups[1].Value; Editable = $true; Match = $m })
        } elseif ($v.StartsWith('<w:noBreakHyphen')) {
            [void]$nodes.Add([pscustomobject]@{ Index = $m.Index; Length = $m.Length; Text = [string][char]0x2011; Editable = $false; Match = $m })
        } elseif ($v.StartsWith('<w:softHyphen')) {
            [void]$nodes.Add([pscustomobject]@{ Index = $m.Index; Length = $m.Length; Text = ''; Editable = $false; Match = $m })
        } elseif ($v.StartsWith('<w:tab')) {
            [void]$nodes.Add([pscustomobject]@{ Index = $m.Index; Length = $m.Length; Text = "`t"; Editable = $false; Match = $m })
        } elseif ($v.StartsWith('<w:br')) {
            [void]$nodes.Add([pscustomobject]@{ Index = $m.Index; Length = $m.Length; Text = "`n"; Editable = $false; Match = $m })
        } else {
            # a symbol character: one glyph, contents unknown to us
            [void]$nodes.Add([pscustomobject]@{ Index = $m.Index; Length = $m.Length; Text = '?'; Editable = $false; Match = $m })
        }
    }
    return ,$nodes
}

function Get-Text {
    param([string]$Block)
    $sb = New-Object System.Text.StringBuilder
    foreach ($n in (Get-TextNodes -Block $Block)) { [void]$sb.Append($n.Text) }
    return $sb.ToString()
}

# Spans. NOTE the parentheses round the sum: @($a, $b + $c) is THREE elements,
# because in an array literal the comma binds tighter than +. Without them
# every containment test is false, which silently misclassified 1,017 table
# cells as body text once already.
function Get-Spans {
    param([string]$Xml, [string]$Pattern)
    $spans = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Xml, $Pattern)) { [void]$spans.Add(@($m.Index, ($m.Index + $m.Length))) }
    return ,$spans
}
function Test-InSpan {
    param($Spans, [int]$Index)
    foreach ($s in $Spans) { if ($Index -ge $s[0] -and $Index -lt $s[1]) { return $true } }
    return $false
}

function Set-PStyle {
    param([string]$Block, [string]$StyleId)
    $sm = [regex]::Match($Block, '<w:pStyle(?:\s[^>]*?)?\s*/>')
    if ($sm.Success) {
        return $Block.Remove($sm.Index, $sm.Length).Insert($sm.Index, "<w:pStyle w:val=`"$StyleId`"/>")
    }
    # CT_PPr order: pStyle is the first child of pPr
    $pp = [regex]::Match($Block, '(?s)^(<w:p\b[^>]*?>)(\s*<w:pPr\b[^>]*?>)')
    if ($pp.Success) { return $Block.Insert($pp.Groups[2].Index + $pp.Groups[2].Length, "<w:pStyle w:val=`"$StyleId`"/>") }
    $open = [regex]::Match($Block, '(?s)^<w:p\b[^>]*?>')
    return $Block.Insert($open.Length, "<w:pPr><w:pStyle w:val=`"$StyleId`"/></w:pPr>")
}

$capRx  = '^\s*(Figure|Table)\s*\u00a0?\s*(?:[A-Z]?\d+(?:[.\-\u2013]\d+)?|[A-Z][.\-\u2013]\d+)\s*[.:]\s+\S'
# What identifies a reference entry is the YEAR IN PARENTHESES near the front,
# not a personal author's name. Requiring "Surname, A." matched 64 of 87
# entries in one chapter and missed every organizational author there is:
# "ANSI. (2023).", "Federal Emergency Management Agency. (2025).",
# "LUMA. (2024a).", "MRLC (Multi-Resolution Land Characteristics Consortium).
# (2016).". The year may carry a letter for same-year works (2024a) or a date
# ("2018, September 21"). The alternative branch keeps the personal-author
# shape, for a style that puts the year elsewhere.
$refRx  = '^.{0,120}?\((?:n\.d\.|\d{4}[a-z]?(?:,[^)]{1,24})?)\)' +
          '|^[A-Z][A-Za-z\u00c0-\u024f''\- ]+,\s+[A-Z]\.'
$leadRx = '^(\s*(?:Figure|Table)\s*\u00a0?\s*[^\s]*?[.:])( *)'

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
    $stylesXml = if ($pkg.Data.ContainsKey('word/styles.xml')) { [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/styles.xml']) } else { '' }
    $numXml    = if ($pkg.Data.ContainsKey('word/numbering.xml')) { [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/numbering.xml']) } else { '' }

    $have = @{}
    foreach ($m in [regex]::Matches($stylesXml, 'w:styleId="([^"]+)"')) { $have[$m.Groups[1].Value] = $true }

    $beforeText = Get-Text -Block $raw

    $tbl  = Get-Spans -Xml $raw -Pattern '(?s)<w:tbl>.*?</w:tbl>'
    $txbx = Get-Spans -Xml $raw -Pattern '(?s)<w:txbxContent>.*?</w:txbxContent>'

    # numId -> ilvl0 numFmt, for choosing ListBullet vs ListNumber
    $num2abs = @{}
    foreach ($m in [regex]::Matches($numXml, '(?s)<w:num\b[^>]*w:numId="(\d+)".*?<w:abstractNumId w:val="(\d+)"')) {
        $num2abs[$m.Groups[1].Value] = $m.Groups[2].Value
    }
    $abs2fmt = @{}
    foreach ($m in [regex]::Matches($numXml, '(?s)<w:abstractNum\b[^>]*w:abstractNumId="(\d+)".*?</w:abstractNum>')) {
        $f = [regex]::Match($m.Value, '(?s)<w:lvl w:ilvl="0".*?<w:numFmt w:val="([^"]+)"')
        if ($f.Success) { $abs2fmt[$m.Groups[1].Value] = $f.Groups[1].Value }
    }

    # The References section, bounded at BOTH ends. The start is the last
    # "References" heading; the end is the next heading-styled paragraph after
    # it. The upper bound is not optional: Ch06 of NWIRP 2-1 carries its
    # mathematical nomenclature AFTER the reference list ("I = Set of nodes",
    # "Parameters", "L = Set of links"), and with no end the sweep would put
    # those on the References style.
    $refSectionStart = -1
    $refSectionEnd = [int]::MaxValue
    foreach ($pm in [regex]::Matches($raw, '(?s)<w:p\b[^>]*?>.*?</w:p>')) {
        if ((Get-Text -Block $pm.Value).Trim() -match '^(?i)references?$') { $refSectionStart = $pm.Index }
    }
    if ($refSectionStart -ge 0) {
        foreach ($pm in [regex]::Matches($raw, '(?s)<w:p\b[^>]*?>.*?</w:p>')) {
            if ($pm.Index -le $refSectionStart) { continue }
            if ($pm.Value -match '<w:pStyle\b[^>]*\bw:val="(Heading\d|AppendixHead|Title)"\s*/>') {
                $refSectionEnd = $pm.Index
                break
            }
        }
    }

    $tally = [ordered]@{}
    $report = [ordered]@{}
    $decide = [ordered]@{}
    function Add-Decide { param([string]$Key,[string]$Text)
        if (-not $decide.Contains($Key)) { $decide[$Key] = New-Object System.Collections.ArrayList }
        if ($decide[$Key].Count -lt 8 -and $Text) { [void]$decide[$Key].Add($Text.Substring(0,[Math]::Min(72,$Text.Length))) }
    }
    $sb = New-Object System.Text.StringBuilder
    $pos = 0

    foreach ($pm in [regex]::Matches($raw, '(?s)<w:p\b[^>]*?>.*?</w:p>')) {
        $p = $pm.Value
        $new = $p

        if (Test-InSpan -Spans $txbx -Index $pm.Index) {
            $report['shape text (left alone)'] = 1 + [int]$report['shape text (left alone)']
            continue
        }
        $inTbl = Test-InSpan -Spans $tbl -Index $pm.Index
        $sm = [regex]::Match($p, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        $sid = if ($sm.Success) { $sm.Groups[1].Value } else { $null }
        $t = (Get-Text -Block $p).Trim()
        $hasImg = $p -match '<a:blip\b|<w:pict\b|<w:drawing\b'

        # --- figures: a paragraph outside a table that carries an image ---
        # Including a PANEL LABEL - an image with a short line of text in the
        # same paragraph, as in a multi-panel figure ("a) Fully functional",
        # "Observed NDWE", "Trees flattened"). There were nine of those across
        # the set and they were reported rather than styled at first, on the
        # grounds that the certified set has none to copy. What settles it is
        # that the paragraph HOLDS THE IMAGE, which makes it figure content
        # whatever its text: graphic centres it, zeroes the spacing and carries
        # keepNext, so the panels stay in their grid and stay with the caption.
        # On BodyText each one took 6pt of space-after and broke the grid.
        #
        # A caption paragraph that holds its own image is left on its caption
        # style. Two figures in Ch05 are built that way, and the caption is the
        # more specific of the two descriptions.
        if ($rules -contains 'Graphic' -and -not $inTbl -and $hasImg -and
            $have.ContainsKey('graphic') -and $sid -notmatch '^Caption' -and
            ($t.Length -eq 0 -or $t.Length -lt 40)) {
            if ($sid -ne 'graphic') {
                $new = Set-PStyle -Block $p -StyleId 'graphic'
                $key = if ($t) { '-> graphic (panel label)' } else { '-> graphic' }
                $tally[$key] = 1 + [int]$tally[$key]
            }
        }
        # --- a long line sharing a paragraph with an image: report only ---
        elseif (-not $inTbl -and $hasImg -and $t -and $sid -notmatch '^Caption') {
            $report['image inside a text paragraph (DECIDE)'] = 1 + [int]$report['image inside a text paragraph (DECIDE)']
            Add-Decide -Key 'image inside a text paragraph (DECIDE)' -Text $t
        }
        # --- captions ---
        # Either caption-SHAPED (label + number + period + title) or already on
        # one of Word's caption styles. The second test matters: a caption whose
        # author left the period out does not match the shape, so judging on
        # shape alone left NWIRP 2-1 Ch03's "Figure 31 Study regions" on the
        # built-in Caption style while still normalising its punctuation.
        elseif ($rules -contains 'Caption' -and
                ($t -match $capRx -or
                 ($sid -match '^Caption' -and $t -match '^\s*(Figure|Table)\b'))) {
            $label = $Matches[1]
            $want = if ($label -eq 'Figure') { 'CaptionFigure' } else { 'CaptionTable' }
            if ($have.ContainsKey($want) -and $sid -ne $want) {
                $new = Set-PStyle -Block $p -StyleId $want
                $tally["-> $want"] = 1 + [int]$tally["-> $want"]
            }
        }
        # --- block quotations ---
        # An extract quoted from a source, set as its own paragraph: wholly
        # italic and opening with a quotation mark. That pair is the signature -
        # body prose is not italic end to end, and a run of italic inside a
        # sentence (a journal title, a species name, a variable) does not start
        # the paragraph with a quote mark. Seven in one chapter, 300 to 790
        # characters each.
        #
        # The italic is NOT stripped. Reapply-Styles.ps1 preserves italic
        # deliberately, and if the Quote style carries it the direct run is
        # merely redundant rather than wrong.
        elseif ($rules -contains 'Quote' -and -not $inTbl -and $t.Length -ge 60 -and
                $have.ContainsKey('Quote') -and ($sid -eq 'BodyText' -or $sid -eq 'Quote' -or $null -eq $sid) -and
                $t -match '^[\u201c\u2018"]') {
            $italicChars = 0; $allChars = 0
            foreach ($rm in [regex]::Matches($p, '(?s)<w:r\b[^>]*?>.*?</w:r>')) {
                $rt = Get-Text -Block $rm.Value
                if (-not $rt) { continue }
                $allChars += $rt.Length
                if ($rm.Value -match '<w:i\b(?:\s[^/>]*)?/>') { $italicChars += $rt.Length }
            }
            if ($allChars -gt 0 -and ($italicChars / $allChars) -ge 0.8) {
                if ($sid -ne 'Quote') {
                    $new = Set-PStyle -Block $p -StyleId 'Quote'
                    $tally['-> Quote'] = 1 + [int]$tally['-> Quote']
                }
            } else {
                Add-Decide -Key 'quoted but not italic throughout (DECIDE)' -Text $t
            }
        }
        # --- reference entries, only inside the References section ---
        elseif ($rules -contains 'Reference' -and -not $inTbl -and $refSectionStart -ge 0 -and
                $pm.Index -gt $refSectionStart -and $pm.Index -lt $refSectionEnd -and
                $t.Length -ge 30 -and $t -match $refRx -and $have.ContainsKey('References')) {
            if ($sid -ne 'References') { $new = Set-PStyle -Block $p -StyleId 'References'; $tally['-> References'] = 1 + [int]$tally['-> References'] }
        }
        # --- list items: a numPr on an unstyled or ListParagraph paragraph ---
        elseif ($rules -contains 'List' -and -not $inTbl -and ($null -eq $sid -or $sid -eq 'Normal' -or $sid -eq 'ListParagraph')) {
            $npr = [regex]::Match($p, '(?s)<w:numPr>.*?</w:numPr>')
            if ($npr.Success) {
                $nid = [regex]::Match($npr.Value, '<w:numId w:val="(\d+)"')
                $ilv = [regex]::Match($npr.Value, '<w:ilvl w:val="(\d+)"')
                if ($nid.Success -and $nid.Groups[1].Value -ne '0') {
                    $fmt = $null
                    if ($num2abs.ContainsKey($nid.Groups[1].Value)) { $fmt = $abs2fmt[$num2abs[$nid.Groups[1].Value]] }
                    $fam = switch ($fmt) {
                        'bullet' { 'ListBullet' }
                        { $_ -in @('decimal','lowerLetter','upperLetter','lowerRoman','upperRoman') } { 'ListNumber' }
                        default { $null }
                    }
                    if ($fam) {
                        $lvl = if ($ilv.Success) { [int]$ilv.Groups[1].Value } else { 0 }
                        $want = if ($lvl -eq 0) { $fam } else { "$fam$([Math]::Min($lvl+1,5))" }
                        if ($have.ContainsKey($want)) {
                            # the List style's own numId supplies the glyph, so the
                            # paragraph's direct numPr comes off - the certified set's
                            # ListBullet paragraphs carry none
                            $new = Set-PStyle -Block $p -StyleId $want
                            $new = $new.Replace($npr.Value, '')
                            $tally["-> $want"] = 1 + [int]$tally["-> $want"]
                        } else { $report["$want absent"] = 1 + [int]$report["$want absent"] }
                    } else { $report["list, numFmt $fmt (DECIDE)"] = 1 + [int]$report["list, numFmt $fmt (DECIDE)"] }
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

    # --- caption prefix: label + NBSP + number + period + TWO spaces + title ---
    # The period is part of the form, so a caption missing it is normalised too
    # ("Table 21 Data sources" -> "Table 21.  Data sources"). A caption whose
    # number is a SEQ field is skipped: the digits in a field are only Word's
    # cached result and rewriting them corrupts nothing but achieves nothing.
    if ($rules -contains 'CaptionSpacing') {
        $fixed = 0
        $sb2 = New-Object System.Text.StringBuilder
        $pos2 = 0
        foreach ($pm in [regex]::Matches($newDoc, '(?s)<w:p\b[^>]*?>.*?</w:p>')) {
            $p = $pm.Value
            if ($p -notmatch '<w:pStyle\b[^>]*\bw:val="Caption\w*"\s*/>') { continue }
            $t = Get-Text -Block $p
            if (-not $t.Trim()) { continue }

            # A caption whose number is a SEQ field keeps the field: only the
            # punctuation AFTER the last field end is normalised, so nothing
            # inside the field is touched and a field update still renumbers it.
            # The chapter prefix (STYLEREF) and the number live in the field, so
            # the label and separator are the author's, not ours, to rewrite.
            # Where the caption's NUMBER ends: the text offset after the last
            # field in the leading region. Both field forms count - a complex
            # field closes with fldCharType="end", a SIMPLE field closes with
            # </w:fldSimple> - and the caption that broke this used one of
            # each, STYLEREF for the chapter and SEQ for the number. Measured
            # on the full text model, so the non-breaking hyphen between them
            # is counted.
            $fieldEnd = -1
            if ($p -match '<w:fldChar w:fldCharType="end"\s*/>|</w:fldSimple>') {
                # A field inside a TRACKED DELETION is not part of the number.
                # Sean's captions carry a deleted SEQ field sitting after the
                # ".  " that already terminates the number, and its fldChar end
                # was being counted: $fieldEnd moved PAST the punctuation, the
                # rule then found no terminator in front of the title, and
                # inserted a second one - "Figure 2-10.  .  Maps showing". In
                # one caption the deleted field sat mid-title and the period
                # landed in the prose: "Same as Figure 3-3.  , but now also".
                # Deleted content contributes no text, so it must not move a
                # text offset either.
                $delSpans = New-Object System.Collections.ArrayList
                foreach ($dm in [regex]::Matches($p, '(?s)<w:del\b[^>]*?>.*?</w:del>')) {
                    [void]$delSpans.Add([pscustomobject]@{ Start = $dm.Index; End = $dm.Index + $dm.Length })
                }
                $marks = @{}
                foreach ($em in [regex]::Matches($p, '<w:fldChar w:fldCharType="end"\s*/>|</w:fldSimple>')) {
                    $inDel = $false
                    foreach ($d in $delSpans) {
                        if ($em.Index -ge $d.Start -and $em.Index -lt $d.End) { $inDel = $true; break }
                    }
                    if (-not $inDel) { $marks[$em.Index] = $true }
                }
                # The window is the number region, not a fixed 40 characters.
                # A caption's number ENDS at its first terminator; anything past
                # that is title. Ch03 has "Figure 3-10.  Same as Figure 3-3, but
                # now also showing ...", where "Figure 3-3" is an author's REF
                # field whose end sits 32 characters in - inside a 40-character
                # window - so it was taken for the end of the number and the
                # terminator was inserted there: "Same as Figure 3-3.  , but now
                # also showing". A cross-reference near the start of a title is
                # ordinary, so the window has to be bounded by meaning rather
                # than by a character count.
                $limit = 40
                $tm = [regex]::Match($t, '[.:](?=[\s\u00a0]|$)')
                if ($tm.Success) { $limit = $tm.Index + 1 }
                $acc = 0
                $lastEnd = 0
                foreach ($n in (Get-TextNodes -Block $p)) {
                    foreach ($k in @($marks.Keys | Where-Object { $_ -ge $lastEnd -and $_ -lt $n.Index })) {
                        if ($acc -le $limit) { $fieldEnd = $acc }
                    }
                    $acc += $n.Text.Length
                    $lastEnd = $n.Index + $n.Length
                }
                foreach ($k in @($marks.Keys | Where-Object { $_ -ge $lastEnd })) {
                    if ($acc -le $limit) { $fieldEnd = $acc }
                }
            }

            if ($fieldEnd -ge 0) {
                # The number may CONTINUE as literal text after the field. A
                # caption numbered by STYLEREF alone reads
                # "Figure " + <STYLEREF>2 + <noBreakHyphen> + "9. Maps showing",
                # so the field ends at "Figure 2" and "-9" is ordinary text.
                # Treating the field end as the end of the number put the
                # period inside it: "Figure 2-.  9.".
                #
                # Absorb only a separator-plus-digits run, which is why a title
                # beginning with a year survives: in "Figure 5.  2020 Census"
                # the separator is followed by spaces, not digits, so nothing
                # is absorbed.
                $cont = [regex]::Match($t.Substring($fieldEnd), '^(?:[\u2010\u2011\u2013.\-]?\d+)+')
                $numEnd = $fieldEnd + $cont.Length
                # A colon is the author's terminator in some chapters
                # ("Figure P-1: Study Regions"); the NIST form is a period, and
                # a pattern that only allowed "\.?" left the colon in place and
                # produced "Figure P-1.  : Study Regions".
                # TWO terminators, not one. The authors' own form is sometimes a
                # colon ("Figure P-1: Study Regions"), and a pattern allowing a
                # single terminator turned that into "Figure P-1.  : Study
                # Regions" - the period went in and the colon stayed. Allowing
                # a second one collapses ".  :" back to a single ".  ", and
                # makes the pass idempotent on anything it has already touched.
                $tail = [regex]::Match($t.Substring($numEnd), '^([\s\u00a0]*[.:]?[\s\u00a0]*[.:]?[\s\u00a0]*)')
                $old = $tail.Groups[1].Value
                $want = '.'
                if ($t.Substring($numEnd + $old.Length).Trim()) { $want = '.  ' }
                $start = $numEnd
                $end = $numEnd + $old.Length
                if ($old -ceq $want) { continue }
            } else {
                # label, separator, number, whatever punctuation and space follow
                $lead = [regex]::Match($t, '^(\s*)(Figure|Table)([\s\u00a0]*)((?:[A-Z][.\-\u2010\u2011\u2013])?\d+(?:[.\-\u2010\u2011\u2013]\d+)?)([\s\u00a0]*[.:]?[\s\u00a0]*[.:]?[\s\u00a0]*)')
                if (-not $lead.Success) {
                    Add-Decide -Key 'caption with no recognisable number (DECIDE)' -Text $t.Trim()
                    continue
                }
                $rest = $t.Substring($lead.Index + $lead.Length)
                $want = $lead.Groups[2].Value + [char]0x00a0 + $lead.Groups[4].Value + '.'
                if ($rest.Trim()) { $want = $want + '  ' }
                $start = $lead.Groups[1].Value.Length
                $end   = $lead.Index + $lead.Length
                $old   = $t.Substring($start, $end - $start)
                if ($old -ceq $want) { continue }
            }

            # Replace the text range [$start,$end) across the paragraph's text
            # nodes, on the same model the offsets were measured with. A range
            # that touches a node this script may not edit - a hyphen, a tab, a
            # symbol - is left alone and reported rather than guessed at.
            $nodes = Get-TextNodes -Block $p
            $blocked = $false
            $cur = 0
            foreach ($n in $nodes) {
                $lo = $cur; $hi = $cur + $n.Text.Length; $cur = $hi
                if (-not $n.Editable -and $lo -lt $end -and $hi -gt $start) { $blocked = $true }
            }
            if ($blocked) {
                Add-Decide -Key 'caption prefix spans a hyphen or tab (DECIDE)' -Text $t.Trim()
                continue
            }

            $cur = 0; $placed = $false; $out = New-Object System.Text.StringBuilder; $last = 0
            foreach ($n in $nodes) {
                $lo = $cur; $hi = $cur + $n.Text.Length; $cur = $hi
                if (-not $n.Editable) { continue }
                $tm = $n.Match
                $txt = $n.Text
                $s = [Math]::Max($lo, $start); $e = [Math]::Min($hi, $end)
                $nt = $txt
                if ($s -lt $e) { $nt = $txt.Substring(0, $s - $lo) + $txt.Substring($e - $lo) }
                # STRICTLY inside. At a node boundary, $start -le $hi is also true
                # of the PRECEDING node, which for a field caption is the cached
                # result inside the field: inserting there writes into the field.
                if (-not $placed -and $start -ge $lo -and ($start -lt $hi -or ($start -eq $hi -and $hi -eq $lo))) {
                    $k = $start - $lo
                    $nt = $nt.Substring(0, $k) + $want + $nt.Substring($k)
                    $placed = $true
                }
                if ($nt -ne $txt) {
                    $attrs = [regex]::Match($tm.Value, '^<w:t((?:\s[^>]*)?)>').Groups[1].Value
                    if ($attrs -notmatch 'xml:space') { $attrs = $attrs + ' xml:space="preserve"' }
                    [void]$out.Append($p.Substring($last, $tm.Index - $last))
                    [void]$out.Append('<w:t' + $attrs + '>' + $nt + '</w:t>')
                    $last = $tm.Index + $tm.Length
                }
            }
            if (-not $placed) { continue }
            [void]$out.Append($p.Substring($last))
            [void]$sb2.Append($newDoc.Substring($pos2, $pm.Index - $pos2))
            [void]$sb2.Append($out.ToString())
            $pos2 = $pm.Index + $pm.Length
            $fixed++
        }
        if ($fixed -gt 0) {
            [void]$sb2.Append($newDoc.Substring($pos2))
            $newDoc = $sb2.ToString()
            $tally['caption prefix normalised'] = $fixed
        }
    }

    foreach ($k in $report.Keys) {
        Write-Host ("  {0,-48} {1}" -f $k, $report[$k]) -ForegroundColor DarkYellow
        if ($decide.Contains($k)) { foreach ($l in $decide[$k]) { Write-Host ("      | {0}" -f $l) -ForegroundColor DarkGray } }
    }
    foreach ($k in $decide.Keys) {
        if ($report.Contains($k)) { continue }
        Write-Host ("  {0,-48} {1}" -f $k, $decide[$k].Count) -ForegroundColor DarkYellow
        foreach ($l in $decide[$k]) { Write-Host ("      | {0}" -f $l) -ForegroundColor DarkGray }
    }
    if ($tally.Count -eq 0) { Write-Host "  Nothing to change." -ForegroundColor Green; continue }
    foreach ($k in $tally.Keys) { Write-Host ("  {0,-48} {1}" -f $k, $tally[$k]) -ForegroundColor DarkGray }

    # guards: only caption spaces may change the text, and only by adding spaces
    # Guard. Only whitespace and the caption period may differ, so compare with
    # both removed: that proves no word, digit or symbol was touched. Then check
    # the period count moved by no more than the number of captions normalised,
    # which catches a period deleted somewhere it was carrying meaning.
    $afterText = Get-Text -Block $newDoc
    # The caption terminator may be SUBSTITUTED, not just added: an author's
    # colon becomes the house period. So the comparison ignores whitespace and
    # both terminators, and the counts of each are then bounded by how many
    # captions were touched - neither test is enough alone.
    $norm = { param($x) ([regex]::Replace($x, '[\s\u00a0.:]+', '')) }
    if ((& $norm $afterText) -ne (& $norm $beforeText)) {
        throw "Rewrite changed text other than whitespace and caption punctuation in ${name}; nothing written."
    }
    $touched = [int]$tally['caption prefix normalised']
    $dots = ([regex]::Matches($afterText, '\.').Count - [regex]::Matches($beforeText, '\.').Count)
    if ([Math]::Abs($dots) -gt $touched) {
        throw "Rewrite changed the period count by $dots in ${name}, more than the $touched caption(s) normalised; nothing written."
    }
    $colons = ([regex]::Matches($beforeText, ':').Count - [regex]::Matches($afterText, ':').Count)
    if ($colons -lt 0 -or $colons -gt $touched) {
        throw "Rewrite changed the colon count by $colons in ${name}, which the $touched caption(s) normalised do not account for; nothing written."
    }
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Rewrite produced malformed document.xml for ${name}: $($_.Exception.Message)" }

    if ($ListOnly) { continue }
    if (-not $NoBackup) {
        $bak = "$target.content.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }

    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "content-$([guid]::NewGuid()).zip")
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

    try { Move-Item -LiteralPath $tempZip -Destination $target -Force; $filesChanged++; Write-Host "  Saved." -ForegroundColor Green }
    catch { Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue; Write-Host "  LOCKED (open in Word) - not written." -ForegroundColor Red }
}

Write-Host ""
Write-Host ("Done. {0} file(s) changed." -f $filesChanged) -ForegroundColor Green
