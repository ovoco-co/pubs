<#
.SYNOPSIS
    Turns typed table references in the body text ("see Table 5-10") into live
    Word cross-reference fields pointing at the caption they name.

.DESCRIPTION
    A chapter written by hand usually carries its table references as plain
    text. They do not move when the tables are renumbered, and Word's
    Cross-reference dialog cannot adopt them: it only inserts new ones.

    This finds the table captions, bookmarks the label and number on each one,
    and replaces every matching mention in the body with

        { REF _Ref123456789 \h }

    which is exactly what Word's "Insert cross-reference, only label and
    number" writes. The caption's current text is cached between the separate
    and end marks, so the document reads correctly before Word has updated
    anything, and Ctrl+A then F9 refreshes it.

    Matching is on the number the caption currently shows, so the references
    say what the text says today. Renumber the captions afterwards, press F9,
    and every reference follows, because a REF field reports what its bookmark
    contains rather than what it said when it was written.

    The bookmark opens immediately after <w:pPr>, ahead of the number runs, so
    that Set-NISTTableCaptions.ps1 (which inserts its numbering runs at exactly
    that point) writes a new number INSIDE the bookmark rather than in front of
    it. Running these two scripts in either order therefore still leaves every
    reference pointing at a live number.

    A caption is a paragraph on -CaptionStyle whose text starts with the label
    and a number ("Table 5-10. ..."). A mention is that same pattern anywhere
    else. A mention already inside a field is left alone, so re-running is
    safe.

.PARAMETER File
    A single .docx to process.

.PARAMETER Folder
    Process every .docx in this folder.

.PARAMETER Filter
    Wildcard used with -Folder. Defaults to *.docx.

.PARAMETER CaptionStyle
    Paragraph styles that carry a table caption.
    Defaults to CaptionTable, CaptionFigure.

.PARAMETER Label
    The caption label. Defaults to Table.

.PARAMETER ListOnly
    Report the caption-to-mention mapping and write nothing. Run this first.

.PARAMETER Force
    Skip the confirmation prompt.

.PARAMETER NoBackup
    Skip the .bak copy and write the file in place.

.EXAMPLE
    .\Set-TableCrossReferences.ps1 -File "Chapter-04.docx" -ListOnly

.EXAMPLE
    .\Set-TableCrossReferences.ps1 -File "Chapter-04.docx" -Force
#>

param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$Filter = '*.docx',

    [string[]]$CaptionStyle = @('CaptionTable', 'CaptionFigure'),

    [string]$Label = 'Table',

    # Also treat an unstyled paragraph as a caption when its text begins with
    # the label, a number and a period ("Table 5-43.  Correlation Analysis").
    # The period is what separates a caption from a sentence that opens with
    # "Table 5-33 presents ...", so it is required.
    [switch]$NoUnstyledCaptions,

    # Only link a mention whose chapter number also matches. Off by default: a
    # renumbered chapter leaves the text saying 5-9 and the caption saying 1-9,
    # and the table number alone identifies it when it is unique.
    [switch]$ExactNumbersOnly,

    [switch]$ListOnly,

    [switch]$Force,

    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$styleAlt  = ($CaptionStyle | ForEach-Object { [regex]::Escape($_) }) -join '|'
$styleRx   = '<w:pStyle\s[^>]*w:val="(?:' + $styleAlt + ')"[^>]*/>'

# "Table" then a space or non-breaking space, chapter, any dash, number.
# The separators are written as regex escapes, not as literal characters:
# Windows PowerShell 5.1 reads a .ps1 with no BOM as ANSI, so a literal
# NBSP or en dash in the source is mis-decoded and the script fails to
# parse at all. \u00A0 NBSP, \u2011 non-breaking hyphen, \u2013 en dash,
# \u2014 em dash. Keeping this file pure ASCII makes the encoding moot.
# The separator includes a DOT. Ch05 and Ch06 were authored with
# "Figure 6.1" in the prose, and the captions now read "Figure 6-1" with an
# en dash, so without the dot the mention never matched its own caption and
# 53 cross-references stayed as dead text. Both forms parse to the same
# chapter/number pair, which is what the lookup key is built from.
$numberRx  = [regex]::Escape($Label) + '[\s\u00A0]*(\d+)\s*[.\-\u2010\u2011\u2013\u2014]\s*(\d+)'
$runRx     = '(?s)<w:r(?:\s[^>]*)?>.*?</w:r>'
$textRx    = '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-RunText {
    param([string]$Run)
    return (-join ([regex]::Matches($Run, $textRx) | ForEach-Object { $_.Groups[1].Value }))
}

function Split-Run {
    # Cut one run in two at a character offset into its text, keeping the run
    # properties on both halves. Returns the two run strings.
    param([string]$Run, [int]$Offset)

    $rPr = ''
    $rPrMatch = [regex]::Match($Run, '(?s)^<w:r(?:\s[^>]*)?>\s*(<w:rPr>.*?</w:rPr>)')
    if ($rPrMatch.Success) { $rPr = $rPrMatch.Groups[1].Value }

    $text = Get-RunText -Run $Run
    $head = $text.Substring(0, $Offset)
    $tail = $text.Substring($Offset)

    $open = [regex]::Match($Run, '^<w:r(?:\s[^>]*)?>').Value
    $mk = {
        param($t)
        if ($t.Length -eq 0) { return '' }
        return $open + $rPr + '<w:t xml:space="preserve">' + $t + '</w:t></w:r>'
    }
    return @((& $mk $head), (& $mk $tail))
}

function New-RefField {
    # The five runs Word writes for a cross-reference, with the caption's
    # current text cached so the document reads correctly before F9.
    param([string]$Bookmark, [string]$Cached, [string]$RPr)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<w:r>$RPr<w:fldChar w:fldCharType=`"begin`"/></w:r>")
    [void]$sb.Append("<w:r>$RPr<w:instrText xml:space=`"preserve`"> REF $Bookmark \h </w:instrText></w:r>")
    [void]$sb.Append("<w:r>$RPr<w:fldChar w:fldCharType=`"separate`"/></w:r>")
    [void]$sb.Append("<w:r>$RPr<w:t>$Cached</w:t></w:r>")
    [void]$sb.Append("<w:r>$RPr<w:fldChar w:fldCharType=`"end`"/></w:r>")
    return $sb.ToString()
}

function Test-SpanReplaceable {
    # A mention is only safe to swap for a field when the run span it covers is
    # a whole number of sibling elements. This document set carries tracked
    # changes, and a span that starts inside <w:ins> and ends outside it would
    # take the closing tag with it and leave the XML malformed.
    #
    # Deleted text is skipped outright: a reference inside a <w:del> is text the
    # author has already removed, and turning it into a live field revives it.
    param([string]$Para, [int]$Start, [int]$Length)

    $span = $Para.Substring($Start, $Length)
    # A tracked-change boundary anywhere inside the mention: leave it typed.
    # Balanced is not enough here - "Figure " in one insertion and "4-10" in
    # the next balance, but the rewrite would drop both insertions' markers and
    # turn an author's tracked insertion into untracked text. A mention wholly
    # inside one insertion has no boundary in its span and is still linked.
    if ($span -match '<w:(ins|del|moveFrom|moveTo)\b|</w:(ins|del|moveFrom|moveTo)>') { return $false }
    foreach ($tag in @('w:ins', 'w:del', 'w:hyperlink', 'w:smartTag', 'w:sdt', 'w:fldSimple')) {
        $open  = ([regex]::Matches($span, "<$tag(?:\s[^>]*)?>")).Count
        $close = ([regex]::Matches($span, "</$tag>")).Count
        if ($open -ne $close) { return $false }
    }

    $before = $Para.Substring(0, $Start)
    $delOpen  = ([regex]::Matches($before, '<w:del(?:\s[^>]*)?>')).Count
    $delClose = ([regex]::Matches($before, '</w:del>')).Count
    if ($delOpen -gt $delClose) { return $false }

    # A mention already inside a complex field is the CACHED RESULT of a field
    # written earlier - a REF put there by a previous pass reads "Table 3-1" in
    # its result runs, and that text is part of the joined run text a mention is
    # matched against. Wrapping it again would nest a field inside its own
    # result. Tested per span rather than per paragraph: the Table pass leaves a
    # REF in a paragraph that may still hold an unlinked Figure mention, and a
    # paragraph-wide skip abandoned that mention. Balanced counts mean the span
    # sits wholly outside every field.
    $fbBefore = ([regex]::Matches($before, 'w:fldCharType="begin"')).Count
    $feBefore = ([regex]::Matches($before, 'w:fldCharType="end"')).Count
    if ($fbBefore -gt $feBefore) { return $false }
    $fbSpan = ([regex]::Matches($span, 'w:fldCharType="begin"')).Count
    $feSpan = ([regex]::Matches($span, 'w:fldCharType="end"')).Count
    if ($fbSpan -ne $feSpan) { return $false }

    return $true
}

function Get-FirstRPr {
    param([string]$Run)
    $m = [regex]::Match($Run, '(?s)^<w:r(?:\s[^>]*)?>\s*(<w:rPr>.*?</w:rPr>)')
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

# ---------------------------------------------------------------------------
# Target list
# ---------------------------------------------------------------------------

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets = @((Resolve-Path -LiteralPath $File).ProviderPath)
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    # Exclude Word's ~$ lock files: they match *.docx and are not documents.
    $targets = @(Get-ChildItem -LiteralPath $Folder -Filter $Filter -File |
        Where-Object { $_.Name -notlike '~$*' } |
                 Where-Object { $_.Name -notlike '~$*' } |
                 ForEach-Object { $_.FullName })
    if ($targets.Count -eq 0) { throw "No files matching $Filter in $Folder" }
}

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


Write-Host "Set Table Cross-References" -ForegroundColor Cyan
Write-Host "  Replaces typed '$Label n-n' mentions with REF fields on the caption." -ForegroundColor Gray
Write-Host "  Targets:" -ForegroundColor Gray
foreach ($t in $targets) { Write-Host "    $t" -ForegroundColor Gray }

if ($ListOnly) {
    Write-Host ""
    Write-Host "  -ListOnly: reporting matches, writing nothing." -ForegroundColor Yellow
} elseif (-not $Force) {
    Write-Host ""
    $reply = Read-Host "Rewrite mentions as cross-reference fields? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $zip = [System.IO.Compression.ZipFile]::Open($target, 'Read')
    $entries = @{}
    foreach ($e in $zip.Entries) {
        $ms = New-Object System.IO.MemoryStream
        $s = $e.Open(); $s.CopyTo($ms); $s.Close()
        $entries[$e.FullName] = $ms.ToArray()
    }
    $zip.Dispose()

    if (-not $entries.ContainsKey('word/document.xml')) { throw "No word/document.xml in $target" }
    $xml = $utf8NoBom.GetString($entries['word/document.xml'])
    if ($xml.Length -gt 0 -and $xml[0] -eq [char]0xFEFF) { $xml = $xml.Substring(1) }

    # ---- pass 1: find the captions and give each one a bookmark -----------
    $paras = [regex]::Matches($xml, '(?s)<w:p(?:\s[^>]*)?>.*?</w:p>')

    # A complex field can span MANY paragraphs. A TOC is exactly that: one
    # fldChar begin, an instrText "TOC \o ...", a separate, then every entry
    # paragraph Word generated, then one fldChar end at the very bottom. Inside
    # an entry paragraph there are no unbalanced fldChars of its own, so a
    # paragraph-local balance test sees nothing wrong and the paragraph looks
    # editable. It is not: an entry is field RESULT, Word rewrites it on every
    # F9, and a REF written there is both lost and, until then, a field nested
    # inside its own result. An earlier run put 20 of 43 REFs inside the TOC
    # that way (16 as TOC > REF, 4 as TOC > HYPERLINK > REF).
    #
    # Depth is therefore carried at DOCUMENT level: count begin/end markers
    # ahead of each paragraph and skip any paragraph that opens while a field
    # is still open. This is also the real reason TOC entries were matching as
    # captions in pass 1 - an entry carries the caption text verbatim - so the
    # gate is applied to both passes and -NoUnstyledCaptions is no longer the
    # only thing standing between a ToC and a bookmark.
    $fieldParas = @{}
    $markers = [regex]::Matches($xml, 'w:fldCharType="(begin|end)"')
    $mi = 0
    $depth = 0
    for ($i = 0; $i -lt $paras.Count; $i++) {
        while ($mi -lt $markers.Count -and $markers[$mi].Index -lt $paras[$i].Index) {
            if ($markers[$mi].Groups[1].Value -eq 'begin') { $depth++ } else { $depth-- }
            if ($depth -lt 0) { $depth = 0 }
            $mi++
        }
        if ($depth -gt 0) { $fieldParas[$i] = $true }
    }
    if ($fieldParas.Count -gt 0) {
        Write-Host "  Paragraphs inside a multi-paragraph field (TOC etc), skipped: $($fieldParas.Count)" -ForegroundColor Gray
    }

    # Bookmark ids have to be unique across the document.
    $maxId = 0
    foreach ($m in [regex]::Matches($xml, '<w:bookmarkStart\s[^>]*w:id="(\d+)"')) {
        $v = [int]$m.Groups[1].Value
        if ($v -gt $maxId) { $maxId = $v }
    }
    $nextId = $maxId + 1
    $refSeed = 540000000

    $captions = @{}      # "5-10" -> @{ Bookmark; Cached; Title }
    $byNumber = @{}      # "10"   -> the same entry, for a renumbered chapter
    $captionParas = @{}  # paragraph index -> $true, so pass 2 skips them
    $edits = @()         # paragraph rewrites, applied back to front

    for ($i = 0; $i -lt $paras.Count; $i++) {
        if ($fieldParas.ContainsKey($i)) { continue }               # TOC entry
        $para = $paras[$i].Value
        $styled = [regex]::IsMatch($para, $styleRx)
        if (-not $styled -and $NoUnstyledCaptions) { continue }

        $runs = [regex]::Matches($para, $runRx)
        if ($runs.Count -eq 0) { continue }
        $joined = -join ($runs | ForEach-Object { Get-RunText -Run $_.Value })

        # A styled caption only has to start with the number. An unstyled one
        # has to carry the period too, or every sentence opening "Table 5-33
        # presents ..." would be mistaken for a caption.
        $pattern = if ($styled) { '^\s*' + $numberRx } else { '^\s*' + $numberRx + '\s*[.)]' }
        $hit = [regex]::Match($joined, $pattern)
        if (-not $hit.Success) { continue }

        $key = "$($hit.Groups[1].Value)-$($hit.Groups[2].Value)"
        # The bookmark covers the label and number only, never the period.
        $numberOnly = [regex]::Match($joined, '^\s*' + $numberRx)
        if ($captions.ContainsKey($key)) {
            Write-Warning "  Duplicate caption number $Label $key; the first one keeps the bookmark."
            continue
        }

        $bookmark = "_Ref$($refSeed + $i)"
        $spanEnd  = $numberOnly.Index + $numberOnly.Length   # end of "Table 5-10"

        # Walk the runs to find where the label and number stop.
        $acc = 0; $endRun = -1; $endOffset = 0
        for ($r = 0; $r -lt $runs.Count; $r++) {
            $len = (Get-RunText -Run $runs[$r].Value).Length
            if ($acc + $len -ge $spanEnd) { $endRun = $r; $endOffset = $spanEnd - $acc; break }
            $acc += $len
        }
        if ($endRun -lt 0) { continue }

        $newPara = $para
        # bookmarkEnd first (later in the string), then bookmarkStart, so the
        # offsets taken from the original paragraph stay valid.
        $endRunMatch = $runs[$endRun]
        $parts = Split-Run -Run $endRunMatch.Value -Offset $endOffset
        $replacement = $parts[0] + "<w:bookmarkEnd w:id=`"$nextId`"/>" + $parts[1]
        $newPara = $newPara.Remove($endRunMatch.Index, $endRunMatch.Length).Insert($endRunMatch.Index, $replacement)

        # Open the bookmark immediately after w:pPr so a later renumber writes
        # its new runs inside it.
        $pPr = [regex]::Match($newPara, '(?s)<w:pPr(?:\s[^>]*)?>.*?</w:pPr>')
        $insertAt = if ($pPr.Success) { $pPr.Index + $pPr.Length } else { [regex]::Match($newPara, '^<w:p(?:\s[^>]*)?>').Length }
        $newPara = $newPara.Insert($insertAt, "<w:bookmarkStart w:id=`"$nextId`" w:name=`"$bookmark`"/>")

        $captions[$key] = @{
            Bookmark = $bookmark
            Cached   = $joined.Substring($numberOnly.Index, $numberOnly.Length).Trim()
            Title    = $joined.Trim()
        }
        $captionParas[$i] = $true
        if (-not $byNumber.ContainsKey($hit.Groups[2].Value)) {
            $byNumber[$hit.Groups[2].Value] = $captions[$key]
        } else {
            $byNumber[$hit.Groups[2].Value] = $null      # ambiguous, do not guess
        }
        $edits += @{ Index = $paras[$i].Index; Length = $paras[$i].Length; Text = $newPara }
        $nextId++
    }

    Write-Host "  Captions bookmarked: $($captions.Count)" -ForegroundColor Gray

    # ---- pass 2: rewrite the mentions -------------------------------------
    $mentionCount = 0
    $renumbered = 0
    $skipped = 0
    $unmatched = @()

    for ($i = 0; $i -lt $paras.Count; $i++) {
        $para = $paras[$i].Value
        if ($captionParas.ContainsKey($i)) { continue }              # captions
        if ($fieldParas.ContainsKey($i)) { continue }                # TOC entry
        if ($para -notmatch [regex]::Escape($Label)) { continue }
        # No paragraph-level skip for an existing REF: Test-SpanReplaceable
        # rejects a mention that is inside a field, so the other mentions in the
        # same paragraph are still linked. Skipping the paragraph left a typed
        # "Figure 3-1" unlinked purely because a Table REF shared the paragraph.

        $runs = [regex]::Matches($para, $runRx)
        if ($runs.Count -eq 0) { continue }
        $joined = -join ($runs | ForEach-Object { Get-RunText -Run $_.Value })

        $hits = [regex]::Matches($joined, $numberRx)
        if ($hits.Count -eq 0) { continue }

        $newPara = $para
        # Back to front, so each rewrite leaves the earlier offsets untouched.
        for ($h = $hits.Count - 1; $h -ge 0; $h--) {
            $hit = $hits[$h]
            $key = "$($hit.Groups[1].Value)-$($hit.Groups[2].Value)"
            $caption = $null
            if ($captions.ContainsKey($key)) {
                $caption = $captions[$key]
            } elseif (-not $ExactNumbersOnly -and $byNumber.ContainsKey($hit.Groups[2].Value) -and $byNumber[$hit.Groups[2].Value]) {
                # The chapter number in the text and in the caption disagree,
                # which is what a renumbered chapter looks like. The table
                # number is unique, so it decides.
                $caption = $byNumber[$hit.Groups[2].Value]
                $renumbered++
            }
            if (-not $caption) {
                if ($unmatched -notcontains $key) { $unmatched += $key }
                continue
            }

            $start = $hit.Index
            $end   = $hit.Index + $hit.Length

            # Find the run span the mention covers.
            $acc = 0; $startRun = -1; $startOff = 0; $endRun = -1; $endOff = 0
            $runsNow = [regex]::Matches($newPara, $runRx)
            for ($r = 0; $r -lt $runsNow.Count; $r++) {
                $len = (Get-RunText -Run $runsNow[$r].Value).Length
                if ($startRun -lt 0 -and $acc + $len -gt $start) { $startRun = $r; $startOff = $start - $acc }
                if ($acc + $len -ge $end) { $endRun = $r; $endOff = $end - $acc; break }
                $acc += $len
            }
            if ($startRun -lt 0 -or $endRun -lt 0) { continue }

            $spanStart = $runsNow[$startRun].Index
            $spanEnd   = $runsNow[$endRun].Index + $runsNow[$endRun].Length
            if (-not (Test-SpanReplaceable -Para $newPara -Start $spanStart -Length ($spanEnd - $spanStart))) {
                $skipped++
                continue
            }

            $rPr = Get-FirstRPr -Run $runsNow[$startRun].Value
            $field = New-RefField -Bookmark $caption.Bookmark -Cached $caption.Cached -RPr $rPr

            $headParts = Split-Run -Run $runsNow[$startRun].Value -Offset $startOff
            $tailParts = Split-Run -Run $runsNow[$endRun].Value -Offset $endOff

            $replacement = $headParts[0] + $field + $tailParts[1]

            $newPara = $newPara.Remove($spanStart, $spanEnd - $spanStart).Insert($spanStart, $replacement)
            $mentionCount++
        }

        if ($newPara -ne $para) {
            $edits += @{ Index = $paras[$i].Index; Length = $paras[$i].Length; Text = $newPara }
        }
    }

    Write-Host "  Mentions linked: $mentionCount" -ForegroundColor Gray
    if ($renumbered -gt 0) {
        Write-Host "  Matched on the table number alone (chapter renumbered): $renumbered" -ForegroundColor DarkGray
    }
    if ($skipped -gt 0) {
        Write-Host "  Left as text (inside a tracked change): $skipped" -ForegroundColor DarkGray
    }
    if ($unmatched.Count -gt 0) {
        Write-Warning "  No caption for: $($unmatched -join ', ') (left as text)"
    }

    if ($ListOnly) {
        foreach ($k in ($captions.Keys | Sort-Object { [int]($_ -split '-')[0], [int]($_ -split '-')[1] })) {
            $t = $captions[$k].Title
            if ($t.Length -gt 62) { $t = $t.Substring(0, 59) + '...' }
            Write-Host ("    $Label $k -> $($captions[$k].Bookmark)  $t") -ForegroundColor DarkGray
        }
        continue
    }

    if ($edits.Count -eq 0) {
        Write-Host "  Nothing to change. File unchanged." -ForegroundColor DarkGray
        continue
    }

    # Apply back to front so every recorded offset is still valid.
    $sb = New-Object System.Text.StringBuilder $xml
    foreach ($edit in ($edits | Sort-Object { $_.Index } -Descending)) {
        [void]$sb.Remove($edit.Index, $edit.Length)
        [void]$sb.Insert($edit.Index, $edit.Text)
    }
    $newXml = $sb.ToString()

    # GUARD. This script splices into PROSE - it replaces a typed mention with a
    # field - and it had no guard of any kind: no parse, no counts, no refusal
    # path. Visible text cannot be the invariant here, because the "matched on
    # the number alone" path deliberately corrects a mention the chapter
    # renumbering left stale. Structure can be.
    $probeXml = $newXml
    if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }
    try {
        $probe = New-Object System.Xml.XmlDocument
        $probe.XmlResolver = $null
        $probe.LoadXml($probeXml)
    } catch {
        throw "Cross-reference rewrite produced malformed document.xml for ${target}: $($_.Exception.Message). Nothing written."
    }
    # NOT bookmarkStart/bookmarkEnd: bookmarking the captions is this script's
    # own job, so those legitimately go UP. Equality was the wrong invariant and
    # refused Ch02 over its own 10 new caption anchors. What must hold is that
    # they pair and that none is LOST, since a start without its end leaves
    # every REF aimed there unresolvable.
    foreach ($tag in @('<a:blip\b', '<w:drawing\b', '<w:pict\b', '<w:p\b', '<w:tbl>',
                       '<w:ins\b', '<w:del\b', '<w:delText\b', '<w:delInstrText\b',
                       '<w:commentReference\b', '<w:footnoteReference\b')) {
        $n = ([regex]::Matches($newXml, $tag)).Count
        $o = ([regex]::Matches($xml, $tag)).Count
        if ($n -ne $o) {
            throw ("Cross-reference rewrite changed a protected element count in ${target}: " +
                   "$tag $o -> $n. Nothing written.")
        }
    }
    $bsNew = ([regex]::Matches($newXml, '<w:bookmarkStart\b')).Count
    $beNew = ([regex]::Matches($newXml, '<w:bookmarkEnd\b')).Count
    $bsOld = ([regex]::Matches($xml,    '<w:bookmarkStart\b')).Count
    if ($bsNew -ne $beNew) {
        throw "Cross-reference rewrite left $bsNew bookmark starts against $beNew ends in ${target}. Nothing written."
    }
    if ($bsNew -lt $bsOld) {
        throw "Cross-reference rewrite dropped $($bsOld - $bsNew) bookmark(s) in ${target}. Nothing written."
    }

    # Every inserted field must be complete. A REF spliced in half renders as an
    # error in Word and cannot be found by looking at the text.
    $begins = ([regex]::Matches($newXml, 'w:fldCharType="begin"')).Count
    $ends   = ([regex]::Matches($newXml, 'w:fldCharType="end"')).Count
    if ($begins -ne $ends) {
        throw "Cross-reference rewrite left $begins field begins against $ends ends in ${target}. Nothing written."
    }
    # NOT "old + one per edit": an edit is a spliced text RANGE, and most of them
    # are caption bookmark anchors that add no field at all. Counting them as
    # fields refused Ch02 over its own 10 anchors. What holds is that fields only
    # appear, never vanish, and that the new ones are REF fields.
    $oldBegins = ([regex]::Matches($xml, 'w:fldCharType="begin"')).Count
    if ($begins -lt $oldBegins) {
        throw ("Cross-reference rewrite lost $($oldBegins - $begins) field(s) in ${target}. Nothing written.")
    }
    $refNew = ([regex]::Matches($newXml, '<w:instrText[^>]*>\s*REF\b')).Count
    $refOld = ([regex]::Matches($xml,    '<w:instrText[^>]*>\s*REF\b')).Count
    if (($begins - $oldBegins) -ne ($refNew - $refOld)) {
        throw ("Cross-reference rewrite added $($begins - $oldBegins) field(s) but " +
               "$($refNew - $refOld) REF instruction(s) in ${target}. Nothing written.")
    }
    if ([regex]::IsMatch($newXml, '/>\s*/>')) {
        throw "Cross-reference rewrite left stray '/>' markup in ${target}. Nothing written."
    }
    # A mention being corrected is the only reason the visible text may move, so
    # say by how much rather than asserting it did not.
    $oldText = -join ([regex]::Matches($xml,    '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
    $newText = -join ([regex]::Matches($newXml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
    if ($oldText -cne $newText) {
        Write-Host ("  Visible text moved by $($newText.Length - $oldText.Length) char(s) " +
                    "across $($edits.Count) corrected mention(s).") -ForegroundColor DarkGray
    }

    $entries['word/document.xml'] = $utf8NoBom.GetBytes($newXml)

    if (-not $NoBackup) {
        if (-not $Force -and (Test-Path -LiteralPath "$target.xref.bak")) {
            Write-Warning "  Backup $target.xref.bak already exists; leaving it as-is."
        } else {
            Copy-Item -LiteralPath $target -Destination "$target.xref.bak" -Force
        }
    }

    $out = [System.IO.Compression.ZipFile]::Open($target, 'Update')
    foreach ($e in @($out.Entries)) { $e.Delete() }
    foreach ($name in $entries.Keys) {
        $entry = $out.CreateEntry($name)
        $s = $entry.Open()
        $s.Write($entries[$name], 0, $entries[$name].Length)
        $s.Close()
    }
    $out.Dispose()

    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Write-Host "Select all and press F9 in Word to update the fields." -ForegroundColor Cyan
