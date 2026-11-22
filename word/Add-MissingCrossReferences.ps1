<#
.SYNOPSIS
    Adds a cross-reference to every caption that nothing points at, as a
    TRACKED INSERTION in a named author's name.

.DESCRIPTION
    A figure or table that the prose never refers to is a defect in a NIST
    report: the reader is given a figure with no idea when to look at it, and a
    renumber cannot be checked because nothing depends on the number.

    Set-TableCrossReferences converts a mention that is ALREADY typed into a
    REF field. It cannot help where no mention was ever written, which is what
    this script is for. The two are complementary: run that one first, then this
    one over whatever is still unreferenced.

    Where the reference goes. A figure in this report follows the prose that
    discusses it, so the reference belongs at the end of the nearest preceding
    body paragraph - inserted before that paragraph's final full stop, which is
    where the authors put their own, e.g. "... out of service in Puerto Rico
    (Figure 2-9(a))". A heading, caption, graphic, list item or empty paragraph
    is skipped on the way back; if nothing suitable is found within -Lookback
    paragraphs the caption is REPORTED and left alone rather than guessed at.

    Every insertion is wrapped in w:ins so it arrives as a tracked change for
    the author to accept or reject one at a time. That is the point: the
    placement is a judgement, and a judgement should be reviewable rather than
    silently applied.

    The REF points at the caption's existing _Ref bookmark, so it renders the
    caption's own number and follows a renumber. A caption with no bookmark is
    skipped - run Set-TableCrossReferences first, which bookmarks them all.

.EXAMPLE
    .\Add-MissingCrossReferences.ps1 -File ch04.docx -Author "Janel Wilsey" -ListOnly
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [string]$Author = 'Janel Wilsey',
    [string]$Initials,
    [int]$Lookback = 12,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
$target = (Resolve-Path -LiteralPath $File).ProviderPath
$lock = Join-Path (Split-Path -Parent $target) ('~$' + (Split-Path -Leaf $target).Substring(2))
if (Test-Path -LiteralPath $lock) { throw "$(Split-Path -Leaf $target) is open in Word. Close it first." }
if (-not $Initials) { $Initials = (($Author -split '\s+') | ForEach-Object { if ($_.Length) { $_.Substring(0,1) } }) -join '' }

function Get-PText {
    param([string]$P)
    $q = [regex]::Replace($P, '(?s)<w:drawing\b.*?</w:drawing>|<w:pict\b.*?</w:pict>', '')
    $sb = New-Object System.Text.StringBuilder
    foreach ($m in [regex]::Matches($q, '(?s)<w:tab\s*/>|<w:noBreakHyphen\s*/>|<w:t(?:\s[^>]*)?>(.*?)</w:t>')) {
        if     ($m.Value -like '<w:tab*')           { [void]$sb.Append("`t") }
        elseif ($m.Value -like '<w:noBreakHyphen*') { [void]$sb.Append([char]0x2011) }
        else   { [void]$sb.Append($m.Groups[1].Value) }
    }
    return $sb.ToString()
}

$entries = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
$order = New-Object System.Collections.ArrayList
$fs = [System.IO.File]::OpenRead($target)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($fs, 'Read')
    foreach ($e in $zip.Entries) {
        $ms = New-Object System.IO.MemoryStream
        $s = $e.Open(); $s.CopyTo($ms); $s.Close()
        $entries[$e.FullName] = $ms.ToArray(); [void]$order.Add($e.FullName)
    }
    $zip.Dispose()
} finally { $fs.Dispose() }

$xml = $utf8NoBom.GetString($entries['word/document.xml'])
$bodyM = [regex]::Match($xml, '(?s)<w:body(?:\s[^>]*)?>.*</w:body>')
if (-not $bodyM.Success) { throw "No <w:body> in $target" }
$body = $bodyM.Value

# Top-level paragraphs and tables, in order.
$els = New-Object System.Collections.ArrayList
$depth = 0; $start = 0; $kind = $null
foreach ($m in [regex]::Matches($body, '(?s)<w:p\b[^>]*/>|<w:p\b[^>]*>|</w:p>|<w:tbl>|</w:tbl>')) {
    $v = $m.Value
    if ($v -match '^<w:p\b[^>]*/>$') { if ($depth -eq 0) { [void]$els.Add([pscustomobject]@{ Kind='p'; Start=$m.Index; End=$m.Index+$m.Length }) }; continue }
    if ($v -eq '</w:p>' -or $v -eq '</w:tbl>') {
        $depth--
        if ($depth -eq 0) { [void]$els.Add([pscustomobject]@{ Kind=$kind; Start=$start; End=$m.Index+$m.Length }) }
        continue
    }
    if ($depth -eq 0) { $start = $m.Index; $kind = if ($v -eq '<w:tbl>') { 'tbl' } else { 'p' } }
    $depth++
}

$liveRefs = @([regex]::Matches($xml, '<w:instrText(?:\s[^>]*)?>\s*REF\s+(\S+)') | ForEach-Object { $_.Groups[1].Value })

$plan = New-Object System.Collections.ArrayList
$skips = New-Object System.Collections.ArrayList
$used = New-Object System.Collections.ArrayList
for ($i = 0; $i -lt $els.Count; $i++) {
    $el = $els[$i]
    if ($el.Kind -ne 'p') { continue }
    $p = $body.Substring($el.Start, $el.End - $el.Start)
    if ($p -notmatch '<w:pStyle\s+w:val="Caption\w*"\s*/>') { continue }
    $t = Get-PText -P $p
    $nm = [regex]::Match($t, '^\s*(Figure|Table)[\s\u00A0]*([0-9A-Za-z]+[\u2013\u2011\u2010.\-][0-9]+)')
    if (-not $nm.Success) { continue }
    $label = $nm.Groups[1].Value; $number = $nm.Groups[2].Value

    $bms = @([regex]::Matches($p, '<w:bookmarkStart\s[^>]*w:name="(_Ref\d+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($bms.Count -eq 0) { [void]$skips.Add("$label $number - no _Ref bookmark; run Set-TableCrossReferences first"); continue }
    $already = $false
    foreach ($b in $bms) { if ($liveRefs -contains $b) { $already = $true; break } }
    if ($already) { continue }
    $bm = $bms[0]

    # Walk back for a body paragraph that can carry the reference.
    $hostIdx = -1
    for ($j = $i - 1; $j -ge [Math]::Max(0, $i - $Lookback); $j--) {
        if ($els[$j].Kind -ne 'p') { continue }
        $q = $body.Substring($els[$j].Start, $els[$j].End - $els[$j].Start)
        if ($q -match '<w:pStyle\s+w:val="(?:Caption\w*|graphic|Heading\d|AppendixHead\w*|Frontmatterhead\w*|ListBullet\w*|ListNumber\w*|ListParagraph|References|TableText|TableandFigureFootnotes)"\s*/>') { continue }
        if ($q -match '<a:blip\b') { continue }
        $qt = Get-PText -P $q
        if ($qt.Trim().Length -lt 40) { continue }
        # It must be a SENTENCE. A paragraph that ends in a URL is a reference
        # entry whatever style it carries - two of Ch04's references were about
        # to be planted inside
        # "https://agupubs.onlinelibrary.wiley.com/doi/epdf/10.1029/..." - and
        # one ending in a colon is introducing the figure, not describing it, so
        # the reference would land after the colon.
        if ($qt -match 'https?://') { continue }
        if ($qt.TrimEnd() -notmatch '[.?!]$') { continue }
        # A host MAY carry more than one. Ch04's six imagery-extent figures share
        # a single intro sentence and have nothing but other captions above
        # them, so a one-per-paragraph rule left all six unreferenced - the gap
        # was worse than the crowding. Several landing on one host are combined
        # into a single parenthetical below rather than stacked.
        $hostIdx = $j; break
    }
    if ($hostIdx -lt 0) { [void]$skips.Add("$label $number - no free prose sentence within $Lookback above it"); continue }

    [void]$used.Add($hostIdx)
    [void]$plan.Add([pscustomobject]@{
        Label = $label; Number = $number; Bookmark = $bm
        CapIdx = $i; HostIdx = $hostIdx
        HostText = (Get-PText -P $body.Substring($els[$hostIdx].Start, $els[$hostIdx].End - $els[$hostIdx].Start))
    })
}

Write-Host "Add Missing Cross-References" -ForegroundColor Cyan
Write-Host "  $(Split-Path -Leaf $target)" -ForegroundColor Gray
Write-Host "  Inserted as a tracked change in the name of: $Author" -ForegroundColor Gray
foreach ($s in $skips) { Write-Warning "    skipped: $s" }
if ($plan.Count -eq 0) { Write-Host "  Every caption already has a cross-reference." -ForegroundColor Green; return }
foreach ($pl in $plan) {
    $tail = $pl.HostText.TrimEnd()
    if ($tail.Length -gt 76) { $tail = '...' + $tail.Substring($tail.Length - 73) }
    Write-Host ("    $($pl.Label) $($pl.Number) -> after: $tail") -ForegroundColor DarkGray
}
Write-Host "  $($plan.Count) reference(s) to insert." -ForegroundColor Gray
if ($ListOnly) { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow; return }
if (-not $Force) {
    $reply = Read-Host "Insert $($plan.Count) tracked cross-reference(s)? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

$date = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss') + 'Z'
$maxId = 0
foreach ($m in [regex]::Matches($xml, 'w:id="(\d+)"')) { $v = [int]$m.Groups[1].Value; if ($v -gt $maxId) { $maxId = $v } }

# Group by host so several figures sharing one sentence become ONE parenthetical.
$byHost = @{}
foreach ($pl in $plan) {
    if (-not $byHost.ContainsKey($pl.HostIdx)) { $byHost[$pl.HostIdx] = New-Object System.Collections.ArrayList }
    [void]$byHost[$pl.HostIdx].Add($pl)
}
$grouped = New-Object System.Collections.ArrayList
foreach ($h in $byHost.Keys) {
    $g = @($byHost[$h])
    if ($g.Count -eq 1) { [void]$grouped.Add($g[0]); continue }
    # Consecutive numbers of the same label read as a range: "Figure 4-28
    # through Figure 4-33". Anything else is listed.
    $nums = @($g | ForEach-Object { [int](($_.Number -split '[\u2013\u2011\u2010.\-]')[-1]) })
    $sorted = @($nums | Sort-Object)
    $isRun = $true
    for ($q = 1; $q -lt $sorted.Count; $q++) { if ($sorted[$q] -ne $sorted[$q-1] + 1) { $isRun = $false; break } }
    $sameLabel = (@($g | ForEach-Object { $_.Label } | Sort-Object -Unique).Count -eq 1)
    $ordered = @($g | Sort-Object { [int](($_.Number -split '[\u2013\u2011\u2010.\-]')[-1]) })
    [void]$grouped.Add([pscustomobject]@{
        HostIdx = $h
        Members = $ordered
        Range   = ($isRun -and $sameLabel -and $ordered.Count -gt 1)
    })
}

# Build each insertion, then splice BACK TO FRONT over the body.
$edits = New-Object System.Collections.ArrayList
foreach ($pl in $plan) {
    $h = $els[$pl.HostIdx]
    $hostXml = $body.Substring($h.Start, $h.End - $h.Start)
    # The last text node that ends the paragraph's final sentence.
    $tnodes = @([regex]::Matches($hostXml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>'))
    if ($tnodes.Count -eq 0) { continue }
    $last = $tnodes[$tnodes.Count - 1]
    $lastText = $last.Groups[1].Value
    # Insert before a trailing full stop (and any trailing spaces), else at the end.
    $m2 = [regex]::Match($lastText, '\.\s*$')
    $cut = if ($m2.Success) { $m2.Index } else { $lastText.Length }
    $maxId++
    $ins = '<w:ins w:id="' + $maxId + '" w:author="' + ($Author -replace '&','&amp;') + '" w:date="' + $date + '">' +
           '<w:r><w:t xml:space="preserve"> (</w:t></w:r>' +
           '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
           '<w:r><w:instrText xml:space="preserve"> REF ' + $pl.Bookmark + ' \h </w:instrText></w:r>' +
           '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
           '<w:r><w:t xml:space="preserve">' + $pl.Label + [char]0x00A0 + $pl.Number + '</w:t></w:r>' +
           '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
           '<w:r><w:t xml:space="preserve">)</w:t></w:r>' +
           '</w:ins>'
    # Split the host's last text node at $cut and drop the insertion between.
    $absTextStart = $h.Start + $last.Groups[1].Index
    $before = $lastText.Substring(0, $cut)
    $after  = $lastText.Substring($cut)
    $openTag = [regex]::Match($hostXml.Substring($last.Index), '^<w:t(?:\s[^>]*)?>').Value
    $runOpen = '<w:r><w:t xml:space="preserve">'
    $replacement = $before + '</w:t></w:r>' + $ins + $runOpen + $after
    [void]$edits.Add([pscustomobject]@{ Start = $absTextStart; Len = $lastText.Length; Text = $replacement })
}

$newBody = $body
foreach ($e in @($edits | Sort-Object Start -Descending)) {
    $newBody = $newBody.Remove($e.Start, $e.Len).Insert($e.Start, $e.Text)
}
$newXml = $xml.Remove($bodyM.Index, $body.Length).Insert($bodyM.Index, $newBody)

# ---- guards ----
$probeXml = $newXml
if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }
try {
    $probe = New-Object System.Xml.XmlDocument
    $probe.XmlResolver = $null
    $probe.LoadXml($probeXml)
} catch { throw "Adding cross-references produced malformed document.xml for ${target}: $($_.Exception.Message). Nothing written." }
foreach ($tag in @('<a:blip\b', '<w:drawing\b', '<w:pict\b', '<w:p\b', '<w:tbl>',
                   '<w:del\b', '<w:delText\b', '<w:commentReference\b', '<w:footnoteReference\b',
                   '<w:bookmarkStart\b', '<w:bookmarkEnd\b')) {
    $a = ([regex]::Matches($newXml, $tag)).Count
    $b = ([regex]::Matches($xml, $tag)).Count
    if ($a -ne $b) { throw "Protected count changed in ${target}: $tag $b -> $a. Nothing written." }
}
$insA = ([regex]::Matches($newXml, '<w:ins\b')).Count
$insB = ([regex]::Matches($xml, '<w:ins\b')).Count
if (($insA - $insB) -ne $plan.Count) { throw "Expected $($plan.Count) new w:ins in ${target}, got $($insA - $insB). Nothing written." }
$fb = ([regex]::Matches($newXml, 'w:fldCharType="begin"')).Count - ([regex]::Matches($xml, 'w:fldCharType="begin"')).Count
$fe = ([regex]::Matches($newXml, 'w:fldCharType="end"')).Count - ([regex]::Matches($xml, 'w:fldCharType="end"')).Count
if ($fb -ne $plan.Count -or $fe -ne $plan.Count) { throw "Field begin/end delta is $fb/$fe, expected $($plan.Count) each in ${target}. Nothing written." }
if ([regex]::IsMatch($newXml, '/>\s*/>')) { throw "Stray '/>' markup in ${target}. Nothing written." }
if ([regex]::IsMatch($newXml, '<w:t[^>]*>\s*</w:t>\s*</w:r>\s*<w:r>\s*<w:t[^>]*>\s*</w:t>')) {
    Write-Warning "  Empty text runs were produced; harmless but untidy."
}

if (-not $NoBackup) {
    $bak = "$target.xrefadd.bak"
    if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
}
$entries['word/document.xml'] = $utf8NoBom.GetBytes($newXml)
$out = [System.IO.Compression.ZipFile]::Open($target, 'Update')
foreach ($e in @($out.Entries)) { $e.Delete() }
foreach ($name in $order) {
    $entry = $out.CreateEntry($name)
    $s = $entry.Open(); $s.Write($entries[$name], 0, $entries[$name].Length); $s.Close()
}
$out.Dispose()
Write-Host "  Saved. $($plan.Count) tracked cross-reference(s) added." -ForegroundColor Green
