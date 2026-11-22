<#
.SYNOPSIS
    Writes a caption's title, optionally absorbing the orphan paragraph the
    title was sitting in, and attaches a Word comment in a named author's name.

.DESCRIPTION
    Two different defects leave a caption without a title, and they need
    different handling:

      - The title does not exist. Someone has to write one, and because that is
        an editorial judgement it must be visible AS a suggestion rather than
        slipped in as fact. -Comment attaches a Word comment to the caption in
        -Author's name saying so.
      - The title DOES exist, orphaned into the paragraph below the caption.
        Ch04's "Aerial imagery extents" figures are built image / caption /
        label, six times over, with the label line carrying the real title.
        -AbsorbNext takes that paragraph's text as the title and removes it, so
        the authors' own words are used instead of invented ones.

    The comment is anchored on the caption's own text, so it opens beside the
    caption in the review pane.

.PARAMETER Number
    The caption number as it RENDERS, e.g. 4-28 or 3-10. Matched against the
    caption's visible text, with any dash form accepted.

.EXAMPLE
    .\Set-CaptionTitle.ps1 -File ch04.docx -Number 4-28 -AbsorbNext `
        -Comment "Suggested caption - please confirm." -Author "Janel Wilsey" -Force
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [Parameter(Mandatory = $true)]
    [string]$Number,

    # Figure and Table number SEPARATELY, so "3-4" is ambiguous: a chapter has
    # both a Figure 3-4 and a Table 3-4. Matching either label took the first in
    # document order - always the figure - and wrote seven table titles onto
    # seven real figure captions, destroying titles that were already correct.
    # -Label is mandatory for that reason; there is no safe default.
    [Parameter(Mandatory = $true)]
    [ValidateSet('Figure','Table')]
    [string]$Label,

    [string]$Title,

    [switch]$AbsorbNext,

    [string]$Comment,
    [string]$Author = 'Janel Wilsey',
    [string]$Initials,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$DASH = '[\u2013\u2011\u2010.\-]'

if (-not $Title -and -not $AbsorbNext) { throw "Give -Title, or -AbsorbNext to take the title from the paragraph below." }
if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
$target = (Resolve-Path -LiteralPath $File).ProviderPath
$lock = Join-Path (Split-Path -Parent $target) ('~$' + (Split-Path -Leaf $target).Substring(2))
if (Test-Path -LiteralPath $lock) { throw "$(Split-Path -Leaf $target) is open in Word. Close it first." }
if (-not $Initials) {
    $Initials = (($Author -split '\s+') | ForEach-Object { if ($_.Length) { $_.Substring(0,1) } }) -join ''
}

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
$paras = @([regex]::Matches($xml, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'))

$numRx = '^\s*' + $Label + '[\s\u00A0]*' + ($Number -replace $DASH, $DASH) + '\s*\.'
$idx = -1
for ($i = 0; $i -lt $paras.Count; $i++) {
    $p = $paras[$i].Value
    if ($p -notmatch '<w:pStyle\s+w:val="Caption\w*"\s*/>') { continue }
    if ((Get-PText -P $p) -match $numRx) { $idx = $i; break }
}
if ($idx -lt 0) { throw "No $Label caption rendering '$Number' found in $target." }

$cap = $paras[$idx].Value
$capText = Get-PText -P $cap

# Where the title begins: after the number's terminator and its spaces.
$lead = [regex]::Match($capText, '^\s*' + $Label + '[\s\u00A0]*' + ($Number -replace $DASH, $DASH) + '\s*[.:]\s*')
if (-not $lead.Success) { throw "Could not find the end of the number in '$capText'." }
$existing = $capText.Substring($lead.Length)

$newTitle = $Title
$absorbIdx = -1
if ($AbsorbNext) {
    for ($j = $idx + 1; $j -lt [Math]::Min($idx + 3, $paras.Count); $j++) {
        $q = $paras[$j].Value
        if ($q -match '<a:blip\b') { break }
        $qt = (Get-PText -P $q).Trim()
        if (-not $qt) { continue }
        if ($q -match '<w:pStyle\s+w:val="Caption\w*"\s*/>') { break }
        $absorbIdx = $j; if (-not $newTitle) { $newTitle = $qt }
        break
    }
    if ($absorbIdx -lt 0) { throw "-AbsorbNext found no text paragraph under the caption rendering '$Number'." }
}
if (-not $newTitle) { throw "No title to write." }

Write-Host "Set Caption Title" -ForegroundColor Cyan
Write-Host "  $(Split-Path -Leaf $target)  caption $Number" -ForegroundColor Gray
Write-Host "    was: $($existing.Trim())" -ForegroundColor DarkGray
Write-Host "    now: $newTitle" -ForegroundColor Green
if ($absorbIdx -ge 0) { Write-Host "    absorbing paragraph $absorbIdx" -ForegroundColor DarkGray }
if ($Comment) { Write-Host "    comment as ${Author}: $Comment" -ForegroundColor DarkGray }
if ($ListOnly) { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow; return }
if (-not $Force) {
    $reply = Read-Host "Apply? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---- build the new caption paragraph ----
# Replace every text node after the number with nothing, then append one run
# carrying the title. The number's own runs and fields are untouched.
$acc = 0; $edits = New-Object System.Collections.ArrayList
foreach ($m in [regex]::Matches($cap, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) {
    $g = $m.Groups[1]
    $s = $acc; $e = $acc + $g.Length; $acc = $e
    if ($e -le $lead.Length) { continue }
    $keepLen = [Math]::Max(0, $lead.Length - $s)
    [void]$edits.Add([pscustomobject]@{ Start = $g.Index; Len = $g.Length; Keep = $g.Value.Substring(0, $keepLen) })
}
$newCap = $cap
foreach ($ed in @($edits | Sort-Object Start -Descending)) {
    $newCap = $newCap.Remove($ed.Start, $ed.Len).Insert($ed.Start, $ed.Keep)
}
$esc = $newTitle -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
$titleRun = '<w:r><w:t xml:space="preserve">' + $esc + '</w:t></w:r>'

# ---- the comment ----
$cid = -1
if ($Comment) {
    if (-not $entries.ContainsKey('word/comments.xml')) { throw "No word/comments.xml in $target; cannot add a comment." }
    $cmts = $utf8NoBom.GetString($entries['word/comments.xml'])
    $maxId = -1
    foreach ($m in [regex]::Matches($cmts, '<w:comment\b[^>]*\bw:id="(\d+)"')) {
        $v = [int]$m.Groups[1].Value; if ($v -gt $maxId) { $maxId = $v }
    }
    foreach ($m in [regex]::Matches($xml, 'w:id="(\d+)"')) {
        $v = [int]$m.Groups[1].Value; if ($v -gt $maxId) { $maxId = $v }
    }
    $cid = $maxId + 1
    $date = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss') + 'Z'
    $ctext = $Comment -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
    $newCmt = '<w:comment w:id="' + $cid + '" w:author="' + ($Author -replace '&','&amp;') +
              '" w:date="' + $date + '" w:initials="' + $Initials + '">' +
              '<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>' +
              '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' +
              '<w:r><w:t xml:space="preserve">' + $ctext + '</w:t></w:r></w:p></w:comment>'
    $closeIdx = $cmts.LastIndexOf('</w:comments>')
    if ($closeIdx -lt 0) { throw "Malformed comments.xml in $target." }
    $newCmts = $cmts.Insert($closeIdx, $newCmt)

    # anchor it: range around the caption's content, reference at the end
    $pprM = [regex]::Match($newCap, '(?s)<w:pPr(?:\s[^>]*)?>.*?</w:pPr>')
    $at = if ($pprM.Success) { $pprM.Index + $pprM.Length } else { [regex]::Match($newCap, '^<w:p\b[^>]*>').Length }
    $newCap = $newCap.Insert($at, '<w:commentRangeStart w:id="' + $cid + '"/>')
}

# append the title run, and the comment end marker, just before </w:p>
$tail = ''
if ($cid -ge 0) {
    $tail = '<w:commentRangeEnd w:id="' + $cid + '"/>' +
            '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="' + $cid + '"/></w:r>'
}
$closeP = $newCap.LastIndexOf('</w:p>')
$newCap = $newCap.Insert($closeP, $titleRun + $tail)

# ---- splice back, back to front ----
$newXml = $xml
if ($absorbIdx -ge 0) {
    $am = $paras[$absorbIdx]
    if ($am.Index -gt $paras[$idx].Index) {
        $newXml = $newXml.Remove($am.Index, $am.Length)
        $newXml = $newXml.Remove($paras[$idx].Index, $paras[$idx].Length).Insert($paras[$idx].Index, $newCap)
    } else {
        $newXml = $newXml.Remove($paras[$idx].Index, $paras[$idx].Length).Insert($paras[$idx].Index, $newCap)
        $newXml = $newXml.Remove($am.Index, $am.Length)
    }
} else {
    $newXml = $newXml.Remove($paras[$idx].Index, $paras[$idx].Length).Insert($paras[$idx].Index, $newCap)
}

# ---- guards ----
$probeXml = $newXml
if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }
try {
    $probe = New-Object System.Xml.XmlDocument
    $probe.XmlResolver = $null
    $probe.LoadXml($probeXml)
} catch { throw "Setting the caption title produced malformed document.xml for ${target}: $($_.Exception.Message). Nothing written." }
foreach ($tag in @('<a:blip\b', '<w:drawing\b', '<w:pict\b', '<w:tbl>', '<w:ins\b', '<w:del\b',
                   '<w:delText\b', '<w:footnoteReference\b', '<w:fldChar\b', '<w:fldSimple\b')) {
    $a = ([regex]::Matches($newXml, $tag)).Count
    $b = ([regex]::Matches($xml, $tag)).Count
    if ($a -ne $b) { throw "Protected count changed in ${target}: $tag $b -> $a. Nothing written." }
}
$expectedParaDelta = if ($absorbIdx -ge 0) { -1 } else { 0 }
$pa = ([regex]::Matches($newXml, '<w:p\b')).Count
$pb = ([regex]::Matches($xml, '<w:p\b')).Count
$cmtPara = if ($cid -ge 0) { 0 } else { 0 }
if (($pa - $pb) -ne $expectedParaDelta) {
    throw "Paragraph count moved $pb -> $pa in ${target}, expected $expectedParaDelta. Nothing written."
}
$ra = ([regex]::Matches($newXml, '<w:commentRangeStart\b')).Count
$re_ = ([regex]::Matches($newXml, '<w:commentRangeEnd\b')).Count
$rf = ([regex]::Matches($newXml, '<w:commentReference\b')).Count
if ($ra -ne $re_ -or $ra -ne $rf) { throw "Comment anchors do not balance in ${target} ($ra starts, $re_ ends, $rf refs). Nothing written." }
if ([regex]::IsMatch($newXml, '/>\s*/>')) { throw "Stray '/>' markup in ${target}. Nothing written." }
# the caption must now read number + title
$check = Get-PText -P ([regex]::Match($newXml.Substring($paras[$idx].Index), '(?s)^<w:p\b[^>]*?>.*?</w:p>').Value)
if ($check -notlike "*$newTitle*") { throw "The caption does not contain the new title in ${target}. Nothing written." }

if (-not $NoBackup) {
    $bak = "$target.captitle.bak"
    if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
}
$entries['word/document.xml'] = $utf8NoBom.GetBytes($newXml)
if ($cid -ge 0) { $entries['word/comments.xml'] = $utf8NoBom.GetBytes($newCmts) }

$out = [System.IO.Compression.ZipFile]::Open($target, 'Update')
foreach ($e in @($out.Entries)) { $e.Delete() }
foreach ($name in $order) {
    $entry = $out.CreateEntry($name)
    $s = $entry.Open(); $s.Write($entries[$name], 0, $entries[$name].Length); $s.Close()
}
$out.Dispose()
Write-Host "  Saved." -ForegroundColor Green
