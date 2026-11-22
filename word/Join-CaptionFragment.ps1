<#
.SYNOPSIS
    Merges a stranded caption fragment back into the caption above it, when the
    author marked it with a leading character.

.DESCRIPTION
    A caption written in two halves is a common author state: the label and
    number end up on the caption style, the title ends up in the next
    paragraph on Body Text, and in NWIRP 2-1 Ch04 the author marked the orphan
    half with a leading "^". On the page it reads as a caption with no title
    followed by a stray sentence, and a caption check that only inspects
    captions that EXIST passes it, because the caption is there - it is just
    missing half its text.

    This moves the fragment's RUNS into the caption paragraph, so italics, a
    superscript or a symbol in the title survive, and removes the marker
    character and the now-empty paragraph.

    Deliberately narrow. The fragment must be the IMMEDIATELY following
    paragraph, it must carry the marker, and the paragraph above it must be on
    a caption style. Anything else is reported and left alone: this moves text
    between paragraphs, which is the one thing a formatting pass normally must
    not do, so it only acts where the author left an explicit instruction.

.PARAMETER Marker
    The leading character the author used. Default "^". Removed on merge.

.PARAMETER FromStyle
    Styles a fragment may be sitting on. Default BodyText.

.PARAMETER CaptionStyle
    Styles that count as a caption to merge into.

.PARAMETER ListOnly
    Report and write nothing. Run this first and read it.

.EXAMPLE
    .\Join-CaptionFragment.ps1 -Folder "C:\...\HM-261109-v2" -ListOnly

.EXAMPLE
    .\Join-CaptionFragment.ps1 -File "Ch04.docx" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$Marker = '^',

    [string[]]$FromStyle = @('BodyText'),

    [string[]]$CaptionStyle = @('CaptionFigure', 'CaptionTable', 'CaptionFigure2', 'Caption'),

    # A fragment may carry no style at all. In NWIRP 2-1 Ch04 one of the two is
    # on Body Text and the other has no pStyle, so requiring -FromStyle found
    # only half of them.
    [switch]$IncludeUnstyled,

    # Text-free paragraphs allowed between the caption and its fragment. The
    # same two captions differ here too: one fragment is adjacent, the other has
    # an empty self-closing paragraph in between.
    [int]$MaxGap = 1,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$fromAlt = ($FromStyle    | ForEach-Object { [regex]::Escape($_) }) -join '|'
$capAlt   = ($CaptionStyle | ForEach-Object { [regex]::Escape($_) }) -join '|'

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


Write-Host "Join Caption Fragment" -ForegroundColor Cyan
Write-Host ("  A {0} paragraph on {1} directly under a caption is merged into it." -f "'$Marker'", ($FromStyle -join '/')) -ForegroundColor Cyan
Write-Host "  The marker is removed. Runs are moved, so inline formatting survives." -ForegroundColor Cyan
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

# Paragraph text, with the drawing stripped first: text inside a txbxContent
# belongs to the picture, not the paragraph, and counting it as paragraph text
# is what made five image paragraphs look like stranded caption halves.
function Get-ParaText {
    param([string]$P)
    $x = [regex]::Replace($P, '(?s)<w:drawing\b.*?</w:drawing>', '')
    $x = [regex]::Replace($x, '(?s)<w:pict\b.*?</w:pict>', '')
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in [regex]::Matches($x, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { [void]$sb.Append($t.Groups[1].Value) }
    return $sb.ToString()
}

function Get-AllText {
    param([string]$Xml)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in [regex]::Matches($Xml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { [void]$sb.Append($t.Groups[1].Value) }
    return $sb.ToString()
}

$filesChanged = 0
foreach ($target in $targets) {
    $name = [System.IO.Path]::GetFileName($target)
    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    if (-not $pkg.Data.ContainsKey('word/document.xml')) { Write-Host "  No document.xml." -ForegroundColor DarkGray; continue }
    $raw = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/document.xml'])
    $hadBom = ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF)
    $doc = if ($hadBom) { $raw.Substring(1) } else { $raw }

    # SELF-CLOSING FIRST. Alternation is ordered, and '<w:p\b[^>]*?>' also
    # matches '<w:p .../>' because [^>]*? absorbs the slash - so with the
    # open/close form first, an empty <w:p/> is consumed by it and the match runs
    # on to the NEXT paragraph's </w:p>, returning two paragraphs as one. That
    # is how this pass first produced malformed XML: the "fragment" it spliced
    # out spanned an empty paragraph plus the real one after it.
    $paras = @([regex]::Matches($doc, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'))
    $merges = New-Object System.Collections.ArrayList
    for ($i = 1; $i -lt $paras.Count; $i++) {
        $frag = $paras[$i].Value
        $ft = Get-ParaText -P $frag
        if (-not $ft.TrimStart().StartsWith($Marker)) { continue }
        # A fragment holding an image is a figure, not stranded caption text.
        if ($frag -match '<w:drawing\b|<w:pict\b') { continue }

        $hasStyle = $frag -match '<w:pStyle\b[^>]*?\bw:val="(?:' + $fromAlt + ')"\s*/>'
        $unstyled = $frag -notmatch '<w:pStyle\b'
        if (-not ($hasStyle -or ($IncludeUnstyled -and $unstyled))) { continue }

        # Walk back over text-free paragraphs to find the caption.
        $j = $i - 1
        $gap = 0
        while ($j -ge 0 -and $gap -le $MaxGap) {
            $cand = $paras[$j].Value
            if ($cand -match '<w:pStyle\b[^>]*?\bw:val="(?:' + $capAlt + ')"\s*/>') { break }
            if ((Get-ParaText -P $cand).Trim()) { $j = -1; break }   # real text: not a gap
            $gap++; $j--
        }
        if ($j -lt 0 -or $gap -gt $MaxGap) { continue }
        if ($paras[$j].Value -notmatch '<w:pStyle\b[^>]*?\bw:val="(?:' + $capAlt + ')"\s*/>') { continue }

        [void]$merges.Add([pscustomobject]@{
            CapIndex = $paras[$j].Index; CapLength = $paras[$j].Length; Cap = $paras[$j].Value
            FragIndex = $paras[$i].Index;  FragLength = $paras[$i].Length;  Frag = $frag
            CapText = (Get-ParaText -P $paras[$j].Value).Trim(); FragText = $ft.Trim()
        })
    }

    if ($merges.Count -eq 0) { Write-Host "  No marked fragments." -ForegroundColor Green; continue }
    foreach ($m in $merges) {
        Write-Host ("    {0}" -f $m.CapText.Substring(0, [Math]::Min(40, $m.CapText.Length))) -ForegroundColor DarkGray
        Write-Host ("      + {0}" -f $m.FragText.Substring(0, [Math]::Min(86, $m.FragText.Length))) -ForegroundColor DarkGray
    }

    # BACK TO FRONT: each merge removes the fragment paragraph and lengthens the
    # caption, so forwards every later offset is stale.
    $newDoc = $doc
    $markersDropped = 0
    foreach ($m in @($merges | Sort-Object -Property FragIndex -Descending)) {
        # The fragment's runs are everything after its pPr (or after the open
        # tag when it has none) up to </w:p>.
        $body = $m.Frag
        $pprM = [regex]::Match($body, '(?s)<w:pPr\b[^>]*?>.*?</w:pPr>|<w:pPr\s*/>')
        $start = if ($pprM.Success) { $pprM.Index + $pprM.Length } else { [regex]::Match($body, '(?s)^<w:p\b[^>]*?>').Length }
        $runs = $body.Substring($start, $body.Length - $start - '</w:p>'.Length)

        # Drop the marker from the FIRST w:t that carries it, nothing else.
        $tm = [regex]::Match($runs, '(?s)(<w:t(?:\s[^>]*)?>)(\s*)' + [regex]::Escape($Marker))
        if ($tm.Success) {
            $runs = $runs.Remove($tm.Index, $tm.Length).Insert($tm.Index, $tm.Groups[1].Value + $tm.Groups[2].Value)
            $markersDropped++
        }

        # Remove the fragment paragraph, then append its runs to the caption.
        $newDoc = $newDoc.Remove($m.FragIndex, $m.FragLength)
        $capEnd = $m.CapIndex + $m.CapLength - '</w:p>'.Length
        $newDoc = $newDoc.Insert($capEnd, $runs)
    }

    # --- guards ----------------------------------------------------------
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Merge produced malformed document.xml for ${name}: $($_.Exception.Message). Nothing written." }

    # Visible text may lose exactly the marker characters and nothing else.
    $before = Get-AllText -Xml $doc
    $after  = Get-AllText -Xml $newDoc
    if ($before.Replace($Marker, '') -cne $after.Replace($Marker, '')) {
        throw "Merge changed text beyond the marker in ${name}. Nothing written."
    }
    $expected = $before.Length - $markersDropped
    if ($after.Length -ne $expected) {
        throw "Merge changed the text length by $($after.Length - $before.Length) in ${name}, expected -$markersDropped. Nothing written."
    }

    $pN = ([regex]::Matches($newDoc, '<w:p\b')).Count
    $pO = ([regex]::Matches($doc, '<w:p\b')).Count
    if ($pN -ne $pO - $merges.Count) { throw "Paragraph count changed by $($pN - $pO) in ${name}, expected -$($merges.Count). Nothing written." }

    foreach ($tag in @('<w:ins\b', '<w:del\b', '<w:bookmarkStart\b', '<w:bookmarkEnd\b',
                       '<a:blip\b', '<w:fldChar\b', '<w:fldSimple\b', '<w:instrText\b', '<w:tbl>', '<w:sectPr\b', '<w:r\b')) {
        $n = ([regex]::Matches($newDoc, $tag)).Count
        $o = ([regex]::Matches($doc, $tag)).Count
        if ($n -ne $o) { throw "Protected element count changed in ${name}: $tag $o -> $n. Nothing written." }
    }
    $stray = ([regex]::Matches($newDoc, '/>\s*/>')).Count - ([regex]::Matches($doc, '/>\s*/>')).Count
    if ($stray -gt 0) { throw "Merge left $stray stray '/>' in ${name}. Nothing written." }

    Write-Host ("  {0} fragment(s) merged, {1} marker(s) dropped." -f $merges.Count, $markersDropped) -ForegroundColor DarkGray
    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        $bak = "$target.capjoin.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "capjoin-$([guid]::NewGuid()).zip")
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
        $ti = Get-Item -LiteralPath $target -Force
        if ($ti.IsReadOnly) { $ti.IsReadOnly = $false }
        [System.IO.File]::Copy($tempZip, $target, $true)
        Remove-Item -LiteralPath $tempZip -Force
        $filesChanged++
        Write-Host "  Saved." -ForegroundColor Green
    } catch {
        Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue
        Write-Host "  LOCKED (open in Word) - not written." -ForegroundColor Red
    }
}

Write-Host ""
Write-Host ("Done. {0} file(s) changed." -f $filesChanged) -ForegroundColor Green
