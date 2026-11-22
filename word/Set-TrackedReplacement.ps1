<#
.SYNOPSIS
    Replaces a paragraph's text as a TRACKED change: the old text becomes a
    tracked deletion and the new text a tracked insertion, both in a named
    author's name.

.DESCRIPTION
    Written for a bracketed placeholder that needs real wording. Supplying the
    wording is an editorial act, so it arrives as a revision the author accepts
    or rejects rather than as a silent edit - and the placeholder goes out the
    same way, visibly, instead of simply vanishing.

    A bookmark in the paragraph is PRESERVED and the insertion is placed inside
    it. An empty heading that is the target of a cross-reference carries its
    anchor there, and a REF aimed at it resolves to nothing if the anchor is
    dropped or if the new text lands outside it.

    The old runs are converted properly: w:t becomes w:delText inside the
    w:del, which is how Word records deleted characters. Leaving them as w:t
    would show the text as still present.

.PARAMETER Matching
    Regex matched against the paragraph's visible text. Must match exactly one
    paragraph, or the script refuses - a replacement applied to the wrong
    paragraph is worse than none.

.EXAMPLE
    .\Set-TrackedReplacement.ps1 -File ch06.docx `
        -Matching '\[PLACEHOLDER - appendix title for Table B\.1\]' `
        -Text 'Optimization models for pre- and post-hazard planning' -Author 'Janel Wilsey' -Force
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [Parameter(Mandatory = $true)]
    [string]$Matching,

    [Parameter(Mandatory = $true)]
    [string]$Text,

    [string]$Author = 'Janel Wilsey',
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

function Get-PText {
    param([string]$P)
    $q = [regex]::Replace($P, '(?s)<w:drawing\b.*?</w:drawing>|<w:pict\b.*?</w:pict>', '')
    $sb = New-Object System.Text.StringBuilder
    foreach ($m in [regex]::Matches($q, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { [void]$sb.Append($m.Groups[1].Value) }
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
$hits = New-Object System.Collections.ArrayList
for ($i = 0; $i -lt $paras.Count; $i++) {
    if ((Get-PText -P $paras[$i].Value) -match $Matching) { [void]$hits.Add($i) }
}
if ($hits.Count -eq 0) { throw "No paragraph matching '$Matching' in $target." }
if ($hits.Count -gt 1) { throw "$($hits.Count) paragraphs match '$Matching' in $target; make the pattern exact. Nothing written." }

$idx = $hits[0]
$pm = $paras[$idx]
$p = $pm.Value
$oldText = Get-PText -P $p
$sty = [regex]::Match($p, '<w:pStyle\s+w:val="([^"]*)"\s*/>')

Write-Host "Set Tracked Replacement" -ForegroundColor Cyan
Write-Host "  $(Split-Path -Leaf $target)  [$(if ($sty.Success) { $sty.Groups[1].Value } else { 'no style' })]" -ForegroundColor Gray
Write-Host "    delete (tracked): $oldText" -ForegroundColor DarkGray
Write-Host "    insert (tracked): $Text" -ForegroundColor Green
Write-Host "    as: $Author" -ForegroundColor DarkGray
$bms = @([regex]::Matches($p, '<w:bookmarkStart\s[^>]*w:name="([^"]*)"') | ForEach-Object { $_.Groups[1].Value })
if ($bms.Count) { Write-Host "    bookmark(s) preserved: $($bms -join ', ')" -ForegroundColor DarkGray }
if ($ListOnly) { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow; return }
if (-not $Force) {
    $reply = Read-Host "Apply? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# The contiguous span of runs that carry the text.
$runs = @([regex]::Matches($p, '(?s)<w:r(?:\s[^>]*)?>.*?</w:r>') | Where-Object { $_.Value -match '<w:t[ >]' })
if ($runs.Count -eq 0) { throw "The matched paragraph has no text runs in $target. Nothing written." }
$first = $runs[0]; $last = $runs[$runs.Count - 1]
$spanStart = $first.Index
$spanEnd = $last.Index + $last.Length
$span = $p.Substring($spanStart, $spanEnd - $spanStart)
if ($span -match '<w:del\b|<w:ins\b') { throw "The matched paragraph already carries a tracked change in $target. Nothing written." }

$maxId = 0
foreach ($m in [regex]::Matches($xml, 'w:id="(\d+)"')) { $v = [int]$m.Groups[1].Value; if ($v -gt $maxId) { $maxId = $v } }
$date = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss') + 'Z'
$auth = $Author -replace '&','&amp;'

# w:t -> w:delText, or the deletion still shows as present text.
$delSpan = [regex]::Replace($span, '<w:t(\s[^>]*)?>', { param($m) '<w:delText' + $m.Groups[1].Value + '>' })
$delSpan = $delSpan -replace '</w:t>', '</w:delText>'
$delId = ++$maxId
$insId = ++$maxId
$esc = $Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
$new = '<w:del w:id="' + $delId + '" w:author="' + $auth + '" w:date="' + $date + '">' + $delSpan + '</w:del>' +
       '<w:ins w:id="' + $insId + '" w:author="' + $auth + '" w:date="' + $date + '">' +
       '<w:r><w:t xml:space="preserve">' + $esc + '</w:t></w:r></w:ins>'

$newPara = $p.Remove($spanStart, $spanEnd - $spanStart).Insert($spanStart, $new)
$newXml = $xml.Remove($pm.Index, $pm.Length).Insert($pm.Index, $newPara)

# ---- guards ----
$probeXml = $newXml
if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }
try {
    $probe = New-Object System.Xml.XmlDocument
    $probe.XmlResolver = $null
    $probe.LoadXml($probeXml)
} catch { throw "The tracked replacement produced malformed document.xml for ${target}: $($_.Exception.Message). Nothing written." }
foreach ($tag in @('<a:blip\b', '<w:drawing\b', '<w:p\b', '<w:tbl>', '<w:commentReference\b',
                   '<w:footnoteReference\b', '<w:bookmarkStart\b', '<w:bookmarkEnd\b', '<w:fldChar\b')) {
    $a = ([regex]::Matches($newXml, $tag)).Count
    $b = ([regex]::Matches($xml, $tag)).Count
    if ($a -ne $b) { throw "Protected count changed in ${target}: $tag $b -> $a. Nothing written." }
}
if ((([regex]::Matches($newXml, '<w:ins\b')).Count - ([regex]::Matches($xml, '<w:ins\b')).Count) -ne 1) {
    throw "Expected exactly one new w:ins in ${target}. Nothing written."
}
if ((([regex]::Matches($newXml, '<w:del\b')).Count - ([regex]::Matches($xml, '<w:del\b')).Count) -ne 1) {
    throw "Expected exactly one new w:del in ${target}. Nothing written."
}
if ([regex]::IsMatch($newXml, '/>\s*/>')) { throw "Stray '/>' markup in ${target}. Nothing written." }
# What a reader sees with revisions ACCEPTED must be the new text, and nothing else.
$check = [regex]::Match($newXml.Substring($pm.Index), '(?s)^<w:p\b[^>]*?>.*?</w:p>').Value
$accepted = Get-PText -P ([regex]::Replace($check, '(?s)<w:del\b.*?</w:del>', ''))
if ($accepted.Trim() -cne $Text.Trim()) {
    throw "With the revision accepted the paragraph would read '$accepted', not '$Text', in ${target}. Nothing written."
}
# And with them REJECTED it must read exactly what it read before.
$rejected = Get-PText -P ([regex]::Replace([regex]::Replace($check, '(?s)<w:ins\b.*?</w:ins>', ''), '<w:delText(\s[^>]*)?>', '<w:t$1>') -replace '</w:delText>', '</w:t>')
if ($rejected -cne $oldText) {
    throw "With the revision rejected the paragraph would read '$rejected', not '$oldText', in ${target}. Nothing written."
}

if (-not $NoBackup) {
    $bak = "$target.tracked.bak"
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
Write-Host "  Saved. Accept or reject it in Word's review pane." -ForegroundColor Green
