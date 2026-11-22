<#
.SYNOPSIS
    Puts paragraphs back from the pre-pipeline original when a pass has damaged
    them, matching them by their text rather than by position.

.DESCRIPTION
    The author chapter set workflow keeps a copy of the input in <folder>\orig\
    so that every claim about what changed can be measured against it. This is
    the other thing that copy is for: when a pass turns out to have damaged
    specific paragraphs, the original of each one is still on disk, and putting
    it back is better than either rebuilding the file through the whole
    pipeline or hand-editing XML.

    What it was written for: a caption normaliser measured its text offsets
    from w:t elements alone, which left out <w:noBreakHyphen/>. In a caption in
    correct NIST form - "Table " + STYLEREF + noBreakHyphen + SEQ + ".  Title" -
    that made every offset past the hyphen one short, so the rewritten prefix
    landed over the number and "Table 3-1. The infrastructure..." came out as
    "Table 31.  -The infrastructure...". Twelve captions across two chapters.

    MATCHING is by the tail of the paragraph's text, not by index: the pipeline
    inserts a banner table and may add or remove paragraphs, so position is not
    stable. A paragraph is restored only when exactly one paragraph in the
    source has the same tail, and a tail that matches several, or none, is
    reported and skipped.

    REFUSED, rather than guessed:
      - a source paragraph naming a bookmark the target no longer has. That
        bookmark was removed deliberately (the inert _Hlk and Google
        navigation markers), and bringing the start back without its end
        leaves the document with an unpaired bookmark.
      - a source paragraph naming a style the target does not define.

    The restored paragraph is the ORIGINAL, so it comes back with whatever
    direct formatting it had. Re-run the idempotent passes afterwards -
    Reapply-Styles.ps1, Remove-DirectParagraphFormat.ps1, Set-ContentStyles.ps1 -
    which is why they are idempotent.

.PARAMETER Source
    The folder holding the originals, normally <folder>\orig.

.PARAMETER Style
    Only consider paragraphs whose pStyle matches this regex, in EITHER file.
    Narrow this: 'Caption\w*' restores captions and nothing else.

.PARAMETER SourceContains
    Only consider a paragraph when its ORIGINAL xml matches this regex. Use it
    to restore exactly the paragraphs a known fault could have damaged, instead
    of every paragraph that has since changed for a good reason. For the offset
    fault above that is '<w:noBreakHyphen\s*/>': a paragraph with no such node
    had correct offsets and its normalisation should stand.

.PARAMETER TailLength
    Characters of trailing text used to match. Default 40.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Restore-WordParagraphs.ps1 -Folder "C:\...\HM-261109-v2" `
        -Source "C:\...\HM-261109\orig" -Style 'Caption\w*' -ListOnly
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [Parameter(Mandatory=$true)]
    [string]$Source,

    [Parameter(Mandatory=$true)]
    [string]$Style,

    [string]$SourceContains,

    [int]$TailLength = 40,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if (-not (Test-Path -LiteralPath $Source -PathType Container)) { throw "Source folder not found: $Source" }
$sourceDir = (Resolve-Path -LiteralPath $Source).ProviderPath

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

Write-Host "Restore Word Paragraphs" -ForegroundColor Cyan
Write-Host "  Source:  $sourceDir" -ForegroundColor Cyan
Write-Host "  Style:   $Style" -ForegroundColor Cyan
if ($SourceContains) { Write-Host "  Only where the original matches: $SourceContains" -ForegroundColor Cyan }
Write-Host "  Matched on the last $TailLength characters of text, not on position." -ForegroundColor Cyan
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

function Read-Doc {
    param([string]$Path)
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open,
              [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $e = $zip.GetEntry('word/document.xml')
            if (-not $e) { return $null }
            $r = New-Object System.IO.StreamReader($e.Open())
            try { $t = $r.ReadToEnd() } finally { $r.Dispose() }
            if ($t.Length -gt 0 -and $t[0] -eq [char]0xFEFF) { $t = $t.Substring(1) }
            return $t
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
}

# The full text model: every element that puts a character on the page, not
# just w:t. Leaving out noBreakHyphen is what caused the damage this script
# exists to undo, so the matcher does not repeat it.
function Get-FullText {
    param([string]$Block)
    $sb = New-Object System.Text.StringBuilder
    $rx = '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>|<w:noBreakHyphen\s*/>|<w:softHyphen\s*/>|<w:tab\s*/>|<w:br(?:\s[^>]*)?/>'
    foreach ($m in [regex]::Matches($Block, $rx)) {
        $v = $m.Value
        # '<w:tab/>'.StartsWith('<w:t') is TRUE. Match the element, '<w:t>' or
        # '<w:t ', so a tab is not read as a text run with empty content.
        if ($v -match '^<w:t[ >]') { [void]$sb.Append($m.Groups[1].Value) }
        elseif ($v.StartsWith('<w:noBreakHyphen')) { [void]$sb.Append([char]0x2011) }
        elseif ($v.StartsWith('<w:tab')) { [void]$sb.Append("`t") }
        elseif ($v.StartsWith('<w:br')) { [void]$sb.Append("`n") }
    }
    return $sb.ToString()
}

function Get-Tail {
    param([string]$Text, [int]$N)
    $t = [regex]::Replace($Text, '\s+', ' ').Trim()
    if ($t.Length -le $N) { return $t }
    return $t.Substring($t.Length - $N)
}

$filesChanged = 0
foreach ($target in $targets) {
    $name = [System.IO.Path]::GetFileName($target)
    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Yellow

    $srcPath = Join-Path $sourceDir $name
    if (-not (Test-Path -LiteralPath $srcPath -PathType Leaf)) {
        Write-Host "  No original of this name in the source folder. Skipped." -ForegroundColor DarkGray
        continue
    }

    $srcDoc = Read-Doc -Path $srcPath
    $pkg = Read-Package -Path $target
    if (-not $pkg.Data.ContainsKey('word/document.xml')) { continue }
    $raw = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/document.xml'])
    $hadBom = ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF)
    if ($hadBom) { $raw = $raw.Substring(1) }

    $haveStyles = @{}
    if ($pkg.Data.ContainsKey('word/styles.xml')) {
        $sx = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/styles.xml'])
        foreach ($m in [regex]::Matches($sx, 'w:styleId="([^"]+)"')) { $haveStyles[$m.Groups[1].Value] = $true }
    }
    $haveBookmarks = @{}
    foreach ($m in [regex]::Matches($raw, '<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"')) { $haveBookmarks[$m.Groups[1].Value] = $true }

    # index the source's candidate paragraphs by text tail
    $srcByTail = @{}
    foreach ($pm in [regex]::Matches($srcDoc, '(?s)<w:p\b[^>]*?>.*?</w:p>')) {
        $sm = [regex]::Match($pm.Value, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        if (-not $sm.Success -or $sm.Groups[1].Value -notmatch "^$Style$") { continue }
        if ($SourceContains -and $pm.Value -notmatch $SourceContains) { continue }
        $tail = Get-Tail -Text (Get-FullText -Block $pm.Value) -N $TailLength
        if (-not $tail) { continue }
        if (-not $srcByTail.ContainsKey($tail)) { $srcByTail[$tail] = New-Object System.Collections.ArrayList }
        [void]$srcByTail[$tail].Add($pm.Value)
    }

    $restored = 0; $same = 0; $skipped = [ordered]@{}
    $sb = New-Object System.Text.StringBuilder
    $pos = 0

    foreach ($pm in [regex]::Matches($raw, '(?s)<w:p\b[^>]*?>.*?</w:p>')) {
        $p = $pm.Value
        $sm = [regex]::Match($p, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        if (-not $sm.Success -or $sm.Groups[1].Value -notmatch "^$Style$") { continue }
        $curStyle = $sm.Groups[1].Value
        $text = Get-FullText -Block $p
        $tail = Get-Tail -Text $text -N $TailLength
        if (-not $tail) { continue }
        if (-not $srcByTail.ContainsKey($tail)) {
            $skipped['no paragraph in the original has this tail'] = 1 + [int]$skipped['no paragraph in the original has this tail']
            continue
        }
        if ($srcByTail[$tail].Count -ne 1) {
            $skipped["the tail matches $($srcByTail[$tail].Count) paragraphs in the original"] =
                1 + [int]$skipped["the tail matches $($srcByTail[$tail].Count) paragraphs in the original"]
            continue
        }
        $orig = $srcByTail[$tail][0]
        if ((Get-FullText -Block $orig) -ceq $text) { $same++; continue }

        # a bookmark in the original that is no longer in this document was
        # removed on purpose; restoring the start alone leaves it unpaired
        $lost = @()
        foreach ($bm in [regex]::Matches($orig, '<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"')) {
            if (-not $haveBookmarks.ContainsKey($bm.Groups[1].Value)) { $lost += $bm.Groups[1].Value }
        }
        if ($lost.Count -gt 0) {
            $skipped["names a bookmark this document no longer has ($($lost -join ', '))"] =
                1 + [int]$skipped["names a bookmark this document no longer has ($($lost -join ', '))"]
            continue
        }
        # The WHOLE element, not up to the attribute quote. A pattern that
        # stops at the closing quote and is replaced with a full element leaves
        # the original's "/>" behind as stray character data inside w:pPr:
        # <w:pStyle w:val="CaptionFigure"/>/>. That is WELL-FORMED XML - those
        # two characters are legal text - so a parse check passes it and only a
        # search for /> followed by /> finds it. It happened to 22 paragraphs.
        $osm = [regex]::Match($orig, '<w:pStyle\b[^>]*?\bw:val="([^"]+)"[^>]*?/>')
        if ($osm.Success -and -not $haveStyles.ContainsKey($osm.Groups[1].Value)) {
            $skipped["the original is on $($osm.Groups[1].Value), which this document does not define"] =
                1 + [int]$skipped["the original is on $($osm.Groups[1].Value), which this document does not define"]
            continue
        }

        # Keep the style the pipeline assigned: that part was right. Only the
        # content comes back.
        $new = $orig
        if ($osm.Success) {
            $new = $new.Remove($osm.Index, $osm.Length).Insert($osm.Index, "<w:pStyle w:val=`"$curStyle`"/>")
        }
        Write-Host ("    restored: {0}" -f (Get-Tail -Text (Get-FullText -Block $new) -N 58)) -ForegroundColor DarkGray
        [void]$sb.Append($raw.Substring($pos, $pm.Index - $pos))
        [void]$sb.Append($new)
        $pos = $pm.Index + $pm.Length
        $restored++
    }
    [void]$sb.Append($raw.Substring($pos))
    $newDoc = $sb.ToString()

    foreach ($k in $skipped.Keys) { Write-Host ("  SKIPPED x{0}: {1}" -f $skipped[$k], $k) -ForegroundColor DarkYellow }
    Write-Host ("  {0} restored, {1} already identical" -f $restored, $same) -ForegroundColor DarkGray
    if ($restored -eq 0) { continue }

    # guards
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Restore produced malformed document.xml for ${name}: $($_.Exception.Message)" }
    $bsN = ([regex]::Matches($newDoc, '<w:bookmarkStart\b')).Count
    $bsO = ([regex]::Matches($raw, '<w:bookmarkStart\b')).Count
    $beN = ([regex]::Matches($newDoc, '<w:bookmarkEnd\b')).Count
    $beO = ([regex]::Matches($raw, '<w:bookmarkEnd\b')).Count
    if ($bsN -ne $bsO -or $beN -ne $beO) { throw "Bookmark counts changed in ${name} ($bsO/$beO -> $bsN/$beN). Nothing written." }
    $pN = ([regex]::Matches($newDoc, '<w:p\b')).Count
    $pO = ([regex]::Matches($raw, '<w:p\b')).Count
    if ($pN -ne $pO) { throw "Paragraph count changed in ${name} ($pO -> $pN). Nothing written." }
    # A replacement built from a partial element match leaves "/>" loose in the
    # markup. It parses, so this is the only thing that catches it.
    $strayClose = ([regex]::Matches($newDoc, '/>\s*/>')).Count - ([regex]::Matches($raw, '/>\s*/>')).Count
    if ($strayClose -gt 0) { throw "Restore left $strayClose stray '/>' in ${name}. Nothing written." }

    if ($ListOnly) { continue }
    if (-not $NoBackup) {
        $bak = "$target.restore.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "restore-$([guid]::NewGuid()).zip")
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
