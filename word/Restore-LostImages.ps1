<#
.SYNOPSIS
    Puts back an image whose reference was dropped from document.xml while its
    media part and relationship are still in the package.

.DESCRIPTION
    An image can be lost without anything looking wrong. The media part stays in
    word/media/, the relationship stays in document.xml.xml.rels, and only the
    <a:blip r:embed="rIdN"> that puts it on the page is gone - so the file opens,
    every XML part parses, the text is intact, and the figure is simply absent.
    Nothing in the suite caught it: NWIRP 2-1 Ch05 shipped two orphaned images
    that way, both of them in the documented "caption paragraph that holds its
    own image" shape, where the drawing run sits first and the caption text
    follows in the same paragraph.

    This finds every relationship pointing at a media part that document.xml no
    longer references, locates the paragraph that held it in -Source, and puts
    the drawing run back into the matching paragraph here.

    MATCHING is by the tail of the paragraph's text, like
    Restore-WordParagraphs.ps1, and for the same reason: position is not stable
    once a pipeline has inserted a banner and renumbered captions. The caption
    PREFIX is expected to differ - "Figure 5.4." in the original against
    "Figure 5-4." now - so only the tail is compared, and a tail that matches
    none or several is reported and skipped rather than guessed.

    The run is reinserted at the END it occupied in the original: first run, or
    last. Nothing else in the paragraph is touched.

.PARAMETER Source
    Folder holding the pre-pipeline originals, normally <folder>\orig.

.PARAMETER TailLength
    Characters of trailing text used to match. Default 40.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Restore-LostImages.ps1 -Folder "C:\...\HM-261109-v2" -Source "C:\...\HM-261109\orig" -ListOnly
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [Parameter(Mandatory=$true)]
    [string]$Source,

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

Write-Host "Restore Lost Images" -ForegroundColor Cyan
Write-Host "  Source: $sourceDir" -ForegroundColor Cyan
Write-Host "  An image is 'lost' when its media part and relationship survive but" -ForegroundColor Cyan
Write-Host "  no a:blip in document.xml references it." -ForegroundColor Cyan
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

function Get-Part {
    param($Pkg, [string]$Name)
    if (-not $Pkg.Data.ContainsKey($Name)) { return $null }
    $t = [System.Text.Encoding]::UTF8.GetString($Pkg.Data[$Name])
    if ($t.Length -gt 0 -and $t[0] -eq [char]0xFEFF) { return $t.Substring(1) }
    return $t
}

# Paragraph text with the picture's own text excluded: a text box inside the
# drawing belongs to the picture, and including it would make two paragraphs
# that differ only in their chart labels look like the same paragraph.
function Get-ParaText {
    param([string]$P)
    $x = [regex]::Replace($P, '(?s)<w:drawing\b.*?</w:drawing>', '')
    $x = [regex]::Replace($x, '(?s)<w:pict\b.*?</w:pict>', '')
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in [regex]::Matches($x, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { [void]$sb.Append($t.Groups[1].Value) }
    return $sb.ToString()
}

function Get-Tail {
    param([string]$Text, [int]$N)
    $t = [regex]::Replace($Text, '\s+', ' ').Trim()
    if ($t.Length -le $N) { return $t }
    return $t.Substring($t.Length - $N)
}

# SELF-CLOSING FIRST: '<w:p\b[^>]*?>' also matches '<w:p .../>', so the
# open/close branch first swallows an empty paragraph and runs to the next
# paragraph's close tag.
$PARA_RX = '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'

$filesChanged = 0
foreach ($target in $targets) {
    $name = [System.IO.Path]::GetFileName($target)
    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    $doc = Get-Part -Pkg $pkg -Name 'word/document.xml'
    $rels = Get-Part -Pkg $pkg -Name 'word/_rels/document.xml.rels'
    if (-not $doc -or -not $rels) { Write-Host "  No document.xml / rels." -ForegroundColor DarkGray; continue }
    $raw = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/document.xml'])
    $hadBom = ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF)

    # Orphans: a relationship to a media part that document.xml never names.
    $orphans = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($rels, '<Relationship\b[^>]*?\bId="([^"]+)"[^>]*?\bTarget="(media/[^"]+)"')) {
        $rid = $m.Groups[1].Value
        $tgt = $m.Groups[2].Value
        if ($doc -match ('r:(?:embed|link|id)="' + [regex]::Escape($rid) + '"')) { continue }
        if (-not $pkg.Data.ContainsKey("word/$tgt")) { continue }
        [void]$orphans.Add([pscustomobject]@{ Rid = $rid; Target = $tgt })
    }
    if ($orphans.Count -eq 0) { Write-Host "  No orphaned images." -ForegroundColor Green; continue }

    $srcPath = Join-Path $sourceDir $name
    if (-not (Test-Path -LiteralPath $srcPath -PathType Leaf)) {
        Write-Host ("  {0} orphaned image(s) but no original of this name in the source folder." -f $orphans.Count) -ForegroundColor DarkYellow
        continue
    }
    $srcPkg = Read-Package -Path $srcPath
    $srcDoc = Get-Part -Pkg $srcPkg -Name 'word/document.xml'
    if (-not $srcDoc) { Write-Host "  Original has no document.xml." -ForegroundColor DarkYellow; continue }

    $curParas = @([regex]::Matches($doc, $PARA_RX))
    $plan = New-Object System.Collections.ArrayList
    $skipped = [ordered]@{}

    foreach ($o in $orphans) {
        # The paragraph that held it in the original.
        $sp = $null
        foreach ($m in [regex]::Matches($srcDoc, $PARA_RX)) {
            if ($m.Value -match ('r:embed="' + [regex]::Escape($o.Rid) + '"')) { $sp = $m.Value; break }
        }
        if (-not $sp) { $skipped["$($o.Rid): not found in the original"] = 1; continue }

        # The drawing run, and which end of the paragraph it sat at.
        $runs = @([regex]::Matches($sp, '(?s)<w:r\b[^>]*?>.*?</w:r>'))
        $drawRuns = @($runs | Where-Object { $_.Value -match ('r:embed="' + [regex]::Escape($o.Rid) + '"') })
        if ($drawRuns.Count -ne 1) { $skipped["$($o.Rid): $($drawRuns.Count) drawing run(s) in the original paragraph"] = 1; continue }
        $atFront = ($drawRuns[0].Index -eq $runs[0].Index)

        $tail = Get-Tail -Text (Get-ParaText -P $sp) -N $TailLength
        if (-not $tail) { $skipped["$($o.Rid): the original paragraph has no text to match on"] = 1; continue }

        $hits = @($curParas | Where-Object { (Get-Tail -Text (Get-ParaText -P $_.Value) -N $TailLength) -eq $tail })
        if ($hits.Count -ne 1) { $skipped["$($o.Rid): the tail matches $($hits.Count) paragraph(s) here"] = 1; continue }

        [void]$plan.Add([pscustomobject]@{
            Rid = $o.Rid; Target = $o.Target; Run = $drawRuns[0].Value
            AtFront = $atFront; Index = $hits[0].Index; Length = $hits[0].Length
            Para = $hits[0].Value; Tail = $tail
        })
    }

    foreach ($k in $skipped.Keys) { Write-Host ("  SKIPPED {0}" -f $k) -ForegroundColor DarkYellow }
    if ($plan.Count -eq 0) { Write-Host "  Nothing restorable." -ForegroundColor DarkYellow; continue }
    foreach ($p in $plan) {
        Write-Host ("    {0} -> {1} ({2} of the paragraph ending '{3}')" -f
            $p.Rid, $p.Target, $(if ($p.AtFront) { 'first run' } else { 'last run' }), $p.Tail) -ForegroundColor DarkGray
    }

    # BACK TO FRONT so each insertion leaves earlier offsets valid.
    $newDoc = $doc
    foreach ($p in @($plan | Sort-Object -Property Index -Descending)) {
        $para = $p.Para
        if ($p.AtFront) {
            $pprM = [regex]::Match($para, '(?s)<w:pPr\b[^>]*?>.*?</w:pPr>|<w:pPr\s*/>')
            $at = if ($pprM.Success) { $pprM.Index + $pprM.Length } else { [regex]::Match($para, '(?s)^<w:p\b[^>]*?>').Length }
            # Step over a bookmark that opens here, so the picture does not land
            # between a bookmarkStart and the text it is meant to wrap.
            while ($true) {
                $bm = [regex]::Match($para.Substring($at), '^<w:bookmarkStart\b[^>]*?/>')
                if (-not $bm.Success) { break }
                $at += $bm.Length
            }
            $np = $para.Insert($at, $p.Run)
        } else {
            $np = $para.Insert($para.Length - '</w:p>'.Length, $p.Run)
        }
        $newDoc = $newDoc.Remove($p.Index, $p.Length).Insert($p.Index, $np)
    }

    # --- guards ----------------------------------------------------------
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Restore produced malformed document.xml for ${name}: $($_.Exception.Message). Nothing written." }

    $bN = ([regex]::Matches($newDoc, '<a:blip\b')).Count
    $bO = ([regex]::Matches($doc, '<a:blip\b')).Count
    if ($bN -ne $bO + $plan.Count) { throw "Image count rose by $($bN - $bO) in ${name}, expected $($plan.Count). Nothing written." }
    foreach ($p in $plan) {
        if ($newDoc -notmatch ('r:embed="' + [regex]::Escape($p.Rid) + '"')) { throw "$($p.Rid) still unreferenced in ${name}. Nothing written." }
    }

    # Visible text must not move at all: only a drawing run was added.
    $textOf = { param($x) (([regex]::Matches($x, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value }) -join '') }
    if ((& $textOf $newDoc) -cne (& $textOf $doc)) { throw "Restore changed visible text in ${name}. Nothing written." }

    $pN = ([regex]::Matches($newDoc, '<w:p\b')).Count
    $pO = ([regex]::Matches($doc, '<w:p\b')).Count
    if ($pN -ne $pO) { throw "Paragraph count changed in ${name} ($pO -> $pN). Nothing written." }
    foreach ($tag in @('<w:ins\b', '<w:del\b', '<w:delText\b', '<w:bookmarkStart\b', '<w:bookmarkEnd\b', '<w:fldChar\b', '<w:fldSimple\b')) {
        $n = ([regex]::Matches($newDoc, $tag)).Count
        $o = ([regex]::Matches($doc, $tag)).Count
        if ($n -ne $o) { throw "Protected element count changed in ${name}: $tag $o -> $n. Nothing written." }
    }

    Write-Host ("  {0} image(s) restored." -f $plan.Count) -ForegroundColor DarkGray
    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        $bak = "$target.lostimg.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "lostimg-$([guid]::NewGuid()).zip")
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
