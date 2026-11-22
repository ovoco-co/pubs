<#
.SYNOPSIS
    Puts a labelled placeholder into a heading paragraph that has no text, so a
    missing title is visible instead of silent.

.DESCRIPTION
    An empty heading is not cosmetic. It is a real section boundary - the
    references scan in Test-NISTDocument.ps1 stops at one, and deleting two of
    them once merged a reference list into a nomenclature section and orphaned
    a live cross-reference bookmark - and it is invisible on the page, so it
    survives review. NWIRP 2-1 carries three: Ch04 has an Appendix_Head2 above
    an "Assessment Field / Description" table, and Ch06 has two Appendix_Head,
    one above its nomenclature and one above Table B.1 (Appendix B).

    The title itself is the author's to write. This only replaces nothing with
    something that says what is missing, in the form the review-marker check in
    Test-NISTDocument.ps1 already matches, so it is reported until it is filled
    in rather than shipped blank.

    A heading paragraph that is empty because it carries a page break, a
    bookmark or a field is left alone: it is doing a job.

.PARAMETER Style
    Heading styles to inspect. Default: the appendix and heading ladders.

.PARAMETER Text
    The placeholder. Must stay in the bracketed form the review-marker check
    looks for, or it will not be reported.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Set-EmptyHeadingText.ps1 -Folder "C:\...\HM-261109-v2" -ListOnly
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string[]]$Style = @('AppendixHead','AppendixHead2','AppendixHead3','AppendixHead4',
                         'Heading1','Heading2','Heading3','Heading4',
                         'Frontmatterhead1','Frontmatterhead2','Frontmatterhead3','Frontmatterhead4'),

    [ValidatePattern('^\[.+\]$')]
    [string]$Text = '[PLACEHOLDER - heading title]',

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$styleAlt = ($Style | ForEach-Object { [regex]::Escape($_) }) -join '|'

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

Write-Host "Set Empty Heading Text" -ForegroundColor Cyan
Write-Host "  Placeholder: $Text" -ForegroundColor Cyan
Write-Host "  An empty heading carrying a break, bookmark or field is left alone." -ForegroundColor Cyan
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

function Get-AllText {
    param([string]$Xml)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in [regex]::Matches($Xml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { [void]$sb.Append($t.Groups[1].Value) }
    return $sb.ToString()
}

$esc = [System.Security.SecurityElement]::Escape($Text)
$run = '<w:r><w:t xml:space="preserve">' + $esc + '</w:t></w:r>'

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

    # Self-closing FIRST: '<w:p\b[^>]*?>' also matches '<w:p .../>' because
    # [^>]*? absorbs the slash, and the open/close form first swallows an empty
    # paragraph through the NEXT paragraph's close tag.
    $paras = @([regex]::Matches($doc, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'))
    $hits = New-Object System.Collections.ArrayList
    foreach ($pm in $paras) {
        $p = $pm.Value
        $sm = [regex]::Match($p, '<w:pStyle\b[^>]*?\bw:val="(' + $styleAlt + ')"\s*/>')
        if (-not $sm.Success) { continue }
        if ((Get-AllText -Xml $p).Trim()) { continue }
        # Doing a job other than carrying a title.
        # A bookmarkStart does NOT disqualify: on a heading that is the target of
        # a cross-reference, that bookmark IS the anchor, and the heading having
        # no text is exactly why a REF to it renders empty. The title has to go
        # INSIDE the bookmark, which the insertion below handles.
        #
        # A fldChar or instrText does disqualify. Ch06's third 'empty appendix
        # heading' is the CLOSING paragraph of the multi-paragraph Mendeley
        # bibliography field - it carries fldCharType="end" and a
        # lastRenderedPageBreak - so it is field machinery that happens to sit on
        # a heading style, not a heading missing its title.
        #
        # w:fldSimple is the OTHER form of a field and carries its instruction as
        # an ATTRIBUTE, so neither instrText nor fldChar appears anywhere in it.
        # A heading holding only a simple field therefore has no w:t at all and
        # reads as empty: the placeholder would be inserted INSIDE a live field.
        # No heading in NWIRP 2-1 is built that way, so this is a latent case,
        # but 125 of this set's fields are fldSimple and any list of field
        # machinery that names only fldChar is half a list.
        if ($p -match '<w:br\b|<w:instrText\b|<w:fldChar\b|<w:fldSimple\b|<w:drawing\b|<w:pict\b|<w:sdt>') { continue }
        [void]$hits.Add([pscustomobject]@{ Index = $pm.Index; Length = $pm.Length; Xml = $p; Style = $sm.Groups[1].Value })
    }

    if ($hits.Count -eq 0) { Write-Host "  No empty headings." -ForegroundColor Green; continue }
    foreach ($h in $hits) { Write-Host ("    {0}" -f $h.Style) -ForegroundColor DarkGray }

    # BACK TO FRONT so each insertion leaves the earlier offsets valid.
    $newDoc = $doc
    foreach ($h in @($hits | Sort-Object -Property Index -Descending)) {
        $p = $h.Xml
        $sc = [regex]::Match($p, '(?s)^<w:p\b([^>]*?)/>$')
        if ($sc.Success) {
            # A self-closing heading paragraph cannot hold a run: reopen it.
            $np = '<w:p' + $sc.Groups[1].Value + '>' + $run + '</w:p>'
        } else {
            $pprM = [regex]::Match($p, '(?s)<w:pPr\b[^>]*?>.*?</w:pPr>|<w:pPr\s*/>')
            $at = if ($pprM.Success) { $pprM.Index + $pprM.Length } else { [regex]::Match($p, '(?s)^<w:p\b[^>]*?>').Length }
            # Step over any bookmark that opens right here, so the text lands
            # INSIDE it. A cross-reference reports what its bookmark contains, so
            # text inserted in front of the bookmarkStart leaves the REF empty.
            while ($true) {
                $bm = [regex]::Match($p.Substring($at), '^<w:bookmarkStart\b[^>]*?/>')
                if (-not $bm.Success) { break }
                $at += $bm.Length
            }
            $np = $p.Insert($at, $run)
        }
        $newDoc = $newDoc.Remove($h.Index, $h.Length).Insert($h.Index, $np)
    }

    # --- guards ----------------------------------------------------------
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Insertion produced malformed document.xml for ${name}: $($_.Exception.Message). Nothing written." }

    $before = Get-AllText -Xml $doc
    $after  = Get-AllText -Xml $newDoc
    $expected = $before.Length + ($hits.Count * $Text.Length)
    if ($after.Length -ne $expected) {
        throw "Insertion changed the text length to $($after.Length) in ${name}, expected $expected. Nothing written."
    }
    if ($after.Replace($Text, '') -cne $before) {
        throw "Insertion changed text other than the placeholders in ${name}. Nothing written."
    }
    $pN = ([regex]::Matches($newDoc, '<w:p\b')).Count
    $pO = ([regex]::Matches($doc, '<w:p\b')).Count
    if ($pN -ne $pO) { throw "Paragraph count changed in ${name} ($pO -> $pN). Nothing written." }
    foreach ($tag in @('<w:ins\b', '<w:del\b', '<w:bookmarkStart\b', '<a:blip\b', '<w:fldChar\b', '<w:fldSimple\b', '<w:pStyle\b')) {
        $n = ([regex]::Matches($newDoc, $tag)).Count
        $o = ([regex]::Matches($doc, $tag)).Count
        if ($n -ne $o) { throw "Protected element count changed in ${name}: $tag $o -> $n. Nothing written." }
    }

    Write-Host ("  {0} empty heading(s) given a placeholder." -f $hits.Count) -ForegroundColor DarkGray
    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        $bak = "$target.emptyhead.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "emptyhead-$([guid]::NewGuid()).zip")
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
Write-Host ("Done. {0} file(s) changed. The placeholders are reported by Test-NISTDocument.ps1 until they are filled in." -f $filesChanged) -ForegroundColor Green
