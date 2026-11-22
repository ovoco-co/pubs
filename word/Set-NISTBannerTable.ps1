<#
.SYNOPSIS
    Puts the NWIRP colour banner on the first page of a chapter, immediately
    before its Heading 1.

.DESCRIPTION
    The banner is a two-row, four-column table anchored to the PAGE (vertAnchor
    page, horzAnchor page, tblpX 1, tblpY 1), so it sits in the top-left corner
    of the sheet rather than in the text column. Row 1 is a dark block beside a
    white cell carrying a TITLE field; row 2 is the four-step colour ramp.

        37627E  457B9D  7DA3BA  C7D7E2

    Those four are the NWIRP banner colours from publishing/nist-color-palette.md.
    They are written as literal w:fill values with NO w:themeFill: the theme was
    abandoned, so a theme-colour reference would resolve against whatever theme
    the author's document happens to carry.

    The definition here is the AUTHORITY. It is not copied out of a reference
    document at run time, for two reasons.

    First, copying a block of XML between documents carries its bookmarks with
    it. The banner in NWIRP 2-1 Ch01 contains a stray
    <w:bookmarkStart w:id="0" w:name="_Hlk108699961"/> with no matching end, so
    cloning it into seven chapters gave each one a duplicate w:id="0" AND an
    unpaired start. (_Hlk bookmarks are Word's inert "last edit position"
    markers. Nothing REFs them, so they simply go.)

    Second, an earlier clone of that banner carried direct
    <w:rFonts w:ascii="Source Sans Pro"> into all five of its cell paragraphs.
    The cells are on TableText, which already supplies that font, so the direct
    run properties were formatting that fights its own style - the thing the
    whole pipeline exists to remove.

    w14:paraId and w14:textId are generated fresh per file and checked against
    the ids already in the document, because duplicates across a merged
    document make Word's co-authoring lose track of paragraphs.

    The TITLE field's cached result is filled in from docProps/core.xml
    <dc:title>, so the banner reads correctly without a field update. Run
    Set-DocTitle.ps1 first; if dc:title is empty the cached result is left
    empty and the field will fill in on the next F9.

.PARAMETER Replace
    Replace a page-anchored table that is already there. Without it, a file
    that already has one is reported and skipped.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Set-NISTBannerTable.ps1 -Folder "C:\...\HM-261109-v2" -ListOnly

.EXAMPLE
    .\Set-NISTBannerTable.ps1 -Folder "C:\...\HM-261109-v2" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$Replace,
    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# --- the canonical banner -------------------------------------------------
# Widths: 12235 twips is the full 8.5 inch sheet. The four columns are
# 3600 / 2878 / 2878 / 2879, and row 1's second cell spans the last three
# (8635 = 2878 + 2878 + 2879), so both rows sum to the same total. Write the
# span width as the SUM or Word re-derives the grid from the contents.
$BANNER_FILLS = @('37627E', '457B9D', '7DA3BA', 'C7D7E2')

function New-BannerXml {
    param([string[]]$ParaIds, [string[]]$TextIds, [string]$TitleCache, [bool]$HasTableGrid = $true)

    # NOT -Pid: $PID is a read-only automatic variable, so a parameter of that
    # name throws "Cannot overwrite variable Pid" at run time, not at parse time.
    function New-BannerPara {
        param([string]$ParaId, [string]$TextId, [string]$Inner = '')
        return '<w:p w14:paraId="' + $ParaId + '" w14:textId="' + $TextId + '">' +
               '<w:pPr><w:pStyle w:val="TableText"/></w:pPr>' + $Inner + '</w:p>'
    }

    $titleField =
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TITLE   \* MERGEFORMAT </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:t xml:space="preserve">' + [System.Security.SecurityElement]::Escape($TitleCache) + '</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

    $xml =
      '<w:tbl><w:tblPr>' +
      # Only when the document defines it. The banner's appearance is fully
      # specified by its own tblBorders and cell fills, so the table style is
      # decoration - but a w:tblStyle naming a style that is not in styles.xml
      # is a dangling reference, and Ch00 of NWIRP 2-1 has no TableGrid.
      $(if ($HasTableGrid) { '<w:tblStyle w:val="TableGrid"/>' } else { '' }) +
      '<w:tblpPr w:leftFromText="187" w:rightFromText="187" w:vertAnchor="page" w:horzAnchor="page" w:tblpX="1" w:tblpY="1"/>' +
      '<w:tblOverlap w:val="never"/>' +
      '<w:tblW w:w="12235" w:type="dxa"/>' +
      '<w:tblBorders>' +
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '</w:tblBorders>' +
      '<w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
      '</w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="2878"/><w:gridCol w:w="2878"/><w:gridCol w:w="2879"/></w:tblGrid>' +
      # row 1: dark block, then the title cell spanning the other three
      '<w:tr w14:paraId="' + $ParaIds[6] + '" w14:textId="' + $TextIds[6] + '">' +
      '<w:trPr><w:tblHeader/><w:trHeight w:val="1080"/></w:trPr>' +
      '<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="' + $BANNER_FILLS[0] + '"/></w:tcPr>' +
        (New-BannerPara -ParaId $ParaIds[0] -TextId $TextIds[0]) + '</w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="8635" w:type="dxa"/><w:gridSpan w:val="3"/>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>' +
        '<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="144" w:type="dxa"/>' +
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="115" w:type="dxa"/></w:tcMar>' +
        '<w:vAlign w:val="center"/></w:tcPr>' +
        (New-BannerPara -ParaId $ParaIds[1] -TextId $TextIds[1] -Inner $titleField) + '</w:tc>' +
      '</w:tr>' +
      # row 2: the four-step ramp
      '<w:tr w14:paraId="' + $ParaIds[7] + '" w14:textId="' + $TextIds[7] + '">' +
      '<w:trPr><w:trHeight w:val="1080"/></w:trPr>'

    $widths = @('3600', '2878', '2878', '2879')
    for ($c = 0; $c -lt 4; $c++) {
        $xml += '<w:tc><w:tcPr><w:tcW w:w="' + $widths[$c] + '" w:type="dxa"/>' +
                '<w:shd w:val="clear" w:color="auto" w:fill="' + $BANNER_FILLS[$c] + '"/></w:tcPr>' +
                (New-BannerPara -ParaId $ParaIds[$c + 2] -TextId $TextIds[$c + 2]) + '</w:tc>'
    }
    $xml += '</w:tr></w:tbl>'
    return $xml
}

# --- targets --------------------------------------------------------------
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

Write-Host "Set NIST Banner Table" -ForegroundColor Cyan
Write-Host ("  Fills: " + ($BANNER_FILLS -join ' ') + " (literal, no themeFill)") -ForegroundColor Cyan
Write-Host "  Inserted immediately before the first Heading 1." -ForegroundColor Cyan
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

$bannerRx = '<w:tblpPr[^>]*w:vertAnchor="page"[^>]*w:horzAnchor="page"|<w:tblpPr[^>]*w:horzAnchor="page"[^>]*w:vertAnchor="page"'
$changed = 0

foreach ($target in $targets) {
    $name = [System.IO.Path]::GetFileName($target)
    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    if (-not $pkg.Data.ContainsKey('word/document.xml')) { Write-Host "  No document.xml." -ForegroundColor DarkGray; continue }
    $raw = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/document.xml'])
    $hadBom = ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF)
    if ($hadBom) { $raw = $raw.Substring(1) }

    # The Heading 1 the banner goes above
    $h1 = [regex]::Match($raw, '(?s)<w:p\b[^>]*?>(?:(?!</w:p>).)*?<w:pStyle w:val="Heading1"\s*/>')
    if (-not $h1.Success) {
        Write-Host "  No Heading1 paragraph; nothing to anchor to. Skipped." -ForegroundColor DarkYellow
        continue
    }
    $insertAt = $h1.Index

    $existing = @([regex]::Matches($raw.Substring(0, $insertAt), $bannerRx))
    if ($existing.Count -gt 0 -and -not $Replace) {
        Write-Host ("  Already has {0} page-anchored table(s) before the Heading 1. Use -Replace to rewrite. Skipped." -f $existing.Count) -ForegroundColor DarkGray
        continue
    }

    $newDoc = $raw
    if ($existing.Count -gt 0) {
        # delete every page-anchored table that precedes the Heading 1, back to
        # front so the earlier offsets stay valid
        $tbls = @([regex]::Matches($raw, '(?s)<w:tbl>.*?</w:tbl>') |
                  Where-Object { $_.Index -lt $insertAt -and [regex]::IsMatch($_.Value, $bannerRx) })
        for ($k = $tbls.Count - 1; $k -ge 0; $k--) {
            $newDoc = $newDoc.Remove($tbls[$k].Index, $tbls[$k].Length)
        }
        $h1b = [regex]::Match($newDoc, '(?s)<w:p\b[^>]*?>(?:(?!</w:p>).)*?<w:pStyle w:val="Heading1"\s*/>')
        $insertAt = $h1b.Index
        Write-Host ("  Removed {0} existing page-anchored table(s)." -f $tbls.Count) -ForegroundColor DarkGray
    }

    # Two tables that touch merge into one in Word. If the banner would land
    # directly after a </w:tbl>, put an empty paragraph between them.
    $before = $newDoc.Substring(0, $insertAt)
    $separator = ''
    if ($before.TrimEnd().EndsWith('</w:tbl>')) {
        $separator = '<w:p/>'
        Write-Host "  Inserted a separator paragraph: the banner would otherwise merge with the table above it." -ForegroundColor DarkGray
    }

    # Fresh ids, checked against the document
    $used = @{}
    foreach ($m in [regex]::Matches($newDoc, 'w14:(?:paraId|textId)="([0-9A-Fa-f]{8})"')) { $used[$m.Groups[1].Value.ToUpper()] = $true }
    $rand = New-Object System.Random
    function New-ParaIdValue {
        param([hashtable]$Used, [System.Random]$Rand)
        do {
            $v = ('{0:X8}' -f $Rand.Next(0x10000000, [int]::MaxValue))
        } while ($Used.ContainsKey($v) -or $v -eq '00000000')
        $Used[$v] = $true
        return $v
    }
    $paraIds = @(); $textIds = @()
    # EIGHT, not seven: six cell paragraphs plus the two w:tr elements, which
    # carry a paraId of their own. With seven, index 5 served both the first
    # w:tr and the fourth ramp cell and the document came out with a duplicate.
    for ($k = 0; $k -lt 8; $k++) { $paraIds += (New-ParaIdValue -Used $used -Rand $rand) }
    for ($k = 0; $k -lt 8; $k++) { $textIds += (New-ParaIdValue -Used $used -Rand $rand) }

    # dc:title for the TITLE field's cached result
    $titleCache = ''
    if ($pkg.Data.ContainsKey('docProps/core.xml')) {
        $core = [System.Text.Encoding]::UTF8.GetString($pkg.Data['docProps/core.xml'])
        $tm = [regex]::Match($core, '(?s)<dc:title>(.*?)</dc:title>')
        if ($tm.Success) { $titleCache = $tm.Groups[1].Value }
    }
    if (-not $titleCache) {
        Write-Host "  dc:title is empty: the TITLE field goes in with no cached result. Run Set-DocTitle.ps1." -ForegroundColor DarkYellow
    }

    $hasGrid = $false
    if ($pkg.Data.ContainsKey('word/styles.xml')) {
        $sx = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/styles.xml'])
        $hasGrid = [regex]::IsMatch($sx, '<w:style\b[^>]*\bw:styleId="TableGrid"')
    }
    if (-not $hasGrid) { Write-Host "  No TableGrid style here; the banner goes in without a w:tblStyle." -ForegroundColor DarkGray }
    $banner = New-BannerXml -ParaIds $paraIds -TextIds $textIds -TitleCache $titleCache -HasTableGrid $hasGrid
    $newDoc = $newDoc.Insert($insertAt, $separator + $banner)

    # Guards
    if ($banner -match '<w:bookmarkStart') { throw "The banner definition contains a bookmark; that is the duplicate-id trap. Nothing written." }
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Insertion produced malformed document.xml for ${name}: $($_.Exception.Message)" }
    # Compare the duplicate SET before and after. Ch04 of NWIRP 2-1 already
    # carries 19 duplicate paraIds from co-authoring, so an absolute "no
    # duplicates" assertion blocks a file whose defect this script did not
    # cause and does not make worse. What matters is that the insertion adds
    # none.
    function Get-DuplicateParaIds {
        param([string]$Xml)
        $counts = @{}
        foreach ($m in [regex]::Matches($Xml, 'w14:paraId="([0-9A-Fa-f]{8})"')) {
            $v = $m.Groups[1].Value.ToUpper()
            if ($v -eq '00000000') { continue }
            $counts[$v] = 1 + [int]$counts[$v]
        }
        return ,@($counts.Keys | Where-Object { $counts[$_] -gt 1 } | Sort-Object)
    }
    $dupeBefore = Get-DuplicateParaIds -Xml $raw
    $dupeAfter  = Get-DuplicateParaIds -Xml $newDoc
    $newDupes = @($dupeAfter | Where-Object { $dupeBefore -notcontains $_ })
    if ($newDupes.Count -gt 0) { throw "Insertion created duplicate w14:paraId in ${name}: $($newDupes -join ', '). Nothing written." }
    if ($dupeBefore.Count -gt 0) {
        Write-Host ("  NOTE: {0} duplicate w14:paraId already in this file, not caused here: {1}" -f $dupeBefore.Count, (($dupeBefore | Select-Object -First 6) -join ', ')) -ForegroundColor DarkYellow
    }

    $bs = ([regex]::Matches($newDoc, '<w:bookmarkStart\b')).Count
    $be = ([regex]::Matches($newDoc, '<w:bookmarkEnd\b')).Count
    $bsOld = ([regex]::Matches($raw, '<w:bookmarkStart\b')).Count
    $beOld = ([regex]::Matches($raw, '<w:bookmarkEnd\b')).Count
    if ($bs -ne $bsOld -or $be -ne $beOld) { throw "Bookmark counts changed in ${name} ($bsOld/$beOld -> $bs/$be). Nothing written." }

    Write-Host ("  Banner inserted before the Heading 1 (title cache: '{0}')." -f $titleCache) -ForegroundColor DarkGray
    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        $bak = "$target.banner.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "banner-$([guid]::NewGuid()).zip")
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
    try { Move-Item -LiteralPath $tempZip -Destination $target -Force; $changed++; Write-Host "  Saved." -ForegroundColor Green }
    catch { Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue; Write-Host "  LOCKED (open in Word) - not written." -ForegroundColor Red }
}

Write-Host ""
Write-Host ("Done. {0} file(s) changed." -f $changed) -ForegroundColor Green
