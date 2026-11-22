<#
.SYNOPSIS
    Shifts a heading ladder a level deeper or shallower, so a chapter whose
    Heading 1 was used as its section level can be put right.

.DESCRIPTION
    An author chapter is often written with Heading 1 as the SECTION level and
    no chapter-title paragraph at all. NWIRP 2-1 Ch04 has eight of them -
    Introduction, Wireless Communications Outages, Concluding Remarks - and the
    consequences run further than the look:

      - Set-NISTChapterNumbering.ps1 refuses to wire the chapter, correctly,
        because the HMList chain would render "Chapter 4" through "Chapter 11".
      - With no chapter number in the document, a caption's STYLEREF has
        nothing to read, so every figure caption falls back to a bare SEQ and
        the chapter's 35 figures number 1..35 instead of 4-1..4-35.
      - The Navigation pane shows eight chapters where there is one.

    So the fix is structural and it unblocks the numbering: shift the ladder
    one level deeper, then give the chapter a title.

    ORDER MATTERS. Shifting deeper (+1) renames Heading1 to Heading2, which
    collides with the Heading2 paragraphs already there, so the DEEPEST level
    is renamed first and the shallowest last. Shifting shallower (-1) goes the
    other way. Getting this backwards merges two levels into one silently.

    Only w:pStyle in document.xml is rewritten. No style definition is touched,
    so the ladder has to exist already: a shift that would land a paragraph on
    a style the file does not define is refused, with the missing styles named,
    rather than written as a reference that renders as Normal.

.PARAMETER Shift
    Levels to move. 1 makes Heading1 into Heading2. -1 makes Heading2 into
    Heading1.

.PARAMETER Family
    Which ladder. Heading, AppendixHead or Frontmatterhead. The styleIds in
    these templates are Heading1..Heading9, AppendixHead (level 1, no digit)
    then AppendixHead2..4, and Frontmatterhead1..4, so the level-to-id mapping
    is per family rather than a suffix rule.

.PARAMETER MaxLevel
    Deepest level the family has. Default 9 for Heading, 4 for the others.

.PARAMETER Title
    Insert a chapter-title paragraph on the family's level-1 style immediately
    before the first shifted heading. Use when the shift leaves the chapter
    with no title, which is the usual reason for shifting at all.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Set-HeadingShift.ps1 -File "Ch04.docx" -Shift 1 -ListOnly

.EXAMPLE
    .\Set-HeadingShift.ps1 -File "Ch04.docx" -Shift 1 -Title "[PLACEHOLDER - chapter title]" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [Parameter(Mandatory=$true)]
    [ValidateRange(-8, 8)]
    [int]$Shift,

    [ValidateSet('Heading','AppendixHead','Frontmatterhead')]
    [string]$Family = 'Heading',

    [int]$MaxLevel = 0,

    [string]$Title,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if ($Shift -eq 0) { throw "-Shift 0 would do nothing." }
if ($MaxLevel -le 0) { $MaxLevel = if ($Family -eq 'Heading') { 9 } else { 4 } }

# Level -> styleId. AppendixHead's level 1 carries no digit, which is why this
# is a function and not "$Family + $level".
function Get-LevelId {
    param([int]$Level)
    if ($Level -lt 1 -or $Level -gt $MaxLevel) { return $null }
    if ($Family -eq 'Heading')         { return "Heading$Level" }
    if ($Family -eq 'Frontmatterhead') { return "Frontmatterhead$Level" }
    if ($Level -eq 1) { return 'AppendixHead' }
    return "AppendixHead$Level"
}

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

Write-Host "Set Heading Shift" -ForegroundColor Cyan
Write-Host ("  {0} ladder, shift {1:+0;-0} level(s)" -f $Family, $Shift) -ForegroundColor Cyan
if ($Title) { Write-Host "  Inserting a level-1 title: $Title" -ForegroundColor Cyan }
Write-Host "  Only w:pStyle is rewritten. No style definition is touched." -ForegroundColor Cyan
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

function Get-Text {
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
    $stylesXml = if ($pkg.Data.ContainsKey('word/styles.xml')) { [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/styles.xml']) } else { '' }
    $have = @{}
    foreach ($m in [regex]::Matches($stylesXml, 'w:styleId="([^"]+)"')) { $have[$m.Groups[1].Value] = $true }

    # What is actually in use, and where each would land.
    $plan = New-Object System.Collections.ArrayList
    $missing = @{}
    for ($lvl = 1; $lvl -le $MaxLevel; $lvl++) {
        $from = Get-LevelId -Level $lvl
        if (-not $from) { continue }
        $n = ([regex]::Matches($doc, '<w:pStyle\b[^>]*?\bw:val="' + [regex]::Escape($from) + '"\s*/>')).Count
        if ($n -eq 0) { continue }
        $to = Get-LevelId -Level ($lvl + $Shift)
        if (-not $to) { throw "$from would shift off the end of the $Family ladder (level $($lvl + $Shift), max $MaxLevel) in ${name}. Nothing written." }
        if (-not $have.ContainsKey($to)) { $missing[$to] = $true }
        [void]$plan.Add([pscustomobject]@{ Level = $lvl; From = $from; To = $to; Count = $n })
    }
    if ($plan.Count -eq 0) { Write-Host "  No $Family paragraphs. Nothing to shift." -ForegroundColor Green; continue }
    if ($missing.Count -gt 0) {
        throw ("The shift needs style(s) this file does not define: " + (($missing.Keys | Sort-Object) -join ', ') +
               ". Import them first (Import-WordStyles.ps1); a pStyle naming a missing style renders as Normal. Nothing written.")
    }

    foreach ($p in $plan) { Write-Host ("    {0,-18} -> {1,-18} {2} paragraph(s)" -f $p.From, $p.To, $p.Count) -ForegroundColor DarkGray }

    # DEEPEST FIRST when shifting deeper, shallowest first when shifting
    # shallower, so a rename never lands on a level still waiting its turn.
    $ordered = if ($Shift -gt 0) { @($plan | Sort-Object Level -Descending) } else { @($plan | Sort-Object Level) }

    $newDoc = $doc
    $moved = 0
    foreach ($p in $ordered) {
        $rx = '<w:pStyle\b[^>]*?\bw:val="' + [regex]::Escape($p.From) + '"\s*/>'
        $newDoc = [regex]::Replace($newDoc, $rx, ('<w:pStyle w:val="' + $p.To + '"/>'))
        $moved += $p.Count
    }

    # Insert the chapter title just before the first shifted heading.
    $titleAdded = 0
    if ($Title) {
        $lvl1 = Get-LevelId -Level 1
        if (-not $have.ContainsKey($lvl1)) { throw "Cannot insert a title: $lvl1 is not defined in ${name}." }
        $firstTo = (Get-LevelId -Level (1 + $Shift))
        $fm = [regex]::Match($newDoc, '(?s)<w:p\b[^>]*?>(?:(?!</w:p>).)*?<w:pStyle\b[^>]*?\bw:val="' + [regex]::Escape($firstTo) + '"\s*/>')
        if (-not $fm.Success) {
            Write-Warning "  Could not find the first shifted heading; title not inserted."
        } else {
            $esc = [System.Security.SecurityElement]::Escape($Title)
            $para = '<w:p><w:pPr><w:pStyle w:val="' + $lvl1 + '"/></w:pPr><w:r><w:t xml:space="preserve">' + $esc + '</w:t></w:r></w:p>'
            $newDoc = $newDoc.Insert($fm.Index, $para)
            $titleAdded = 1
            Write-Host ("    inserted {0}: {1}" -f $lvl1, $Title) -ForegroundColor DarkGray
        }
    }

    # --- guards ----------------------------------------------------------
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Shift produced malformed document.xml for ${name}: $($_.Exception.Message). Nothing written." }

    $before = Get-Text -Xml $doc
    $after = Get-Text -Xml $newDoc
    if ($titleAdded -eq 0) {
        if ($after -cne $before) { throw "Shift changed visible text in ${name}. Nothing written." }
    } else {
        # Strip the inserted title ONCE, wherever it landed, and the rest must be
        # byte-identical. String.Replace takes no count argument - passing one
        # throws "Cannot find an overload for Replace and the argument count 3"
        # - and a regex replace would need the title escaped, so cut by offset.
        $idx = $after.IndexOf($Title)
        if ($idx -lt 0) { throw "Inserted title not found in the result text for ${name}. Nothing written." }
        if ($after.Remove($idx, $Title.Length) -cne $before) {
            throw "Title insertion changed more than the inserted title in ${name}. Nothing written."
        }
    }

    $pN = ([regex]::Matches($newDoc, '<w:p\b')).Count
    $pO = ([regex]::Matches($doc, '<w:p\b')).Count
    if ($pN -ne $pO + $titleAdded) { throw "Paragraph count changed by $($pN - $pO) in ${name}, expected $titleAdded. Nothing written." }

    $sN = ([regex]::Matches($newDoc, '<w:pStyle\b')).Count
    $sO = ([regex]::Matches($doc, '<w:pStyle\b')).Count
    if ($sN -ne $sO + $titleAdded) { throw "pStyle count changed by $($sN - $sO) in ${name}, expected $titleAdded. Nothing written." }

    foreach ($tag in @('<w:ins\b', '<w:del\b', '<w:bookmarkStart\b', '<w:bookmarkEnd\b',
                       '<a:blip\b', '<w:fldChar\b', '<w:fldSimple\b', '<w:tbl>', '<w:sectPr\b')) {
        $n = ([regex]::Matches($newDoc, $tag)).Count
        $o = ([regex]::Matches($doc, $tag)).Count
        if ($n -ne $o) { throw "Protected element count changed in ${name}: $tag $o -> $n. Nothing written." }
    }
    $stray = ([regex]::Matches($newDoc, '/>\s*/>')).Count - ([regex]::Matches($doc, '/>\s*/>')).Count
    if ($stray -gt 0) { throw "Shift left $stray stray '/>' in ${name}. Nothing written." }

    Write-Host ("  {0} heading(s) shifted." -f $moved) -ForegroundColor DarkGray
    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        $bak = "$target.headingshift.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "hshift-$([guid]::NewGuid()).zip")
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
Write-Host ("Done. {0} file(s) changed. Run Set-NISTChapterNumbering.ps1 next: a chapter that now has ONE Heading 1 can be wired." -f $filesChanged) -ForegroundColor Green
