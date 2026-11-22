<#
.SYNOPSIS
    Copies named styles from a source document into targets, bringing the
    styles they depend on with them.

.DESCRIPTION
    The step the author chapter set workflow used to leave as a manual note:
    "copy any missing styles out of the reference document before step 1".
    Set-NISTStyleArchitecture.ps1 writes full definitions for the xx* family,
    Header/Footer, References, Title, MTDisplayEquation, NIST-Data,
    SidebarTable and the List ladders, but it only EDITS these:

        BodyText  CaptionTable  CaptionFigure  graphic  TableText
        TableandFigureFootnotes  FootnoteText  HMList  Quote

    and an edit pass no-ops silently when the style is not there. A chapter
    missing one gets no warning and no style: Ch04 of NWIRP 2-1 had no Quote,
    so seven block quotations stayed on Body Text and the pass that would have
    moved them reported "nothing to change". This closes that gap.

    DEPENDENCY CLOSURE. A style is useless without the styles it points at, so
    w:basedOn, w:next and w:link are followed transitively and those
    definitions come across too - a dangling basedOn is ignored by Word rather
    than reported, and the style silently falls back to Normal. A linked
    character style (BodyTextChar, QuoteChar) is part of the style, not a
    separate decision.

    Definitions are copied VERBATIM. A hand-built equivalent comes out almost
    right, differs in ways nobody notices for a week, and has to be redone.

    What it will NOT do:
      - overwrite a style the target already has, unless -Overwrite. The
        target's own version may have been corrected since.
      - import xxgraphic. It is toolkit drift and the architecture deletes it.
      - import numbering. A style's w:numPr points at a numId that means
        something different in the target document - that is the trap that put
        "Appendix %1." on List Bullet 2 - so a numbering-bearing style arrives
        and Set-NISTStyleArchitecture.ps1 then allocates it a correct
        definition. Run that afterwards.

.PARAMETER Source
    The document to copy from.

.PARAMETER Style
    Style ids to copy. Their dependencies come too.

.PARAMETER Missing
    Instead of a list, copy every style the TARGET is missing that the source
    has and the target's document.xml actually references.

.PARAMETER Overwrite
    Replace a style the target already defines.

.EXAMPLE
    .\Import-WordStyles.ps1 -Folder "C:\...\set" -Source "NIST-reference.docx" -Style Quote,BlockText

.EXAMPLE
    .\Import-WordStyles.ps1 -File "Ch04.docx" -Source "Ch03.docx" -Missing -ListOnly
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [Parameter(Mandatory=$true)]
    [string]$Source,

    [string[]]$Style,

    [switch]$Missing,

    [switch]$Overwrite,
    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if (-not $Style -and -not $Missing) { throw "Give -Style <ids> or -Missing." }
if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) { throw "Source not found: $Source" }
$sourcePath = (Resolve-Path -LiteralPath $Source).ProviderPath

$NEVER_IMPORT = @('xxgraphic')

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $resolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($resolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|docm|dotx)$' -and $_.Name -notlike '~$*' } |
        ForEach-Object { $_.FullName } | Sort-Object
}
$targets = @($targets | Where-Object { $_ -ne $sourcePath })
if ($targets.Count -eq 0) { throw "No Word files found (the source is excluded)." }

Write-Host "Import Word Styles" -ForegroundColor Cyan
Write-Host "  Source: $sourcePath" -ForegroundColor Cyan
Write-Host ("  " + $(if ($Missing) { 'Every referenced style the target is missing' } else { 'Styles: ' + ($Style -join ', ') })) -ForegroundColor Cyan
Write-Host "  basedOn / next / link are followed transitively." -ForegroundColor Cyan
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
        } finally { $stream2 = $null; $zip.Dispose() }
    } finally { $stream.Dispose() }
    return ,[pscustomobject]@{ Names = $names; Data = $entries }
}

function Get-StyleMap {
    param([string]$Xml)
    $map = @{}
    foreach ($m in [regex]::Matches($Xml, '(?s)<w:style\b[^>]*\bw:styleId="([^"]+)"[^>]*>.*?</w:style>')) {
        $map[$m.Groups[1].Value] = $m.Value
    }
    return ,$map
}

$srcPkg = Read-Package -Path $sourcePath
if (-not $srcPkg.Data.ContainsKey('word/styles.xml')) { throw "Source has no word/styles.xml." }
$srcStyles = [System.Text.Encoding]::UTF8.GetString($srcPkg.Data['word/styles.xml'])
if ($srcStyles.Length -gt 0 -and $srcStyles[0] -eq [char]0xFEFF) { $srcStyles = $srcStyles.Substring(1) }
$srcMap = Get-StyleMap -Xml $srcStyles
Write-Host ("  Source defines {0} style(s)." -f $srcMap.Count) -ForegroundColor DarkGray

$filesChanged = 0
foreach ($target in $targets) {
    $name = [System.IO.Path]::GetFileName($target)
    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    if (-not $pkg.Data.ContainsKey('word/styles.xml')) { Write-Host "  No styles.xml." -ForegroundColor DarkGray; continue }
    $xml = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/styles.xml'])
    $hadBom = ($xml.Length -gt 0 -and $xml[0] -eq [char]0xFEFF)
    if ($hadBom) { $xml = $xml.Substring(1) }
    $tgtMap = Get-StyleMap -Xml $xml

    $wanted = New-Object System.Collections.ArrayList
    if ($Missing) {
        $refd = @{}
        foreach ($n in $pkg.Names) {
            if ($n -notmatch '^word/[^/]+\.xml$') { continue }
            if ($n -match '^word/(styles|numbering|settings|fontTable|theme\d*|webSettings|people)\.xml$') { continue }
            $part = [System.Text.Encoding]::UTF8.GetString($pkg.Data[$n])
            foreach ($m in [regex]::Matches($part, '<w:(?:pStyle|rStyle|tblStyle)\b[^>]*\bw:val="([^"]+)"')) { $refd[$m.Groups[1].Value] = $true }
        }
        foreach ($k in $refd.Keys) {
            if ($tgtMap.ContainsKey($k)) { continue }
            if ($srcMap.ContainsKey($k)) { [void]$wanted.Add($k) }
        }
    } else {
        foreach ($k in $Style) { [void]$wanted.Add($k) }
    }

    # transitive closure over basedOn / next / link
    $need = New-Object System.Collections.ArrayList
    $seen = @{}
    $queue = New-Object System.Collections.Queue
    foreach ($k in $wanted) { $queue.Enqueue($k) }
    $absent = @()
    while ($queue.Count -gt 0) {
        $id = $queue.Dequeue()
        if ($seen.ContainsKey($id)) { continue }
        $seen[$id] = $true
        if ($NEVER_IMPORT -contains $id) { continue }
        if (-not $srcMap.ContainsKey($id)) { $absent += $id; continue }
        if (-not $tgtMap.ContainsKey($id) -or $Overwrite) { [void]$need.Add($id) }
        foreach ($rel in @('basedOn', 'next', 'link')) {
            $rm = [regex]::Match($srcMap[$id], '<w:' + $rel + '\b[^>]*\bw:val="([^"]+)"')
            if ($rm.Success) { $queue.Enqueue($rm.Groups[1].Value) }
        }
    }
    if ($absent.Count -gt 0) {
        Write-Host ("  Not in the source, so not imported: {0}" -f (($absent | Sort-Object -Unique) -join ', ')) -ForegroundColor DarkYellow
    }
    if ($need.Count -eq 0) { Write-Host "  Nothing to import." -ForegroundColor Green; continue }

    $added = @(); $replaced = @()
    foreach ($id in $need) {
        $def = $srcMap[$id]
        if ($tgtMap.ContainsKey($id)) {
            $em = [regex]::Match($xml, '(?s)<w:style\b[^>]*\bw:styleId="' + [regex]::Escape($id) + '"[^>]*>.*?</w:style>')
            if (-not $em.Success) { continue }
            $xml = $xml.Remove($em.Index, $em.Length).Insert($em.Index, $def)
            $replaced += $id
        } else {
            $xml = [regex]::Replace($xml, '</w:styles>', ([System.Text.RegularExpressions.Regex]::Escape($def) -replace '\\(.)', '$1') + '</w:styles>', 1)
            $added += $id
        }
        $tgtMap[$id] = $def
    }
    if ($added.Count -gt 0)    { Write-Host ("  Added:    {0}" -f ($added -join ', ')) -ForegroundColor DarkGray }
    if ($replaced.Count -gt 0) { Write-Host ("  Replaced: {0}" -f ($replaced -join ', ')) -ForegroundColor DarkGray }

    # guards
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $xml; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Import produced malformed styles.xml for ${name}: $($_.Exception.Message)" }
    $finalMap = Get-StyleMap -Xml $xml
    $dangling = @()
    foreach ($sid in $finalMap.Keys) {
        foreach ($rel in @('basedOn', 'link', 'next')) {
            $rm = [regex]::Match($finalMap[$sid], '<w:' + $rel + '\b[^>]*\bw:val="([^"]+)"')
            if ($rm.Success -and -not $finalMap.ContainsKey($rm.Groups[1].Value)) { $dangling += "$sid ($rel -> $($rm.Groups[1].Value))" }
        }
    }
    if ($dangling.Count -gt 0) { throw "Import would leave a dangling reference in ${name}: $($dangling -join '; '). Nothing written." }
    $dupes = @($finalMap.Keys | Where-Object { ([regex]::Matches($xml, 'w:styleId="' + [regex]::Escape($_) + '"')).Count -gt 1 })
    if ($dupes.Count -gt 0) { throw "Duplicate styleId after import in ${name}: $($dupes -join ', '). Nothing written." }

    if ($ListOnly) { continue }
    if (-not $NoBackup) {
        $bak = "$target.styleimport.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/styles.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $xml } else { $xml }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "styleimport-$([guid]::NewGuid()).zip")
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
Write-Host ("Done. {0} file(s) changed. Run Set-NISTStyleArchitecture.ps1 next: an imported style's numPr points at a numId that means something else here." -f $filesChanged) -ForegroundColor Green
