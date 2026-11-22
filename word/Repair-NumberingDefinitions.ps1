<#
.SYNOPSIS
    Repairs word/numbering.xml: removes duplicate abstractNum/num definitions
    and restores the element order the schema requires. Fixes documents whose
    bullets have vanished and whose multilevel list no longer appears in the
    gallery.

.DESCRIPTION
    CT_Numbering is a SEQUENCE, not a choice:

        numPicBullet* , abstractNum* , num* , numIdMacAtCleanup?

    so every abstractNum must precede every num. Word does not prompt to repair
    a file that breaks this. It silently discards the list definitions, which
    presents as:

      - paragraphs on ListBullet rendering with no bullet at all
      - a multilevel list style (HMList) missing from the Multilevel List
        gallery, so it cannot be selected
      - list indents that look applied in the style but do nothing

    Duplicate ids are the other half of the same failure. When two
    <w:abstractNum> elements declare the same w:abstractNumId, Word resolves
    the reference to the FIRST one, so a stale first copy wins over every
    corrected copy written after it.

    Both faults were introduced by Set-NISTStyleArchitecture.ps1, which built
    its insertion with

        [regex]::Replace($numberingXml, '(<w:num\b)', $newAbstracts + '$1', 1)

    The static [regex]::Replace has no overload taking a replacement count -
    its 4th parameter is RegexOptions - so that 1 bound to
    RegexOptions.IgnoreCase and the call replaced EVERY match. One abstractNum
    bundle went in ahead of every num element already in the file. NWIRP
    chapter 5 had six num elements at the time and ended up with six copies of
    each new abstractNum, interleaved with its num elements. That script now
    routes both insertions through Add-NumberingDefinitions, which uses the
    Regex instance method and inserts once; this script cleans up the files
    already written.

    Deduplication keeps the LAST definition of a repeated id. Successive
    writes append corrected versions, so the last copy is the current intent,
    while the first is the stale one Word was actually honouring. In chapter 5
    that rule is what replaces a Wingdings "u" diamond on ListBullet with the
    Symbol disc every later copy carried.

    Reference integrity is checked before writing: every numId referenced by
    styles.xml or document.xml must still resolve to a surviving num, and every
    num must point at a surviving abstractNum. A repair that would break a
    reference throws instead of saving.

.PARAMETER File
    A single .docx/.docm to repair.

.PARAMETER Folder
    A folder of .docx/.docm files to repair.

.PARAMETER ListOnly
    Report what is wrong and write nothing.

.PARAMETER Force
    Skip the confirmation prompt.

.PARAMETER NoBackup
    Do not write <name>.numbering.bak before saving.

.EXAMPLE
    .\Repair-NumberingDefinitions.ps1 -File "C:\...\Chapter-05.docx" -ListOnly

.EXAMPLE
    .\Repair-NumberingDefinitions.ps1 -Folder "C:\...\HM-261102" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$ListOnly,

    [switch]$Force,

    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Targets
# ---------------------------------------------------------------------------

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

Write-Host "Repair Numbering Definitions" -ForegroundColor Cyan
Write-Host "  Removes duplicate abstractNum/num ids (keeping the LAST definition)" -ForegroundColor Cyan
Write-Host "  and restores the schema order: numPicBullet, abstractNum, num." -ForegroundColor Cyan
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""
if ($ListOnly) { Write-Host "  -ListOnly: reporting only, writing nothing." -ForegroundColor Yellow }

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Read-Package {
    param([string]$Path)
    $names   = New-Object System.Collections.ArrayList
    $entries = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $stream  = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
                [void]$names.Add($e.FullName)
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
    return ,[pscustomobject]@{ Names = $names; Data = $entries }
}

function Get-TopLevelBlocks {
    # Pulls every <w:NAME>...</w:NAME> (or self-closing) block for one element
    # name. \b after the name is what stops <w:num> matching <w:numbering>,
    # <w:numPicBullet>, <w:numPr> or <w:numFmt>: the character after "num" in
    # each of those is a word character, so there is no boundary there.
    # None of these three elements nest inside one another, so a non-greedy
    # body match is safe.
    param([string]$Xml, [string]$Name)
    $rx = '(?s)<w:' + $Name + '\b[^>]*?(?:/>|>.*?</w:' + $Name + '>)'
    return ,@([regex]::Matches($Xml, $rx) | ForEach-Object { $_.Value })
}

function Get-BlockId {
    param([string]$Block, [string]$Attribute)
    $m = [regex]::Match($Block, '^<w:[A-Za-z]+\b[^>]*?\b' + $Attribute + '="([^"]+)"')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Select-LastById {
    # Keeps the LAST definition of each id, preserving first-appearance order
    # so a file with no duplicates is returned byte-identical.
    param([string[]]$Blocks, [string]$Attribute)
    $order = New-Object System.Collections.ArrayList
    $keep  = @{}
    foreach ($b in $Blocks) {
        $id = Get-BlockId -Block $b -Attribute $Attribute
        if (-not $id) { $id = '__noid__' + $order.Count }
        if (-not $keep.ContainsKey($id)) { [void]$order.Add($id) }
        $keep[$id] = $b
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($id in $order) { [void]$out.Add($keep[$id]) }
    return ,[pscustomobject]@{ Blocks = @($out); Ids = @($order) }
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$filesChanged = 0

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    if (-not $pkg.Data.ContainsKey('word/numbering.xml')) {
        Write-Host "  No word/numbering.xml. Skipping." -ForegroundColor DarkGray
        continue
    }

    $raw = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/numbering.xml'])
    # $raw.StartsWith([char]0xFEFF) is always true on .NET Framework - no
    # String.StartsWith(char) overload exists, so the char is coerced to a
    # string and the default culture-sensitive comparison treats U+FEFF as
    # ignorable. Test the character itself.
    $hadBom = ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF)
    if ($hadBom) { $raw = $raw.Substring(1) }

    $openM = [regex]::Match($raw, '(?s)^.*?<w:numbering\b[^>]*>')
    if (-not $openM.Success) { Write-Warning "  No <w:numbering> root. Skipping."; continue }
    $prologue = $openM.Value
    $closeIx  = $raw.LastIndexOf('</w:numbering>')
    if ($closeIx -lt 0) { Write-Warning "  No </w:numbering>. Skipping."; continue }
    $body = $raw.Substring($prologue.Length, $closeIx - $prologue.Length)

    $picBullets = Get-TopLevelBlocks -Xml $body -Name 'numPicBullet'
    $abstracts  = Get-TopLevelBlocks -Xml $body -Name 'abstractNum'
    $nums       = Get-TopLevelBlocks -Xml $body -Name 'num'

    # Anything the three passes did not claim. Only whitespace and the optional
    # numIdMacAtCleanup are legal here; a surprise means the parse missed
    # something and the file must be left alone rather than rebuilt from a
    # partial picture.
    $remainder = $body
    foreach ($b in ($picBullets + $abstracts + $nums)) {
        $ix = $remainder.IndexOf($b)
        if ($ix -ge 0) { $remainder = $remainder.Remove($ix, $b.Length) }
    }
    $tailExtra = ''
    $macM = [regex]::Match($remainder, '<w:numIdMacAtCleanup\b[^>]*/>')
    if ($macM.Success) {
        $tailExtra = $macM.Value
        $remainder = $remainder.Remove($macM.Index, $macM.Length)
    }
    if ($remainder.Trim().Length -gt 0) {
        $sample = $remainder.Trim()
        if ($sample.Length -gt 120) { $sample = $sample.Substring(0, 120) + '...' }
        Write-Warning "  Unrecognized top-level content in numbering.xml; leaving file alone: $sample"
        continue
    }

    $absSel = Select-LastById -Blocks $abstracts -Attribute 'w:abstractNumId'
    $numSel = Select-LastById -Blocks $nums      -Attribute 'w:numId'

    $dupAbs = $abstracts.Count - $absSel.Blocks.Count
    $dupNum = $nums.Count      - $numSel.Blocks.Count

    # Order check against the original body.
    $seq = @([regex]::Matches($body, '<w:(numPicBullet|abstractNum|num)\b') |
             ForEach-Object { $_.Groups[1].Value })
    $rank = @{ 'numPicBullet' = 0; 'abstractNum' = 1; 'num' = 2 }
    $orderBad = $false
    $prevRank = -1
    foreach ($t in $seq) {
        if ($rank[$t] -lt $prevRank) { $orderBad = $true; break }
        $prevRank = [Math]::Max($prevRank, $rank[$t])
    }

    Write-Host ("  abstractNum: {0} present, {1} distinct" -f $abstracts.Count, $absSel.Blocks.Count) -ForegroundColor DarkGray
    Write-Host ("  num:         {0} present, {1} distinct" -f $nums.Count, $numSel.Blocks.Count) -ForegroundColor DarkGray
    Write-Host ("  element order: {0}" -f $(if ($orderBad) { 'INVALID (abstractNum after num)' } else { 'valid' })) -ForegroundColor $(if ($orderBad) { 'Red' } else { 'DarkGray' })

    if ($dupAbs -gt 0) {
        $g = @{}
        foreach ($b in $abstracts) {
            $id = Get-BlockId -Block $b -Attribute 'w:abstractNumId'
            if ($id) { if ($g.ContainsKey($id)) { $g[$id]++ } else { $g[$id] = 1 } }
        }
        $dups = @($g.GetEnumerator() | Where-Object { $_.Value -gt 1 } | Sort-Object { [int]$_.Key })
        foreach ($d in $dups) {
            $copies = @($abstracts | Where-Object { (Get-BlockId -Block $_ -Attribute 'w:abstractNumId') -eq $d.Key })
            $identical = (@($copies | Select-Object -Unique).Count -eq 1)
            $note = if ($identical) { 'identical copies' } else { 'COPIES DIFFER - first copy was winning' }
            Write-Host ("    abstractNumId {0}: {1} copies ({2})" -f $d.Key, $d.Value, $note) -ForegroundColor DarkYellow
        }
    }
    if ($dupNum -gt 0) { Write-Host ("    duplicate numId definitions: {0}" -f $dupNum) -ForegroundColor DarkYellow }

    if ($dupAbs -eq 0 -and $dupNum -eq 0 -and -not $orderBad) {
        Write-Host "  Nothing to repair." -ForegroundColor Green
        continue
    }

    # --- Reference integrity, checked before anything is written -----------
    $survivingNumIds = @{}
    foreach ($id in $numSel.Ids) { $survivingNumIds[$id] = $true }
    $survivingAbsIds = @{}
    foreach ($id in $absSel.Ids) { $survivingAbsIds[$id] = $true }

    $refSources = @('word/styles.xml', 'word/document.xml')
    $brokenRefs = New-Object System.Collections.ArrayList
    foreach ($part in $refSources) {
        if (-not $pkg.Data.ContainsKey($part)) { continue }
        $t = [System.Text.Encoding]::UTF8.GetString($pkg.Data[$part]).TrimStart([char]0xFEFF)
        foreach ($m in [regex]::Matches($t, '<w:numId\b[^>]*\bw:val="(\d+)"')) {
            $v = $m.Groups[1].Value
            if ($v -eq '0') { continue }   # numId 0 means "no numbering"
            if (-not $survivingNumIds.ContainsKey($v)) { [void]$brokenRefs.Add("$part -> numId $v") }
        }
    }
    foreach ($b in $numSel.Blocks) {
        $am = [regex]::Match($b, '<w:abstractNumId\b[^>]*\bw:val="(\d+)"')
        if ($am.Success -and -not $survivingAbsIds.ContainsKey($am.Groups[1].Value)) {
            $nid = Get-BlockId -Block $b -Attribute 'w:numId'
            [void]$brokenRefs.Add("num $nid -> abstractNum $($am.Groups[1].Value)")
        }
    }
    if ($brokenRefs.Count -gt 0) {
        $uniq = @($brokenRefs | Select-Object -Unique)
        throw ("Repair would leave {0} unresolved numbering reference(s) in {1}; nothing written. First few: {2}" -f
            $uniq.Count, $target, (($uniq | Select-Object -First 5) -join '; '))
    }

    $rebuilt = $prologue +
               ($picBullets -join '') +
               ($absSel.Blocks -join '') +
               ($numSel.Blocks -join '') +
               $tailExtra +
               '</w:numbering>'

    # Must parse, and must still declare everything it declared before.
    try {
        $probe = New-Object System.Xml.XmlDocument
        $probe.XmlResolver = $null
        $probeXml = $rebuilt; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml)
    } catch {
        throw "Rebuild produced malformed numbering.xml for ${target}: $($_.Exception.Message)"
    }

    $postSeq = @([regex]::Matches($rebuilt, '<w:(numPicBullet|abstractNum|num)\b') |
                 ForEach-Object { $_.Groups[1].Value })
    $prevRank = -1
    foreach ($t in $postSeq) {
        if ($rank[$t] -lt $prevRank) { throw "Rebuild left numbering.xml out of order for ${target}." }
        $prevRank = [Math]::Max($prevRank, $rank[$t])
    }

    Write-Host ("  Repaired: removed {0} duplicate abstractNum and {1} duplicate num definition(s); order {2}" -f
        $dupAbs, $dupNum, $(if ($orderBad) { 'rebuilt' } else { 'already valid' })) -ForegroundColor Green

    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        if (-not $Force -and (Test-Path -LiteralPath "$target.numbering.bak")) {
            Write-Warning "  Backup $target.numbering.bak already exists; leaving it as-is."
        } else {
            Copy-Item -LiteralPath $target -Destination "$target.numbering.bak" -Force
        }
    }

    $pkg.Data['word/numbering.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $rebuilt } else { $rebuilt }))

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "numrepair-$([guid]::NewGuid()).zip")
    $outStream = [System.IO.File]::Create($tempZip)
    try {
        $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $pkg.Names) {
                $ne = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try { $os.Write($pkg.Data[$name], 0, $pkg.Data[$name].Length) } finally { $os.Dispose() }
            }
        } finally { $outZip.Dispose() }
    } finally { $outStream.Dispose() }

    Move-Item -LiteralPath $tempZip -Destination $target -Force
    $filesChanged++
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host ("Done. {0} file(s) repaired." -f $filesChanged) -ForegroundColor Green
