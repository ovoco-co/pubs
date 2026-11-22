<#
.SYNOPSIS
    Removes a direct w:sz/w:szCs font size from runs inside given paragraph
    styles, so the paragraph style governs the size again. Touches nothing else
    in the run.

.DESCRIPTION
    A direct run size always beats the style, so a document can have a correct
    10pt BodyText style and still render at 11pt on every page. NWIRP chapter 3
    was exactly that: BodyText inherits 10pt from xxbase, but 539 of its 611
    body runs carried a hard 11pt.

    This removes ONLY w:sz and w:szCs, and only where the value matches
    -SizePt. Bold, italic, superscript, colour, highlight, spacing and
    everything else in w:rPr survive untouched - the deliberate contrast with
    Reapply-Styles.ps1, which wipes w:rPr wholesale and is therefore excluded
    from the standard pipeline. An w:rPr left empty by the removal is dropped,
    since an empty w:rPr carries no meaning.

    Runs inside w:pPrChange are skipped: that is recorded revision history of
    what the formatting USED to be, not live formatting, and rewriting it
    falsifies the tracked change.

.PARAMETER ParagraphStyle
    Paragraph styleIds whose runs are in scope. Default: BodyText.

.PARAMETER SizePt
    The direct size to remove, in points. 11 removes w:val="22". Use -AnySize
    to remove every direct size instead.

.PARAMETER AnySize
    Remove every direct run size in scope, whatever its value, rather than one.

.PARAMETER ListOnly
    Report what would change and write nothing.

.EXAMPLE
    .\Remove-DirectRunSize.ps1 -File "C:\...\Chapter-03.docx" -SizePt 11 -ListOnly

.EXAMPLE
    .\Remove-DirectRunSize.ps1 -File "C:\...\Chapter-03.docx" -SizePt 11 -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string[]]$ParagraphStyle = @('BodyText'),

    [ValidateRange(1, 200)]
    [double]$SizePt = 0,

    [switch]$AnySize,

    [switch]$ListOnly,

    [switch]$Force,

    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not $AnySize -and $SizePt -eq 0) { throw "Give -SizePt, or -AnySize to remove every direct size." }
if ($AnySize -and $SizePt -ne 0)      { throw "-SizePt and -AnySize are mutually exclusive." }

# w:sz is in HALF-points: 11pt is w:val="22".
$halfPoints = [int]([math]::Round($SizePt * 2))

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


Write-Host "Remove Direct Run Size" -ForegroundColor Cyan
if ($AnySize) {
    Write-Host "  Removing EVERY direct w:sz/w:szCs in: $($ParagraphStyle -join ', ')" -ForegroundColor Cyan
} else {
    Write-Host ("  Removing direct w:sz/w:szCs = {0}pt (w:val=`"{1}`") in: {2}" -f $SizePt, $halfPoints, ($ParagraphStyle -join ', ')) -ForegroundColor Cyan
}
Write-Host "  Bold, italic, colour and every other run property are left alone." -ForegroundColor Cyan
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""
if ($ListOnly) { Write-Host "  -ListOnly: reporting only, writing nothing." -ForegroundColor Yellow }

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

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

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$styleSet = @{}
foreach ($s in $ParagraphStyle) { $styleSet[$s] = $true }
$filesChanged = 0

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    if (-not $pkg.Data.ContainsKey('word/document.xml')) {
        Write-Host "  No word/document.xml. Skipping." -ForegroundColor DarkGray; continue
    }
    $raw = [System.Text.Encoding]::UTF8.GetString($pkg.Data['word/document.xml'])
    # No String.StartsWith(char) overload on .NET Framework: the char is coerced
    # and U+FEFF compares as ignorable, so that test is always true. Check the
    # character itself.
    $hadBom = ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF)
    if ($hadBom) { $raw = $raw.Substring(1) }

    $szPattern = if ($AnySize) { '<w:sz(?:Cs)?\b[^>]*/>' }
                 else { '<w:sz(?:Cs)?\b[^>]*\bw:val="' + $halfPoints + '"\s*/>' }

    # Plain loops, not a MatchEvaluator: a counter incremented inside a
    # .GetNewClosure()'d scriptblock never reaches the caller, because the
    # closure gets its own module scope. The work happens and the count reads 0.
    $runsTouched = 0
    $parasTouched = 0
    $sb = New-Object System.Text.StringBuilder
    $pos = 0
    foreach ($pm in [regex]::Matches($raw, '(?s)<w:p\b[^>]*>.*?</w:p>')) {
        $para = $pm.Value
        $st = [regex]::Match($para, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        if (-not $st.Success -or -not $styleSet.ContainsKey($st.Groups[1].Value)) { continue }

        # Protect w:pPrChange: it records the OLD formatting of a tracked
        # change, so its runs are history, not live content.
        $guard = [regex]::Match($para, '(?s)<w:pPrChange\b.*?</w:pPrChange>')
        $head = $para
        $tailKept = ''
        if ($guard.Success) {
            $head = $para.Substring(0, $guard.Index)
            $tailKept = $para.Substring($guard.Index)
        }

        $before = $head
        $head = [regex]::Replace($head, $szPattern, '')
        if ($head -eq $before) { continue }

        $runsTouched += ([regex]::Matches($before, $szPattern)).Count
        # An w:rPr emptied by the removal carries nothing; drop it.
        $head = [regex]::Replace($head, '<w:rPr\s*>\s*</w:rPr>', '')
        $newPara = $head + $tailKept

        [void]$sb.Append($raw.Substring($pos, $pm.Index - $pos))
        [void]$sb.Append($newPara)
        $pos = $pm.Index + $pm.Length
        $parasTouched++
    }
    [void]$sb.Append($raw.Substring($pos))
    $newDoc = $sb.ToString()

    if ($parasTouched -eq 0) {
        Write-Host "  No matching direct sizes found. File unchanged." -ForegroundColor Green
        continue
    }
    Write-Host ("  Removed {0} direct size element(s) across {1} paragraph(s)" -f $runsTouched, $parasTouched) -ForegroundColor DarkGray

    try {
        $probe = New-Object System.Xml.XmlDocument
        $probe.XmlResolver = $null
        $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml)
    } catch {
        throw "Rewrite produced malformed document.xml for ${target}: $($_.Exception.Message)"
    }
    # Text content must be byte-identical: this changes formatting only.
    $textOf = { param($x) (([regex]::Matches($x, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') |
                 ForEach-Object { $_.Groups[1].Value }) -join '') }
    if ((& $textOf $raw) -ne (& $textOf $newDoc)) {
        throw "Rewrite altered document text for ${target}; nothing written."
    }

    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        if (-not $Force -and (Test-Path -LiteralPath "$target.runsize.bak")) {
            Write-Warning "  Backup $target.runsize.bak already exists; leaving it as-is."
        } else {
            Copy-Item -LiteralPath $target -Destination "$target.runsize.bak" -Force
        }
    }

    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "runsize-$([guid]::NewGuid()).zip")
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
Write-Host ("Done. {0} file(s) changed." -f $filesChanged) -ForegroundColor Green
