<#
.SYNOPSIS
    Collapses runs of two or more spaces in a Word file's visible text to one,
    keeping the two spaces a NIST caption requires after its number.

.DESCRIPTION
    The double space after a sentence is a typewriter convention and the NIST
    reports do not use it in prose. One place DOES need it: a caption reads
    label + number + period + TWO spaces + title, so the whitespace that
    follows a caption's number is protected and everything else collapses.

    Only visible text is touched:

      - <w:t> content, and nothing else. Field instructions (<w:instrText>,
        <w:fldSimple w:instr>) are syntax, and collapsing a space inside
        'SEQ Figure \* ARABIC' changes what the field means.
      - Text inside <w:drawing> and <w:pict> belongs to the PICTURE - chart
        axis labels, callouts, a text box over an image - not to the
        paragraph, so those spans are excluded.
      - <w:delText> is a tracked DELETION. The characters are already gone
        from the page and editing them rewrites an author's revision history.
      - A non-breaking space (U+00A0) is deliberate - the caption label is
        joined to its number with one - and is never collapsed.

    A run of spaces may STRADDLE two runs ('end. ' + ' Next'), so the scan
    carries its state across run boundaries within the paragraph rather than
    looking at each run on its own.

.PARAMETER ProtectStyles
    Paragraph styles whose caption number keeps its two spaces. Wildcards are
    allowed. Default: Caption*

.EXAMPLE
    .\Remove-DoubleSpaces.ps1 -Folder "C:\...\HM-261110-v2" -ListOnly

.EXAMPLE
    .\Remove-DoubleSpaces.ps1 -File "C:\...\Chapter-02.docx" -Force
#>
[CmdletBinding(DefaultParameterSetName = 'Folder')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'File')]
    [string]$File,

    [Parameter(Mandatory = $true, ParameterSetName = 'Folder')]
    [string]$Folder,

    [string[]]$ProtectStyles = @('Caption*'),

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$PARTS = '^word/(document|footnotes|endnotes)\.xml$'

function Get-EditableTextNodes {
    # <w:t> nodes in document order, excluding any that sit inside a drawing or
    # a picture. Returns the inner-text span so an edit can be spliced exactly.
    param([string]$Para)
    $skip = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Para, '(?s)<w:drawing\b.*?</w:drawing>|<w:pict\b.*?</w:pict>|<w:object\b.*?</w:object>')) {
        [void]$skip.Add([pscustomobject]@{ Start = $m.Index; End = $m.Index + $m.Length })
    }
    $nodes = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Para, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) {
        $inside = $false
        foreach ($s in $skip) { if ($m.Index -ge $s.Start -and $m.Index -lt $s.End) { $inside = $true; break } }
        if ($inside) { continue }
        $g = $m.Groups[1]
        [void]$nodes.Add([pscustomobject]@{
            ElemStart = $m.Index; ElemEnd = $m.Index + $m.Length
            TextStart = $g.Index; TextEnd = $g.Index + $g.Length
            Text = $g.Value
        })
    }
    return ,$nodes
}

function Get-ProtectedRange {
    # The whitespace right after a caption's number, as an offset range into the
    # paragraph's combined text. The number ends at its FIRST terminator; a
    # cross-reference later in the title is ordinary and must not be mistaken
    # for it.
    param([string]$Text)
    $m = [regex]::Match($Text, '[.:]( +)')
    if (-not $m.Success) { return $null }
    $g = $m.Groups[1]
    return [pscustomobject]@{ Start = $g.Index; End = $g.Index + $g.Length }
}

function Collapse-Paragraph {
    param([string]$Para, [bool]$Protect)

    $nodes = Get-EditableTextNodes -Para $Para
    if ($nodes.Count -eq 0) { return [pscustomobject]@{ Xml = $Para; Count = 0 } }

    $combined = -join ($nodes | ForEach-Object { $_.Text })
    if ($combined -notmatch '  ') { return [pscustomobject]@{ Xml = $Para; Count = 0 } }

    $prot = if ($Protect) { Get-ProtectedRange -Text $combined } else { $null }

    # Walk the combined stream once. A space is dropped when the character
    # before it in the stream was also a space, which is what makes a run
    # spanning two <w:t> elements collapse correctly.
    $keep = New-Object 'System.Collections.Generic.List[bool]'
    $prevSpace = $false
    $collapsed = 0
    for ($i = 0; $i -lt $combined.Length; $i++) {
        $ch = $combined[$i]
        $isSpace = ($ch -eq ' ')
        $inProt = ($null -ne $prot -and $i -ge $prot.Start -and $i -lt $prot.End)
        if ($isSpace -and $prevSpace -and -not $inProt) {
            $keep.Add($false); $collapsed++
        } else {
            $keep.Add($true)
        }
        $prevSpace = $isSpace
    }
    if ($collapsed -eq 0) { return [pscustomobject]@{ Xml = $Para; Count = 0 } }

    # Rebuild each node's text from the keep mask, then splice BACK TO FRONT so
    # every recorded offset is still valid.
    $sb = New-Object System.Text.StringBuilder $Para
    $offset = 0
    $newTexts = @{}
    for ($n = 0; $n -lt $nodes.Count; $n++) {
        $t = $nodes[$n].Text
        $out = New-Object System.Text.StringBuilder
        for ($k = 0; $k -lt $t.Length; $k++) {
            if ($keep[$offset + $k]) { [void]$out.Append($t[$k]) }
        }
        $newTexts[$n] = $out.ToString()
        $offset += $t.Length
    }
    for ($n = $nodes.Count - 1; $n -ge 0; $n--) {
        $node = $nodes[$n]
        if ($newTexts[$n] -ceq $node.Text) { continue }
        [void]$sb.Remove($node.TextStart, $node.TextEnd - $node.TextStart)
        [void]$sb.Insert($node.TextStart, $newTexts[$n])
    }
    $out = $sb.ToString()

    # A <w:t> that still begins or ends with a space needs xml:space="preserve",
    # or Word drops it on the next save.
    $out = [regex]::Replace($out, '(?s)<w:t>( [^<]*|[^<]* )</w:t>', '<w:t xml:space="preserve">$1</w:t>')

    return [pscustomobject]@{ Xml = $out; Count = $collapsed }
}

# -------- Target list --------
$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $resolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($resolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|docm)$' -and $_.Name -notlike '~$*' } |
        Sort-Object Name | ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }

# A ~$ lock file means an author has that chapter open. A write either fails or
# loses their work, and leaves a half-written .tmp behind.
$locked = @()
foreach ($t in $targets) {
    $lock = Join-Path (Split-Path -Parent $t) ('~$' + (Split-Path -Leaf $t).Substring(2))
    if (Test-Path -LiteralPath $lock) { $locked += $t }
}
if ($locked.Count -gt 0) {
    foreach ($l in $locked) { Write-Warning "  OPEN IN WORD, skipping: $(Split-Path -Leaf $l)" }
    $targets = @($targets | Where-Object { $locked -notcontains $_ })
    if ($targets.Count -eq 0) { throw "Every target is open in Word. Nothing done." }
}

Write-Host "Remove Double Spaces" -ForegroundColor Cyan
Write-Host "  Collapses 2+ spaces to one in visible text." -ForegroundColor Gray
Write-Host "  Keeps the two spaces after a caption number on: $($ProtectStyles -join ', ')" -ForegroundColor Gray
Write-Host "  Never touches field codes, tracked deletions, picture text or U+00A0." -ForegroundColor Gray
if ($ListOnly) { Write-Host "  -ListOnly: reporting only, writing nothing." -ForegroundColor Yellow }
if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Collapse double spaces in $($targets.Count) file(s)? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

$grand = 0
foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $(Split-Path -Leaf $target)..." -ForegroundColor White

    $entries = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $order = New-Object System.Collections.ArrayList
    $srcStream = [System.IO.File]::OpenRead($target)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($srcStream, 'Read')
        foreach ($e in $zip.Entries) {
            $ms = New-Object System.IO.MemoryStream
            $s = $e.Open(); $s.CopyTo($ms); $s.Close()
            $entries[$e.FullName] = $ms.ToArray()
            [void]$order.Add($e.FullName)
        }
        $zip.Dispose()
    } finally { $srcStream.Dispose() }

    $fileCount = 0
    $edits = @{}
    foreach ($part in @($order | Where-Object { $_ -match $PARTS })) {
        $xml = $utf8NoBom.GetString($entries[$part])
        $partCount = 0
        $sb = New-Object System.Text.StringBuilder $xml
        $paras = @([regex]::Matches($xml, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'))
        for ($i = $paras.Count - 1; $i -ge 0; $i--) {
            $pm = $paras[$i]
            $para = $pm.Value
            $styM = [regex]::Match($para, '<w:pStyle\s+w:val="([^"]*)"\s*/>')
            $sty = if ($styM.Success) { $styM.Groups[1].Value } else { '' }
            $protect = $false
            foreach ($ps in $ProtectStyles) { if ($sty -like $ps) { $protect = $true; break } }
            $res = Collapse-Paragraph -Para $para -Protect $protect
            if ($res.Count -eq 0) { continue }
            $partCount += $res.Count
            [void]$sb.Remove($pm.Index, $pm.Length)
            [void]$sb.Insert($pm.Index, $res.Xml)
        }
        if ($partCount -eq 0) { continue }
        $new = $sb.ToString()

        # -------- guards --------
        $probeXml = $new
        if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }
        try {
            $probe = New-Object System.Xml.XmlDocument
            $probe.XmlResolver = $null
            $probe.LoadXml($probeXml)
        } catch {
            throw "Collapsing spaces produced malformed $part in ${target}: $($_.Exception.Message). Nothing written."
        }
        foreach ($tag in @('<a:blip\b', '<w:drawing\b', '<w:pict\b', '<w:p\b', '<w:tbl>',
                           '<w:ins\b', '<w:del\b', '<w:delText\b', '<w:delInstrText\b',
                           '<w:instrText\b', '<w:fldSimple\b', '<w:commentReference\b',
                           '<w:footnoteReference\b', '<w:bookmarkStart\b', '<w:bookmarkEnd\b',
                           '<w:noBreakHyphen\b', '<w:tab\b')) {
            $n = ([regex]::Matches($new, $tag)).Count
            $o = ([regex]::Matches($xml, $tag)).Count
            if ($n -ne $o) {
                throw "Collapsing spaces changed a protected element count in ${target} ($part): $tag $o -> $n. Nothing written."
            }
        }
        if ([regex]::IsMatch($new, '/>\s*/>')) {
            throw "Collapsing spaces left stray '/>' markup in ${target} ($part). Nothing written."
        }
        # The strong invariant: the only characters that may differ are the
        # dropped spaces. Collapse the ORIGINAL text the same way and the two
        # must agree exactly - a single lost letter fails this.
        $oldT = -join ([regex]::Matches($xml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
        $newT = -join ([regex]::Matches($new, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
        if (($oldT -replace ' {2,}', ' ') -cne ($newT -replace ' {2,}', ' ')) {
            throw "Collapsing spaces changed more than whitespace in ${target} ($part). Nothing written."
        }
        if ($newT.Length -ge $oldT.Length) {
            throw "Collapsing spaces did not shorten the text of $part in ${target} ($($oldT.Length) -> $($newT.Length)). Nothing written."
        }

        $edits[$part] = $utf8NoBom.GetBytes($new)
        Write-Host ("    {0,-22} collapsed {1}" -f $part, $partCount) -ForegroundColor DarkGray
        $fileCount += $partCount
    }

    if ($fileCount -eq 0) { Write-Host "  No double spaces. File unchanged." -ForegroundColor DarkGray; continue }
    Write-Host "  Collapsed $fileCount run(s) of spaces." -ForegroundColor Green
    $grand += $fileCount
    if ($ListOnly) { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow; continue }

    if (-not $NoBackup) {
        $bak = "$target.spaces.bak"
        if ((Test-Path -LiteralPath $bak) -and -not $Force) {
            Write-Warning "  Backup $bak already exists; leaving it as-is."
        } else {
            Copy-Item -LiteralPath $target -Destination $bak -Force
        }
    }

    foreach ($k in $edits.Keys) { $entries[$k] = $edits[$k] }
    $out = [System.IO.Compression.ZipFile]::Open($target, 'Update')
    foreach ($e in @($out.Entries)) { $e.Delete() }
    foreach ($name in $order) {
        $entry = $out.CreateEntry($name)
        $s = $entry.Open()
        $s.Write($entries[$name], 0, $entries[$name].Length)
        $s.Close()
    }
    $out.Dispose()
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done. $grand run(s) of spaces collapsed." -ForegroundColor Cyan
