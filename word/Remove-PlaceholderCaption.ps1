<#
.SYNOPSIS
    Removes a placeholder caption paragraph that was inserted where no caption
    belongs, and keeps any cross-reference anchor it carried alive.

.DESCRIPTION
    A "figure has no caption" pass inserts a bracketed placeholder beside every
    image that has none. That is right for an image genuinely missing its
    caption and WRONG for one panel of a multi-panel figure: panels a), b) and
    c) share a single caption, so a placeholder lands between the panels and
    takes a number of its own. Everything after it then shifts by one.

    That shift is not cosmetic. Set-TableCrossReferences matches a typed prose
    mention against the number a caption currently SHOWS, and the prose was
    written against the true numbering, so a placeholder makes every REF after
    it anchor to the neighbouring figure. In NWIRP 2-1 Ch06 one placeholder
    took 6-1, pushing the three-panel map to 6-2 and the distribution plot to
    6-3, and "Figure 6-2 plots the distribution" then resolved to the map.

    Removing the paragraph is therefore only half the job: the figure numbers
    have to be rebuilt and the cross-references redone afterwards, against the
    corrected numbers.

    A cross-reference anchor (_Ref*) inside the removed paragraph is MOVED into
    the next caption paragraph rather than deleted, so an existing REF aimed
    here resolves to the caption that actually owns those panels instead of
    coming back empty. An _Hlk* bookmark is Word's inert edit-location marker
    and is simply dropped.

.PARAMETER Matching
    Regex the paragraph's whole visible text must match. Default is the
    bracketed placeholder this toolkit writes.

.EXAMPLE
    .\Remove-PlaceholderCaption.ps1 -File "...\Chapter-06.docx" -ListOnly

.EXAMPLE
    .\Remove-PlaceholderCaption.ps1 -File "...\Chapter-06.docx" -Occurrence 1 -Force
#>
[CmdletBinding(DefaultParameterSetName = 'Folder')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'File')]
    [string]$File,

    [Parameter(Mandatory = $true, ParameterSetName = 'Folder')]
    [string]$Folder,

    [string]$Style = 'Caption',

    [string]$Matching = '^\s*(Figure|Table)\b.*\[PLACEHOLDER[^\]]*\]\s*$',

    [int[]]$Occurrence,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Get-ParaText {
    param([string]$Para)
    $p = [regex]::Replace($Para, '(?s)<w:drawing\b.*?</w:drawing>|<w:pict\b.*?</w:pict>', '')
    $sb = New-Object System.Text.StringBuilder
    foreach ($m in [regex]::Matches($p, '(?s)<w:tab\s*/>|<w:noBreakHyphen\s*/>|<w:t(?:\s[^>]*)?>(.*?)</w:t>')) {
        if     ($m.Value -like '<w:tab*')           { [void]$sb.Append("`t") }
        elseif ($m.Value -like '<w:noBreakHyphen*') { [void]$sb.Append([char]0x2011) }
        else   { [void]$sb.Append($m.Groups[1].Value) }
    }
    return $sb.ToString()
}

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $r = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($r).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|docm)$' -and $_.Name -notlike '~$*' } |
        Sort-Object Name | ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }
foreach ($t in $targets) {
    $lock = Join-Path (Split-Path -Parent $t) ('~$' + (Split-Path -Leaf $t).Substring(2))
    if (Test-Path -LiteralPath $lock) { throw "$(Split-Path -Leaf $t) is open in Word. Close it first; a write would lose that work." }
}

Write-Host "Remove Placeholder Caption" -ForegroundColor Cyan
Write-Host "  Style starts with: $Style" -ForegroundColor Gray
Write-Host "  Text matches:      $Matching" -ForegroundColor Gray
Write-Host "  A _Ref anchor inside it moves to the NEXT caption; _Hlk is dropped." -ForegroundColor Gray
if ($ListOnly) { Write-Host "  -ListOnly: reporting only, writing nothing." -ForegroundColor Yellow }

$grand = 0
foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $(Split-Path -Leaf $target)..." -ForegroundColor White

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
    $bodyM = [regex]::Match($xml, '(?s)<w:body(?:\s[^>]*)?>.*</w:body>')
    if (-not $bodyM.Success) { throw "No <w:body> in $target" }
    $body = $bodyM.Value

    $paras = @([regex]::Matches($body, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'))
    $hits = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $paras.Count; $i++) {
        $p = $paras[$i].Value
        $sm = [regex]::Match($p, '<w:pStyle\s+w:val="([^"]*)"\s*/>')
        if (-not $sm.Success -or -not $sm.Groups[1].Value.StartsWith($Style)) { continue }
        $t = Get-ParaText -Para $p
        if ($t -notmatch $Matching) { continue }
        # A paragraph holding an image is never a stray caption.
        if ($p -match '<a:blip\b') {
            Write-Warning "  Skipping a match that holds an image: $($t.Trim())"
            continue
        }
        [void]$hits.Add([pscustomobject]@{ Num = $hits.Count + 1; Index = $i; Text = $t.Trim() })
    }

    if ($hits.Count -eq 0) { Write-Host "  No placeholder captions." -ForegroundColor DarkGray; continue }
    foreach ($h in $hits) { Write-Host ("    [$($h.Num)] para $($h.Index): $($h.Text)") -ForegroundColor DarkGray }

    # Select on an explicit Num assigned when the hit was recorded.
    # ArrayList.IndexOf($_) inside Where-Object returned nothing usable, so
    # $chosen came back null, its .Count rendered as an empty string, and the
    # paragraph-count guard refused the write with "Expected to remove
    #  paragraph(s)". A selector should never depend on finding an object's own
    # position in the collection it is being enumerated from.
    # Assign INSIDE each branch. '$chosen = if (...) { @(...) }' sends the array
    # through the if statement's output stream, which unrolls a one-element
    # array back to the bare element: @() is lost, $chosen.Count is $null, it
    # renders as an empty string ("Removing  of 1.") and the paragraph-count
    # guard then compares 1 -ne $null and refuses the write. Same family as the
    # 'return ,$collection' rule for functions - a collection does not survive
    # PowerShell's output stream intact.
    if ($Occurrence) { $chosen = @($hits | Where-Object { $Occurrence -contains $_.Num }) }
    else             { $chosen = @($hits) }
    if ($chosen.Count -eq 0) { Write-Host "  -Occurrence selected none. File unchanged." -ForegroundColor Yellow; continue }
    Write-Host "  Removing $($chosen.Count) of $($hits.Count)." -ForegroundColor Gray
    if ($ListOnly) { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow; continue }
    if (-not $Force) {
        $reply = Read-Host "Remove $($chosen.Count) paragraph(s) from $(Split-Path -Leaf $target)? (y/N)"
        if ($reply -notmatch '^[yY]') { throw "Aborted." }
    }

    # Back to front so recorded offsets stay valid.
    $newBody = $body
    foreach ($h in @($chosen | Sort-Object Index -Descending)) {
        $pm = $paras[$h.Index]
        $p  = $pm.Value
        # Keep any _Ref anchor by moving it into the next caption paragraph.
        $carry = ''
        foreach ($bm in [regex]::Matches($p, '<w:bookmarkStart\s[^>]*/>')) {
            $nm = [regex]::Match($bm.Value, 'w:name="([^"]+)"')
            if ($nm.Success -and $nm.Groups[1].Value -like '_Ref*') { $carry += $bm.Value }
        }
        $ids = @()
        foreach ($bm in [regex]::Matches($carry, 'w:id="(\d+)"')) { $ids += $bm.Groups[1].Value }
        $newBody = $newBody.Remove($pm.Index, $pm.Length)
        if ($carry) {
            # The next caption paragraph after the removal point.
            $after = $newBody.Substring($pm.Index)
            $nextCap = [regex]::Match($after, '(?s)<w:p\b[^>]*?>(?:(?!</w:p>).)*?<w:pStyle\s+w:val="' + [regex]::Escape($Style) + '[^"]*"\s*/>.*?</w:p>')
            if (-not $nextCap.Success) { throw "Removed paragraph carried $($ids -join ',') but no following $Style paragraph exists to hold it in ${target}. Nothing written." }
            $insAt = $pm.Index + $nextCap.Index
            $pprEnd = [regex]::Match($newBody.Substring($insAt), '(?s)<w:pPr(?:\s[^>]*)?>.*?</w:pPr>')
            if (-not $pprEnd.Success) { throw "Following caption has no pPr in ${target}. Nothing written." }
            $at = $insAt + $pprEnd.Index + $pprEnd.Length
            $ends = ''
            foreach ($id in $ids) { $ends += '<w:bookmarkEnd w:id="' + $id + '"/>' }
            $newBody = $newBody.Insert($at, $carry + $ends)
            Write-Host "    moved anchor(s) $($ids -join ', ') into the next caption" -ForegroundColor DarkGray
            # The original ends are now duplicated; drop the ones left behind.
            foreach ($id in $ids) {
                $dupRx = '<w:bookmarkEnd\s[^>]*w:id="' + [regex]::Escape($id) + '"[^>]*/>'
                $all = @([regex]::Matches($newBody, $dupRx))
                if ($all.Count -gt 1) {
                    $last = $all[$all.Count - 1]
                    $newBody = $newBody.Remove($last.Index, $last.Length)
                }
            }
        }
    }

    $new = $xml.Remove($bodyM.Index, $body.Length).Insert($bodyM.Index, $newBody)

    # -------- guards --------
    $probeXml = $new
    if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }
    try {
        $probe = New-Object System.Xml.XmlDocument
        $probe.XmlResolver = $null
        $probe.LoadXml($probeXml)
    } catch {
        throw "Removing a placeholder produced malformed document.xml for ${target}: $($_.Exception.Message). Nothing written."
    }
    foreach ($tag in @('<a:blip\b', '<w:drawing\b', '<w:pict\b', '<w:tbl>', '<w:ins\b', '<w:del\b',
                       '<w:delText\b', '<w:commentReference\b', '<w:footnoteReference\b')) {
        $a = ([regex]::Matches($new, $tag)).Count
        $b = ([regex]::Matches($xml, $tag)).Count
        if ($a -ne $b) { throw "Removing a placeholder changed $tag from $b to $a in ${target}. Nothing written." }
    }
    $pOld = ([regex]::Matches($xml, '<w:p\b')).Count
    $pNew = ([regex]::Matches($new, '<w:p\b')).Count
    if (($pOld - $pNew) -ne $chosen.Count) {
        throw "Expected to remove $($chosen.Count) paragraph(s) from ${target} but the count moved $pOld -> $pNew. Nothing written."
    }
    $bs = ([regex]::Matches($new, '<w:bookmarkStart\b')).Count
    $be = ([regex]::Matches($new, '<w:bookmarkEnd\b')).Count
    if ($bs -ne $be) { throw "Bookmarks do not pair in ${target} after removal ($bs starts, $be ends). Nothing written." }
    if ([regex]::IsMatch($new, '/>\s*/>')) { throw "Stray '/>' markup left in ${target}. Nothing written." }

    if (-not $NoBackup) {
        $bak = "$target.placeholder.bak"
        if ((Test-Path -LiteralPath $bak) -and -not $Force) { Write-Warning "  Backup exists; leaving it." }
        else { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }

    $entries['word/document.xml'] = $utf8NoBom.GetBytes($new)
    $out = [System.IO.Compression.ZipFile]::Open($target, 'Update')
    foreach ($e in @($out.Entries)) { $e.Delete() }
    foreach ($name in $order) {
        $entry = $out.CreateEntry($name)
        $s = $entry.Open(); $s.Write($entries[$name], 0, $entries[$name].Length); $s.Close()
    }
    $out.Dispose()
    Write-Host "  Saved. Rebuild the figure numbers and REDO the cross-references." -ForegroundColor Green
    $grand += $chosen.Count
}

Write-Host ""
Write-Host "Done. $grand paragraph(s) removed." -ForegroundColor Cyan
Write-Host "Figure numbers after a removal are stale: re-run Set-NISTTableCaptions and" -ForegroundColor Yellow
Write-Host "Set-TableCrossReferences, or every mention after it anchors one figure off." -ForegroundColor Yellow
