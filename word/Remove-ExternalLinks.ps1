<#
.SYNOPSIS
    Removes external references from a Word file: the attached template, a
    chart's external data workbook, and the relationships behind them.

.DESCRIPTION
    An author chapter arrives pointing at files on the machine it was written
    on. NWIRP 2-1 carried two kinds:

      - w:attachedTemplate -> file:///C:\Users\kwh1\Downloads\FINAL-2022_
        TechPub_template.dotx, in six of eight chapters. The path exists on no
        other machine. With w:linkStyles set it is worse than dead, because
        Word UPDATES the document's styles from that template on open and would
        overwrite the whole NIST style architecture.
      - c:externalData -> a workbook on SharePoint, behind one chart. The chart
        keeps its own cached values (c:numCache / c:strCache), so it still
        renders; it simply stops offering to refresh from a file nobody else
        can reach.

    Together these are what makes Word ask "This document contains fields that
    may refer to other files. Do you want to update the fields in this
    document?" on open. Nothing in a delivery should depend on a path outside
    the package, and nothing is left riding on Normal.dotm either: every
    paragraph style in use inherits through the xx* base styles, so removing
    these changes no rendering at all.

    Hyperlinks are left alone - they are content.

    A reference this script does not know how to neutralize safely is REPORTED
    and skipped, never guessed at. A linked image (a:blip r:link) is the case
    that matters: dropping its relationship would make the picture vanish.

.EXAMPLE
    .\Remove-ExternalLinks.ps1 -Folder "C:\...\HM-261110-v2" -ListOnly
#>
[CmdletBinding(DefaultParameterSetName = 'Folder')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'File')]
    [string]$File,

    [Parameter(Mandatory = $true, ParameterSetName = 'Folder')]
    [string]$Folder,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $r = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($r).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|docm|dotx)$' -and $_.Name -notlike '~$*' } |
        Sort-Object Name | ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }
foreach ($t in $targets) {
    $lock = Join-Path (Split-Path -Parent $t) ('~$' + (Split-Path -Leaf $t).Substring(2))
    if (Test-Path -LiteralPath $lock) { throw "$(Split-Path -Leaf $t) is open in Word. Close it first." }
}

Write-Host "Remove External Links" -ForegroundColor Cyan
Write-Host "  Drops the attached template and any chart external-data link." -ForegroundColor Gray
Write-Host "  Hyperlinks are content and are left alone. A linked image is reported, never cut." -ForegroundColor Gray
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

    $newParts = @{}          # part name -> new text
    $removed = 0
    $skipped = 0

    foreach ($relName in @($order | Where-Object { $_ -match '\.rels$' })) {
        $relText = if ($newParts.ContainsKey($relName)) { $newParts[$relName] } else { $utf8NoBom.GetString($entries[$relName]) }
        # The part this .rels belongs to: word/_rels/settings.xml.rels -> word/settings.xml
        $owner = $relName -replace '_rels/', '' -replace '\.rels$', ''

        foreach ($rm in @([regex]::Matches($relText, '<Relationship[^>]*TargetMode="External"[^>]*/>'))) {
            $typeM = [regex]::Match($rm.Value, 'Type="[^"]*/(\w+)"')
            $idM   = [regex]::Match($rm.Value, 'Id="([^"]+)"')
            $tgtM  = [regex]::Match($rm.Value, 'Target="([^"]*)"')
            if (-not $idM.Success) { continue }
            $kind = if ($typeM.Success) { $typeM.Groups[1].Value } else { '?' }
            if ($kind -eq 'hyperlink') { continue }
            $rid = $idM.Groups[1].Value
            $tgt = if ($tgtM.Success) { $tgtM.Groups[1].Value } else { '' }

            if (-not $entries.ContainsKey($owner)) {
                Write-Warning "    $kind $rid -> no owner part $owner; skipped."
                $skipped++; continue
            }
            $ownText = if ($newParts.ContainsKey($owner)) { $newParts[$owner] } else { $utf8NoBom.GetString($entries[$owner]) }

            # Only neutralize shapes known to be safe.
            $newOwn = $ownText
            $handled = $false
            $at = [regex]::Match($newOwn, '<w:attachedTemplate\b[^>]*r:id="' + [regex]::Escape($rid) + '"[^>]*/>')
            if ($at.Success) {
                $newOwn = $newOwn.Remove($at.Index, $at.Length)
                $ls = [regex]::Match($newOwn, '<w:linkStyles\b[^>]*/>')
                if ($ls.Success) { $newOwn = $newOwn.Remove($ls.Index, $ls.Length) }
                $handled = $true
            }
            if (-not $handled) {
                $ed = [regex]::Match($newOwn, '(?s)<c:externalData\b[^>]*r:id="' + [regex]::Escape($rid) + '"[^>]*>.*?</c:externalData>')
                if (-not $ed.Success) { $ed = [regex]::Match($newOwn, '<c:externalData\b[^>]*r:id="' + [regex]::Escape($rid) + '"[^>]*/>') }
                if ($ed.Success) { $newOwn = $newOwn.Remove($ed.Index, $ed.Length); $handled = $true }
            }
            if (-not $handled) {
                if ($ownText -match 'r:link="' + [regex]::Escape($rid) + '"') {
                    Write-Warning "    $kind $rid is a LINKED IMAGE or object; removing it would make the content vanish. Skipped: $tgt"
                } else {
                    Write-Warning "    $kind $rid has no handler in this script. Skipped: $tgt"
                }
                $skipped++; continue
            }

            Write-Host "    $kind $rid -> $($tgt.Substring(0, [Math]::Min(70, $tgt.Length)))" -ForegroundColor DarkGray
            $newParts[$owner] = $newOwn
            $relText = $relText.Remove($rm.Index, $rm.Length)
            $newParts[$relName] = $relText
            $removed++
        }
    }

    if ($removed -eq 0) {
        Write-Host "  No removable external references$(if ($skipped) { " ($skipped skipped)" })." -ForegroundColor DarkGray
        continue
    }
    Write-Host "  $removed reference(s) to remove$(if ($skipped) { ", $skipped skipped" })." -ForegroundColor Gray
    if ($ListOnly) { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow; continue }
    if (-not $Force) {
        $reply = Read-Host "Remove them from $(Split-Path -Leaf $target)? (y/N)"
        if ($reply -notmatch '^[yY]') { throw "Aborted." }
    }

    # -------- guards --------
    foreach ($nm in $newParts.Keys) {
        $txt = $newParts[$nm]
        $probeXml = $txt
        if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }
        try {
            $probe = New-Object System.Xml.XmlDocument
            $probe.XmlResolver = $null
            $probe.LoadXml($probeXml)
        } catch {
            throw "Removing external links produced malformed $nm in ${target}: $($_.Exception.Message). Nothing written."
        }
    }
    # No r:id or r:link may be left pointing at a relationship that is gone.
    foreach ($relName in @($newParts.Keys | Where-Object { $_ -match '\.rels$' })) {
        $owner = $relName -replace '_rels/', '' -replace '\.rels$', ''
        if (-not $newParts.ContainsKey($owner) -and -not $entries.ContainsKey($owner)) { continue }
        $ownText = if ($newParts.ContainsKey($owner)) { $newParts[$owner] } else { $utf8NoBom.GetString($entries[$owner]) }
        $have = @([regex]::Matches($newParts[$relName], 'Id="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
        foreach ($m in [regex]::Matches($ownText, 'r:(?:id|link|embed)="([^"]+)"')) {
            if ($have -notcontains $m.Groups[1].Value) {
                throw "$owner still references $($m.Groups[1].Value) but that relationship is gone in ${target}. Nothing written."
            }
        }
    }
    # Pictures and chart caches must be untouched.
    foreach ($nm in $newParts.Keys) {
        if ($nm -match '\.rels$') { continue }
        $oldText = $utf8NoBom.GetString($entries[$nm])
        foreach ($tag in @('<a:blip\b', '<c:numCache\b', '<c:strCache\b', '<c:ser>', '<c:pt\b')) {
            $a = ([regex]::Matches($newParts[$nm], $tag)).Count
            $b = ([regex]::Matches($oldText, $tag)).Count
            if ($a -ne $b) { throw "Removing external links changed $tag in $nm of ${target}: $b -> $a. Nothing written." }
        }
    }

    if (-not $NoBackup) {
        $bak = "$target.extlinks.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }

    foreach ($nm in $newParts.Keys) { $entries[$nm] = $utf8NoBom.GetBytes($newParts[$nm]) }
    $out = [System.IO.Compression.ZipFile]::Open($target, 'Update')
    foreach ($e in @($out.Entries)) { $e.Delete() }
    foreach ($name in $order) {
        $entry = $out.CreateEntry($name)
        $s = $entry.Open(); $s.Write($entries[$name], 0, $entries[$name].Length); $s.Close()
    }
    $out.Dispose()
    Write-Host "  Saved." -ForegroundColor Green
    $grand += $removed
}

Write-Host ""
Write-Host "Done. $grand external reference(s) removed." -ForegroundColor Cyan
