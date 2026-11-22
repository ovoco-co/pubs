<#
.SYNOPSIS
    Converts a non-breaking space to an ordinary space where it is only acting
    as a word or sentence separator, and removes one that is doing nothing.

.DESCRIPTION
    A non-breaking space is deliberate in three places and accidental
    everywhere else. The deliberate ones are KEPT:

      - A caption's label joined to its number, "Figure" + U+00A0 + the number,
        which is the canonical NIST caption form and must never break across a
        line. 113 of them in NWIRP 2-1.
      - A value joined to its unit - "20 %", "5 km", "35 m/s", "230 kV" - which
        the NIST Guide to SI Units asks for. Recognised either by a digit on one
        side or by a unit from -Units on the right.
      - Anything else the caller protects with -KeepMatching.

    The accidental ones come in from an HTML or PDF paste and are invisible in
    Word until a line wraps oddly. Three shapes are handled:

      - Between two words, or after a sentence's full stop: becomes an ordinary
        space. "NIST[NBSP]is guided by", "(FEMA 2017).[NBSP]Also citing".
      - Leading or trailing in a paragraph: removed outright. It indents the
        first line or pads the last, and no style can explain either.
      - A paragraph whose entire content is non-breaking spaces: emptied. One
        sits in Ch05's reference list, where it reads as a stray indented line.

    Spelled-out values are why a unit list is needed rather than a rule about
    letters: "five[NBSP]km" has letters on both sides and must stay joined,
    while "NIST[NBSP]is" must not.

.EXAMPLE
    .\Repair-NonBreakingSpaces.ps1 -Folder "C:\...\HM-261110-v2" -ListOnly
#>
[CmdletBinding(DefaultParameterSetName = 'Folder')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'File')]
    [string]$File,

    [Parameter(Mandatory = $true, ParameterSetName = 'Folder')]
    [string]$Folder,

    [string[]]$Units = @('%','km','m','mm','cm','ft','in','mi','s','h','min','ms','kV','V','MW','kW','kWh','mph','kph','Hz','kg','g','lb','C','F','km2','m2','ha','pt','px','deg'),

    [string]$KeepMatching,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$NB = [char]0x00A0
$PARTS = '^word/(document|footnotes|endnotes)\.xml$'
$unitRx = '^(?:' + (($Units | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')(?![A-Za-z])'

function Get-TextNodes {
    # <w:t> nodes outside any drawing or picture, with their inner-text spans.
    param([string]$Para)
    $skip = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Para, '(?s)<w:drawing\b.*?</w:drawing>|<w:pict\b.*?</w:pict>|<w:object\b.*?</w:object>')) {
        [void]$skip.Add([pscustomobject]@{ S = $m.Index; E = $m.Index + $m.Length })
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Para, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) {
        $in = $false
        foreach ($s in $skip) { if ($m.Index -ge $s.S -and $m.Index -lt $s.E) { $in = $true; break } }
        if ($in) { continue }
        $g = $m.Groups[1]
        [void]$out.Add([pscustomobject]@{ Start = $g.Index; Len = $g.Length; Text = $g.Value })
    }
    return ,$out
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
    if (Test-Path -LiteralPath $lock) { throw "$(Split-Path -Leaf $t) is open in Word. Close it first." }
}

Write-Host "Repair Non-Breaking Spaces" -ForegroundColor Cyan
Write-Host "  Kept: a caption's label-to-number, and a value joined to its unit." -ForegroundColor Gray
Write-Host "  Converted to an ordinary space: a word or sentence separator." -ForegroundColor Gray
Write-Host "  Removed: leading, trailing, or a paragraph that is nothing else." -ForegroundColor Gray
if ($ListOnly) { Write-Host "  -ListOnly: reporting only, writing nothing." -ForegroundColor Yellow }

$grandConv = 0; $grandDel = 0; $grandKept = 0
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

    $edits = @{}
    $fConv = 0; $fDel = 0; $fKept = 0
    $samples = New-Object System.Collections.ArrayList

    foreach ($part in @($order | Where-Object { $_ -match $PARTS })) {
        $xml = $utf8NoBom.GetString($entries[$part])
        $sb = New-Object System.Text.StringBuilder $xml
        $paras = @([regex]::Matches($xml, '(?s)<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'))
        $pConv = 0; $pDel = 0
        for ($i = $paras.Count - 1; $i -ge 0; $i--) {
            $pm = $paras[$i]; $p = $pm.Value
            if ($p -notmatch [regex]::Escape($NB)) { continue }
            $isCap = $p -match '<w:pStyle\s+w:val="Caption\w*"\s*/>'
            $nodes = Get-TextNodes -Para $p
            if ($nodes.Count -eq 0) { continue }
            $combined = -join ($nodes | ForEach-Object { $_.Text })
            if ($combined -notmatch [regex]::Escape($NB)) { continue }

            # Decide each NBSP once, over the combined text.
            $action = @{}
            $allNb = ($combined -replace [regex]::Escape($NB), '') -replace '\s', ''
            for ($k = 0; $k -lt $combined.Length; $k++) {
                if ($combined[$k] -ne $NB) { continue }
                $before = $combined.Substring(0, $k)
                $after  = $combined.Substring($k + 1)
                if ($allNb -eq '') { $action[$k] = 'del'; continue }          # nothing but spaces
                if ($before.Trim() -eq '') { $action[$k] = 'del'; continue }  # leading
                if ($after.Trim() -eq '')  { $action[$k] = 'del'; continue }  # trailing
                if ($KeepMatching -and $combined -match $KeepMatching) { $action[$k] = 'keep'; continue }
                if ($isCap -and $before -match '^\s*(Figure|Table)\s*$') { $action[$k] = 'keep'; continue }
                $b1 = $before.Substring($before.Length - 1, 1)
                if ($b1 -match '\d') { $action[$k] = 'keep'; continue }
                if ($after -match '^\d') { $action[$k] = 'keep'; continue }
                if ($after -match $unitRx) { $action[$k] = 'keep'; continue }
                $action[$k] = 'conv'
            }
            $conv = @($action.Keys | Where-Object { $action[$_] -eq 'conv' })
            $del  = @($action.Keys | Where-Object { $action[$_] -eq 'del' })
            $fKept += @($action.Keys | Where-Object { $action[$_] -eq 'keep' }).Count
            if ($conv.Count -eq 0 -and $del.Count -eq 0) { continue }

            foreach ($k in $conv) {
                if ($samples.Count -lt 8) {
                    $lo = [Math]::Max(0, $k - 30); $hi = [Math]::Min($combined.Length, $k + 24)
                    [void]$samples.Add('conv: ...' + $combined.Substring($lo, $k - $lo) + '[NBSP]' + $combined.Substring($k + 1, $hi - $k - 1) + '...')
                }
            }
            foreach ($k in $del) { if ($samples.Count -lt 8) { [void]$samples.Add('del : a leading, trailing or lone NBSP') } }

            # Rebuild each node from the decisions.
            $offset = 0
            $newTexts = @{}
            for ($n = 0; $n -lt $nodes.Count; $n++) {
                $t = $nodes[$n].Text
                $outSb = New-Object System.Text.StringBuilder
                for ($c = 0; $c -lt $t.Length; $c++) {
                    $gi = $offset + $c
                    if ($action.ContainsKey($gi)) {
                        if     ($action[$gi] -eq 'conv') { [void]$outSb.Append(' ') }
                        elseif ($action[$gi] -eq 'del')  { }
                        else   { [void]$outSb.Append($t[$c]) }
                    } else { [void]$outSb.Append($t[$c]) }
                }
                $newTexts[$n] = $outSb.ToString()
                $offset += $t.Length
            }
            $newPara = $p
            for ($n = $nodes.Count - 1; $n -ge 0; $n--) {
                if ($newTexts[$n] -ceq $nodes[$n].Text) { continue }
                $newPara = $newPara.Remove($nodes[$n].Start, $nodes[$n].Len).Insert($nodes[$n].Start, $newTexts[$n])
            }
            $newPara = [regex]::Replace($newPara, '(?s)<w:t>( [^<]*|[^<]* )</w:t>', '<w:t xml:space="preserve">$1</w:t>')
            [void]$sb.Remove($pm.Index, $pm.Length)
            [void]$sb.Insert($pm.Index, $newPara)
            $pConv += $conv.Count; $pDel += $del.Count
        }
        if ($pConv -eq 0 -and $pDel -eq 0) { continue }
        $new = $sb.ToString()

        # ---- guards ----
        $probeXml = $new
        if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }
        try {
            $probe = New-Object System.Xml.XmlDocument
            $probe.XmlResolver = $null
            $probe.LoadXml($probeXml)
        } catch { throw "Repairing non-breaking spaces produced malformed $part in ${target}: $($_.Exception.Message). Nothing written." }
        foreach ($tag in @('<a:blip\b', '<w:drawing\b', '<w:pict\b', '<w:p\b', '<w:tbl>', '<w:ins\b', '<w:del\b',
                           '<w:delText\b', '<w:instrText\b', '<w:fldSimple\b', '<w:fldChar\b',
                           '<w:commentReference\b', '<w:footnoteReference\b', '<w:noBreakHyphen\b', '<w:tab\b')) {
            $a = ([regex]::Matches($new, $tag)).Count
            $b = ([regex]::Matches($xml, $tag)).Count
            if ($a -ne $b) { throw "Protected count changed in $part of ${target}: $tag $b -> $a. Nothing written." }
        }
        if ([regex]::IsMatch($new, '/>\s*/>')) { throw "Stray '/>' markup in $part of ${target}. Nothing written." }
        # The ONLY characters that may differ are the NBSPs acted on: normalise
        # every NBSP and every run of whitespace in both and they must agree.
        $tOld = -join ([regex]::Matches($xml, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
        $tNew = -join ([regex]::Matches($new, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
        # Compare the NON-WHITESPACE characters. Collapsing runs of whitespace
        # instead was wrong: the text of every w:t is joined with no paragraph
        # separator, so deleting a LEADING non-breaking space legitimately turns
        # "...text None" into "...textNone" and the check fired on correct work.
        # Strip all whitespace from both and nothing else may differ; the NBSP
        # tally below then pins down how many actually went.
        $normOld = $tOld -replace '[\s\u00A0]+', ''
        $normNew = $tNew -replace '[\s\u00A0]+', ''
        if ($normOld -cne $normNew) { throw "More than whitespace changed in $part of ${target}. Nothing written." }
        $nbOld = ([regex]::Matches($tOld, [regex]::Escape($NB))).Count
        $nbNew = ([regex]::Matches($tNew, [regex]::Escape($NB))).Count
        if (($nbOld - $nbNew) -ne ($pConv + $pDel)) {
            throw "Non-breaking spaces fell by $($nbOld - $nbNew) in $part of ${target}, expected $($pConv + $pDel). Nothing written."
        }

        $edits[$part] = $utf8NoBom.GetBytes($new)
        $fConv += $pConv; $fDel += $pDel
    }

    if ($edits.Count -eq 0) { Write-Host "  Nothing to repair ($fKept kept)." -ForegroundColor DarkGray; continue }
    foreach ($s in $samples) { Write-Host "    $s" -ForegroundColor DarkGray }
    Write-Host "  $fConv converted, $fDel removed, $fKept kept." -ForegroundColor Green
    $grandConv += $fConv; $grandDel += $fDel; $grandKept += $fKept
    if ($ListOnly) { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow; continue }

    if (-not $NoBackup) {
        $bak = "$target.nbsp.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    foreach ($k in $edits.Keys) { $entries[$k] = $edits[$k] }
    $out = [System.IO.Compression.ZipFile]::Open($target, 'Update')
    foreach ($e in @($out.Entries)) { $e.Delete() }
    foreach ($name in $order) {
        $entry = $out.CreateEntry($name)
        $s = $entry.Open(); $s.Write($entries[$name], 0, $entries[$name].Length); $s.Close()
    }
    $out.Dispose()
    Write-Host "  Saved." -ForegroundColor Green
}
Write-Host ""
Write-Host "Done. $grandConv converted, $grandDel removed, $grandKept kept." -ForegroundColor Cyan
