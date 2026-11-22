<#
.SYNOPSIS
    Removes direct PARAGRAPH formatting - spacing, indents, justification -
    from paragraphs that already carry a style, so the style owns the layout.

.DESCRIPTION
    Reapply-Styles.ps1 clears direct RUN formatting (w:rPr: font, size,
    colour). It does not touch w:pPr, so a paragraph on Body Text can still
    carry its own spacing, first-line indent and justification, and Word's
    style inspector shows exactly what the authors kept reporting:

        Body Text + Justified, First line: 0.25"
        Body Text + Line spacing: Double
        Table Text + Line spacing: 1.5 lines
        Table Text + Right
        References + Hanging: 0.33"

    In NWIRP 2-1 that was 1,716 overrides: 467 justifications on Table Text,
    645 spacing/indent/justification settings on Body Text, 112 indents on
    References. None of it is visible as a style, so nobody can fix it by
    editing one.

    This is a DELIBERATE, per-document-set step. It is not in
    Build-NISTDocuments.ps1 and must not be added to it: a blanket pipeline
    that strips formatting will eventually run over a file where some of that
    formatting was the point.

    NEVER TOUCHED:

      w:pPrChange   A tracked FORMAT revision. Stripping the property while
                    leaving the revision record makes Word offer to accept a
                    change into a paragraph that no longer has it. Those
                    paragraphs are skipped and reported.
      w:sectPr      Section properties live in the pPr of the paragraph that
                    ends the section. Removing one merges two sections.
      w:numPr       List membership.
      w:rPr         The paragraph mark's own run properties.
      w:tabs        A decimal tab stop is how a numeric column aligns
                    (Set-TableDecimalTabs.ps1), and it is carried here.
      w:shd, w:pBdr The shading and borders in these documents are deliberate.
      w:cnfStyle    Table conditional formatting - which banding row this is.
      w:keepNext / w:keepLines / w:pageBreakBefore
                    Pagination the author chose, which this script has no way
                    to second-guess.
      text boxes    Paragraphs inside w:txbxContent are chart furniture: axis
                    labels, callouts, panel letters. Their alignment is the
                    only thing positioning them inside the shape.

    An UNSTYLED paragraph is skipped by default: it has no style to defer to,
    so there is nothing to say its own formatting is wrong. Pass
    -IncludeUnstyled when the set has been through Set-StyleRemap.ps1 and every
    paragraph that should carry a style already does. What is left unstyled
    then is the empty paragraphs - the certified chapter set leaves 128 of its
    149 that way deliberately - and in NWIRP 2-1 seventy-one of them carried
    direct double line spacing, so every blank line in the document was a
    different height.

.PARAMETER Property
    Which w:pPr children to remove. Default: spacing, ind, jc,
    contextualSpacing. "Noise" adds the flags a converter leaves behind that
    say nothing about layout: widowControl, autoSpaceDE, autoSpaceDN,
    adjustRightInd, snapToGrid, textAlignment.

.PARAMETER Style
    Only paragraphs on these styles. Default: every paragraph that has a
    pStyle at all, since having one is what makes the style the authority.
    A paragraph with NO style is left alone - it has no style to defer to.

.PARAMETER IncludeUnstyled
    Also strip from paragraphs with no pStyle. See the note above.

.PARAMETER ParagraphMark
    Also clear the PARAGRAPH MARK's run properties (the w:rPr inside w:pPr),
    which is the font, size and colour of the pilcrow itself.

    Reapply-Styles.ps1 leaves this alone by design - its pattern only matches
    an rPr that follows a <w:r> open tag - so it survives the whole pipeline.
    That is what Word's style inspector is reporting as "Normal + 11 pt" on an
    empty paragraph: on a paragraph with no text, the mark's own size IS the
    height of the blank line, so forty-seven of them at 11pt in one chapter
    meant forty-seven blank lines that did not match the rest.
    NWIRP 2-1 carried 1,316 direct fonts, 527 sizes and 508 colours there.

    PRESERVED, matching Reapply-Styles.ps1 exactly so the two agree:
    w:ins / w:del / w:moveFrom / w:moveTo (a tracked paragraph insertion or
    deletion is recorded ON THE MARK - there are 61 of them here, and dropping
    one silently rewrites the revision history), the Footnote/EndnoteReference
    character style, w:i and w:iCs, w:highlight, and w:vertAlign.
    w:rPrChange goes, because it records the formatting being erased.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Remove-DirectParagraphFormat.ps1 -Folder "C:\...\HM-261109-v2" -ListOnly

.EXAMPLE
    .\Remove-DirectParagraphFormat.ps1 -File "Ch06.docx" -Style TableText -Property jc,spacing -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [ValidateSet('spacing','ind','jc','contextualSpacing','Noise')]
    [string[]]$Property = @('spacing','ind','jc','contextualSpacing'),

    [string[]]$Style,

    [switch]$IncludeUnstyled,

    [switch]$ParagraphMark,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$props = New-Object System.Collections.ArrayList
foreach ($p in $Property) {
    if ($p -eq 'Noise') {
        foreach ($n in @('widowControl','autoSpaceDE','autoSpaceDN','adjustRightInd','snapToGrid','textAlignment')) {
            if (-not $props.Contains($n)) { [void]$props.Add($n) }
        }
    } elseif (-not $props.Contains($p)) { [void]$props.Add($p) }
}

# Belt and braces: even if one of these is passed, it does not come out.
$NEVER = @('pStyle','numPr','rPr','sectPr','tabs','shd','pBdr','cnfStyle',
           'keepNext','keepLines','pageBreakBefore','framePr','pPrChange','outlineLvl')
$props = @($props | Where-Object { $NEVER -notcontains $_ })
if ($props.Count -eq 0) { throw "Nothing left to remove after the protected list was applied." }

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

Write-Host "Remove Direct Paragraph Format" -ForegroundColor Cyan
Write-Host ("  Removing: " + (($props | ForEach-Object { "w:$_" }) -join ', ')) -ForegroundColor Cyan
Write-Host ("  Scope: " + $(if ($Style) { "styles " + ($Style -join ', ') } else { 'every paragraph that carries a pStyle' })) -ForegroundColor Cyan
if ($ParagraphMark) { Write-Host "  Also clearing the paragraph mark's run properties (tracked-change markers kept)." -ForegroundColor Cyan }
Write-Host ("  Skipped: tracked format revisions, text box content" + $(if ($IncludeUnstyled) { '.' } else { ', unstyled paragraphs.' })) -ForegroundColor Cyan
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
    param([string]$Block)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in [regex]::Matches($Block, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { [void]$sb.Append($t.Groups[1].Value) }
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
    if ($hadBom) { $raw = $raw.Substring(1) }

    # text box spans, by offset
    $txbx = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($raw, '(?s)<w:txbxContent>.*?</w:txbxContent>')) {
        [void]$txbx.Add(@($m.Index, ($m.Index + $m.Length)))
    }

    $tally = [ordered]@{}
    $skipRev = 0; $skipUnstyled = 0; $touched = 0
    $sb = New-Object System.Text.StringBuilder
    $pos = 0

    foreach ($pm in [regex]::Matches($raw, '(?s)<w:p\b[^>]*?>.*?</w:p>')) {
        $p = $pm.Value
        $inTxbx = $false
        foreach ($s in $txbx) { if ($pm.Index -ge $s[0] -and $pm.Index -lt $s[1]) { $inTxbx = $true; break } }
        if ($inTxbx) { continue }

        $pprM = [regex]::Match($p, '(?s)<w:pPr>.*?</w:pPr>')
        if (-not $pprM.Success) { continue }
        $ppr = $pprM.Value

        $sm = [regex]::Match($ppr, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        if (-not $sm.Success -and -not $IncludeUnstyled) { $skipUnstyled++; continue }
        $sid = if ($sm.Success) { $sm.Groups[1].Value } else { '(unstyled)' }
        if ($Style -and $Style -notcontains $sid) { continue }

        # A tracked format revision records what the formatting USED to be.
        # Removing the current formatting underneath it leaves Word offering to
        # accept a change to a property that is gone.
        if ($ppr -match '<w:pPrChange\b') { $skipRev++; continue }

        $newPpr = $ppr
        foreach ($prop in $props) {
            # Match the WHOLE element, both self-closing forms. Word writes
            # either <w:jc w:val="both"/> or <w:jc w:val="both" />, and the two
            # appear in different files of the same set: an exact-string match
            # on the first silently misses every spaced one.
            $before = $newPpr
            $newPpr = [regex]::Replace($newPpr, '<w:' + $prop + '\b[^>]*?/>', '')
            $newPpr = [regex]::Replace($newPpr, '(?s)<w:' + $prop + '\b[^>]*?>.*?</w:' + $prop + '>', '')
            if ($newPpr -ne $before) {
                $k = "$sid + w:$prop"
                $tally[$k] = 1 + [int]$tally[$k]
            }
        }
        # --- the paragraph mark's own run properties ---
        if ($ParagraphMark) {
            $mrk = [regex]::Match($newPpr, '(?s)<w:rPr\b(?:\s[^>]*)?(?:/>|>(.*?)</w:rPr>)')
            if ($mrk.Success) {
                $inner = if ($mrk.Groups[1].Success) { $mrk.Groups[1].Value } else { '' }
                # Drop the tracked FORMAT record first. Word nests a second
                # w:rPr inside it, so a non-greedy match for </w:rPr> above
                # would have stopped at the inner one; take the whole record
                # out before reading anything else.
                $inner = [regex]::Replace($inner, '(?s)<w:rPrChange\b.*?</w:rPrChange>', '')
                $inner = [regex]::Replace($inner, '<w:rPrChange\b[^>]*/>', '')

                $kept = New-Object System.Text.StringBuilder
                # CT_ParaRPr puts the revision markers first, then rStyle, then
                # the character properties in CT_RPr order.
                foreach ($rev in @('ins','del','moveFrom','moveTo')) {
                    $rm = [regex]::Match($inner, '<w:' + $rev + '\b[^>]*/>')
                    if ($rm.Success) { [void]$kept.Append($rm.Value) }
                }
                $rs = [regex]::Match($inner, '(?i)<w:rStyle\b[^>]*\bw:val="(?:Footnote|Endnote)Reference"[^>]*/>')
                if ($rs.Success) { [void]$kept.Append($rs.Value) }
                # w:shd is kept as well: the shading in these documents is
                # deliberate, and nine paragraph marks in one chapter carry it.
                # CT_RPr order is i, iCs, highlight, u, shd, vertAlign.
                foreach ($keepTag in @('i','iCs','highlight','shd','vertAlign')) {
                    $km = [regex]::Match($inner, '<w:' + $keepTag + '\b(?:\s[^/>]*)?/>')
                    if ($km.Success) { [void]$kept.Append($km.Value) }
                }
                $rebuilt = if ($kept.Length -gt 0) { '<w:rPr>' + $kept.ToString() + '</w:rPr>' } else { '' }
                if ($rebuilt -ne $mrk.Value) {
                    $newPpr = $newPpr.Remove($mrk.Index, $mrk.Length).Insert($mrk.Index, $rebuilt)
                    $k2 = "$sid + paragraph mark"
                    $tally[$k2] = 1 + [int]$tally[$k2]
                }
            }
        }

        if ($newPpr -eq $ppr) { continue }

        # An empty pPr is legal but pointless; drop it when only the pStyle is
        # left it still has to stay, so only an entirely empty one goes.
        if ($newPpr -match '^<w:pPr>\s*</w:pPr>$') { $newPpr = '' }

        $newP = $p.Remove($pprM.Index, $pprM.Length).Insert($pprM.Index, $newPpr)
        [void]$sb.Append($raw.Substring($pos, $pm.Index - $pos))
        [void]$sb.Append($newP)
        $pos = $pm.Index + $pm.Length
        $touched++
    }
    [void]$sb.Append($raw.Substring($pos))
    $newDoc = $sb.ToString()

    if ($skipRev -gt 0) { Write-Host ("  Skipped {0} paragraph(s) carrying a tracked format revision." -f $skipRev) -ForegroundColor DarkYellow }
    if ($touched -eq 0) { Write-Host "  Nothing to remove." -ForegroundColor Green; continue }

    $total = 0
    foreach ($k in $tally.Keys) { $total += $tally[$k] }
    foreach ($k in ($tally.Keys | Sort-Object { -$tally[$_] })) {
        Write-Host ("  {0,-34} {1}" -f $k, $tally[$k]) -ForegroundColor DarkGray
    }
    Write-Host ("  {0} override(s) off {1} paragraph(s)." -f $total, $touched) -ForegroundColor DarkGray

    # --- guards ----------------------------------------------------------
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $newDoc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Removal produced malformed document.xml for ${name}: $($_.Exception.Message)" }

    if ((Get-Text -Block $newDoc) -cne (Get-Text -Block $raw)) { throw "Removal changed visible text in ${name}; nothing written." }

    # The revision markers inside a paragraph mark's rPr are w:ins / w:del
    # elements, counted by the same patterns below as the body ones, so a
    # dropped marker shows up there.
    foreach ($tag in @('<w:p\b', '<w:tbl>', '<w:tc>', '<w:sectPr\b', '<w:numPr>', '<w:tabs>',
                       '<w:shd\b', '<w:pBdr>', '<w:cnfStyle\b', '<w:pPrChange\b', '<w:keepNext\b',
                       '<w:ins\b', '<w:del\b', '<a:blip\b', '<w:bookmarkStart\b',
                       '<w:fldChar\b', '<w:fldSimple\b', '<w:pStyle\b')) {
        $n = ([regex]::Matches($newDoc, $tag)).Count
        $o = ([regex]::Matches($raw, $tag)).Count
        if ($n -ne $o) { throw "Protected element count changed in ${name}: $tag $o -> $n. Nothing written." }
    }

    if ($ListOnly) { continue }
    if (-not $NoBackup) {
        $bak = "$target.ppr.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $newDoc } else { $newDoc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ppr-$([guid]::NewGuid()).zip")
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
Write-Host ("Done. {0} file(s) changed." -f $filesChanged) -ForegroundColor Green
