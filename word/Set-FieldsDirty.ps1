<#
.SYNOPSIS
    Marks fields dirty so Word recomputes them the next time the document is
    opened. No Word installation required.

.DESCRIPTION
    A field caches its last computed result between the "separate" and "end"
    marks, and that cached text is what a reader sees until something
    recalculates it. A freshly written field therefore displays whatever
    placeholder its author put in the cache - which is how eight NWIRP 2-1
    chapters came to carry 'Chapter Title' in every running head and
    'Report Title1' in every footer, while the field INSTRUCTIONS were correct
    all along.

    <w:fldChar w:fldCharType="begin" w:dirty="true"/> tells Word to evaluate
    that field on load. It is the mechanism Build-NISTReport.ps1 and
    Merge-WordDocuments.ps1 already use (their private Mark-FieldsDirty), but
    both scope it to TOC / TABLE OF / LIST OF and both run it only over the
    assembled body. This exposes it as a tool that can be aimed at a part.

    WHY THIS AND NOT A WORD FIELD UPDATE.
    Set-NISTHeaderFooterContent.ps1 can drive Word to recalculate, and the
    readme and NIST-DOCUMENT-CHECKLIST both say to pass -SkipFieldUpdate
    instead, because a Word field update is what corrupted NCSTAR Ch04: a bare
    'REF bookmark \h' whose bookmarkEnd sat thousands of paragraphs downstream
    spanned ~1.75 MB, and one F9 injected all of it, taking that chapter from
    6,257 paragraphs to 22,958.

    w:dirty is per-field, so marking the header and footer fields recomputes
    exactly those and touches nothing in the body. That makes it safe where a
    whole-document update is not, it needs no Word on this machine, and it
    fires on whatever machine the reviewer opens the file on.

    Check Test-NISTDocument.ps1's 'no REF field pastes a runaway bookmark'
    before marking body fields dirty. For header and footer parts it does not
    apply - those carry STYLEREF, PAGE and TITLE, none of which paste a span.

    NEVER TOUCHED: a field inside a tracked deletion. Its instruction is a
    w:delInstrText, it is an author's pending edit, and Word must not be told
    to evaluate it.

.PARAMETER Part
    Regex matched against package part names. The default covers both naming
    schemes: Word's own header1.xml / footer2.xml AND the template's
    nistHeaderDefault.xml / nistFooterEven.xml, which a '^word/header\d+\.xml$'
    pattern misses entirely.

.PARAMETER Instruction
    Only mark a field whose instruction matches this regex. Unset means every
    field in the selected parts. Use it to reproduce Mark-FieldsDirty's scope:
    -Instruction 'TOC|TABLE OF|LIST OF'

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    # The running heads and footers, which is what this was written for.
    .\Set-FieldsDirty.ps1 -Folder "C:\...\HM-261109-v2" -ListOnly

.EXAMPLE
    .\Set-FieldsDirty.ps1 -Folder "C:\...\HM-261109-v2" -Force

.EXAMPLE
    # A table of contents in an assembled deliverable.
    .\Set-FieldsDirty.ps1 -File "Combined.docx" -Part '^word/document\.xml$' `
        -Instruction 'TOC|TABLE OF|LIST OF' -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$Part = '^word/[^/]*(?:header|footer)[^/]*\.xml$',

    [string]$Instruction,

    # Remove w:dirty instead of adding it. A dirty field makes Word update
    # fields on open, and Word announces that with "This document contains
    # fields that may refer to other files. Do you want to update the fields in
    # this document?" - on every open, for every file. That prompt is worse than
    # the problem it solves: an author who is asked a scary question about
    # external files every time they open a chapter will click No, and the
    # numbers never update anyway. The better answer is a correct CACHED value,
    # which needs no update at all.
    [switch]$Clear,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$DIRTY = ' w:dirty="true"'

function Remove-DirtyAttr {
    # Strip the attribute in whatever form Word or this script wrote it:
    # w:dirty="true", ="1", ="on", with or without surrounding space.
    param([string]$Element)
    return [regex]::Replace($Element, '\s+w:dirty\s*=\s*"[^"]*"', '')
}

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


Write-Host "Set Fields Dirty" -ForegroundColor Cyan
Write-Host "  Parts: /$Part/" -ForegroundColor Cyan
Write-Host ("  Fields: " + $(if ($Instruction) { "instruction matching /$Instruction/" } else { 'every field in those parts' })) -ForegroundColor Cyan
Write-Host "  Word recomputes a dirty field on open. Tracked deletions are never marked." -ForegroundColor Cyan
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

# Add w:dirty to a self-closing element's open tag. Word writes either
# '<w:fldChar .../>' or '<w:fldChar ... />' and both forms appear in one
# document set, so the closing is matched rather than assumed.
function Add-DirtyAttr {
    param([string]$Element)
    if ($Element -match '\bw:dirty\s*=') { return $Element }
    # Replace ONLY the final '/>', never any whitespace in front of it. Word
    # writes both '<w:fldChar .../>' and '<w:fldChar ... />' and both forms
    # appear in this document set, so a pattern of '\s*/>$' consumed that space
    # on the spaced form and the added length came out 14 instead of 15 - which
    # is what tripped the accounting guard on Ch02's 38 spaced elements.
    #
    # A w:fldSimple open tag is NOT self-closing - it is '<w:fldSimple w:instr="
    # ... ">' with content and a closing tag - so there is no '/>' to replace and
    # this returned the element UNTOUCHED while the caller still counted it as
    # marked. The accounting guard caught it exactly: 2,370 characters added
    # against 162 fields marked is 158 at 15 characters each, and Ch04 holds
    # precisely 4 fldSimple. Without that guard the script would have reported
    # marking every field while silently skipping every simple one, which is the
    # worst version of this bug: the header numbers stay stale and the run says
    # it fixed them.
    if ($Element -match '/>$') {
        return [regex]::Replace($Element, '/>$', ($DIRTY + '/>'), 1)
    }
    return [regex]::Replace($Element, '>$', ($DIRTY + '>'), 1)
}

$filesChanged = 0
foreach ($target in $targets) {
    $name = [System.IO.Path]::GetFileName($target)
    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    $parts = @($pkg.Names | Where-Object { $_ -match $Part })
    if ($parts.Count -eq 0) {
        Write-Host "  No part matches /$Part/." -ForegroundColor DarkYellow
        continue
    }

    $marked = 0
    $already = 0
    $skippedDel = 0
    $skippedInstr = 0
    $edits = @{}
    $detail = New-Object System.Collections.ArrayList

    foreach ($p in $parts) {
        $raw = [System.Text.Encoding]::UTF8.GetString($pkg.Data[$p])
        # No String.StartsWith(char) on .NET Framework: the char is coerced and
        # the culture-sensitive compare treats U+FEFF as ignorable, so that test
        # is always true and the rewrite adds a BOM that was never there.
        $hadBom = ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF)
        $xml = if ($hadBom) { $raw.Substring(1) } else { $raw }

        # A field's instruction is not inside its begin element, so the
        # instruction filter needs the whole begin..end span. Walk the runs with
        # a depth counter the way Repair-WordFields does, so a nested field is
        # attributed to its own begin rather than its parent's.
        $fields = New-Object System.Collections.ArrayList
        $stack = New-Object System.Collections.ArrayList
        foreach ($rm in [regex]::Matches($xml, '(?s)<w:r\b(?:\s[^>]*)?>.*?</w:r>|<w:r\b[^>]*/>')) {
            $r = $rm.Value
            if ($r -notmatch '<w:fldChar\b|<w:instrText\b|<w:delInstrText\b') { continue }
            $bm = [regex]::Match($r, '<w:fldChar\b[^>]*?\bw:fldCharType="begin"[^>]*?/>')
            if ($bm.Success) {
                [void]$stack.Add([pscustomobject]@{
                    Index = $rm.Index + $bm.Index; Length = $bm.Length
                    Element = $bm.Value; Instr = ''; HasDel = $false
                })
                continue
            }
            if ($stack.Count -eq 0) { continue }
            $cur = $stack[$stack.Count - 1]
            if ($r -match '<w:delInstrText\b|<w:delText\b') { $cur.HasDel = $true }
            $im = [regex]::Match($r, '(?s)<w:instrText(?:\s[^>]*)?>(.*?)</w:instrText>')
            if ($im.Success) { $cur.Instr = $cur.Instr + $im.Groups[1].Value }
            if ($r -match '<w:fldChar\b[^>]*?\bw:fldCharType="end"') {
                [void]$stack.RemoveAt($stack.Count - 1)
                [void]$fields.Add($cur)
            }
        }
        # An unterminated field never reached the end branch. Mark it anyway if
        # it qualifies: the instruction is already collected, and leaving it is
        # what left the placeholder on screen in the first place.
        foreach ($cur in $stack) { [void]$fields.Add($cur) }

        # A one-element simple field carries its instruction as an attribute.
        # HasDel cannot be read off the open tag the way it is for a fldChar
        # field, where the deletion shows up in a following run: a simple field
        # sits INSIDE the w:del instead. Collect the deletion spans first and
        # flag any simple field that starts inside one, or this script breaks
        # its own rule and marks a field an author has deleted. NWIRP 2-1 has
        # none, so nothing here changes - but 125 of its fields are fldSimple,
        # and a rule that holds only for the other form is not a rule.
        $delSpans = New-Object System.Collections.ArrayList
        foreach ($dm in [regex]::Matches($xml, '(?s)<w:del\b[^>]*?>.*?</w:del>')) {
            [void]$delSpans.Add([pscustomobject]@{ Start = $dm.Index; End = $dm.Index + $dm.Length })
        }
        foreach ($sm in [regex]::Matches($xml, '<w:fldSimple\b[^>]*?>')) {
            $iv = [regex]::Match($sm.Value, '\bw:instr="([^"]*)"')
            $inDel = $false
            foreach ($d in $delSpans) {
                if ($sm.Index -ge $d.Start -and $sm.Index -lt $d.End) { $inDel = $true; break }
            }
            [void]$fields.Add([pscustomobject]@{
                Index = $sm.Index; Length = $sm.Length
                Element = $sm.Value; Instr = $iv.Groups[1].Value; HasDel = $inDel
            })
        }

        $wanted = New-Object System.Collections.ArrayList
        foreach ($f in $fields) {
            if ($f.HasDel) { $skippedDel++; continue }
            if ($Clear) {
                # Only fields that CARRY the attribute are of interest now.
                if ($f.Element -notmatch '\bw:dirty\s*=') { $already++; continue }
            } else {
                if ($f.Element -match '\bw:dirty\s*=') { $already++; continue }
            }
            if ($Instruction -and $f.Instr -notmatch $Instruction) { $skippedInstr++; continue }
            [void]$wanted.Add($f)
            $instr = ($f.Instr -replace '\s+', ' ').Trim()
            if ($detail.Count -lt 8 -and $instr) { [void]$detail.Add("$(Split-Path -Leaf $p): $instr") }
        }
        if ($wanted.Count -eq 0) { continue }

        # BACK TO FRONT. Every edit lengthens the string, so forwards each later
        # Index is stale by however much has been inserted ahead of it.
        $new = $xml
        foreach ($f in @($wanted | Sort-Object -Property Index -Descending)) {
            $repl = if ($Clear) { Remove-DirtyAttr -Element $f.Element } else { Add-DirtyAttr -Element $f.Element }
            $new = $new.Remove($f.Index, $f.Length).Insert($f.Index, $repl)
        }
        $marked += $wanted.Count
        $removedChars = 0
        if ($Clear) {
            foreach ($f in $wanted) { $removedChars += ($f.Element.Length - (Remove-DirtyAttr -Element $f.Element).Length) }
        }
        $edits[$p] = [pscustomobject]@{ Old = $xml; New = $new; HadBom = $hadBom; Count = $wanted.Count; Removed = $removedChars }
    }

    foreach ($d in $detail) { Write-Host "    $d" -ForegroundColor DarkGray }
    $bits = @($(if ($Clear) { "cleared $marked" } else { "marked $marked" }))
    if ($already -gt 0)      { $bits += $(if ($Clear) { "$already not dirty" } else { "$already already dirty" }) }
    if ($skippedDel -gt 0)   { $bits += "$skippedDel inside a tracked deletion, left alone" }
    if ($skippedInstr -gt 0) { $bits += "$skippedInstr did not match the instruction filter" }
    Write-Host ("  " + ($bits -join ', ')) -ForegroundColor DarkGray

    if ($marked -eq 0) { Write-Host $(if ($Clear) { "  Nothing to clear." } else { "  Nothing to mark." }) -ForegroundColor Green; continue }

    # --- guards ----------------------------------------------------------
    # This pass adds an attribute and nothing else, so every one of these is an
    # exact equality rather than a tolerance.
    foreach ($p in $edits.Keys) {
        $e = $edits[$p]
        try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $e.New; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
        catch { throw "Marking produced malformed $p for ${name}: $($_.Exception.Message). Nothing written." }

        if ((Get-Text -Xml $e.New) -cne (Get-Text -Xml $e.Old)) {
            throw "Marking changed visible text in $p of ${name}. Nothing written."
        }
        $actual = $e.New.Length - $e.Old.Length
        if ($Clear) {
            # The attribute is not always the 15 characters this script writes:
            # Word may have written w:dirty="1". Count what actually went.
            $expected = -1 * $e.Removed
            if ($actual -ne $expected) {
                throw "Clearing removed $(-$actual) character(s) from $p of ${name} but accounted for $(-$expected). Nothing written."
            }
            if (([regex]::Matches($e.New, '\bw:dirty\s*=')).Count -ne 0) {
                throw "Clearing left w:dirty behind in $p of ${name}. Nothing written."
            }
        } else {
            $expected = $e.Count * $DIRTY.Length
            if ($actual -ne $expected) {
                throw "Marking added $actual character(s) to $p of ${name} but accounted for $expected. Nothing written."
            }
        }
        foreach ($tag in @('<w:fldChar\b', 'w:fldCharType="begin"', 'w:fldCharType="end"',
                           '<w:fldSimple\b', '<w:instrText\b', '<w:delInstrText\b',
                           '<w:p\b', '<w:t[ >]')) {
            $n = ([regex]::Matches($e.New, $tag)).Count
            $o = ([regex]::Matches($e.Old, $tag)).Count
            if ($n -ne $o) { throw "Element count changed in $p of ${name}: $tag $o -> $n. Nothing written." }
        }
        $stray = ([regex]::Matches($e.New, '/>\s*/>')).Count - ([regex]::Matches($e.Old, '/>\s*/>')).Count
        if ($stray -gt 0) { throw "Marking left $stray stray '/>' in $p of ${name}. Nothing written." }
    }

    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        $bak = "$target.dirty.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }

    foreach ($p in $edits.Keys) {
        $e = $edits[$p]
        $pkg.Data[$p] = $utf8NoBom.GetBytes($(if ($e.HadBom) { [string][char]0xFEFF + $e.New } else { $e.New }))
    }

    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "fielddirty-$([guid]::NewGuid()).zip")
    $out = [System.IO.File]::Create($tempZip)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($out, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($n in $pkg.Names) {
                $ze = $zip.CreateEntry($n, [System.IO.Compression.CompressionLevel]::Optimal)
                $s = $ze.Open(); try { $s.Write($pkg.Data[$n], 0, $pkg.Data[$n].Length) } finally { $s.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $out.Dispose() }

    try {
        # Move-Item -Force cannot replace an existing file across the Windows
        # drive mount, so copy over and delete, clearing any read-only bit.
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
Write-Host ("Done. {0} file(s) changed. Word recomputes the marked fields the next time each file is opened." -f $filesChanged) -ForegroundColor Green
