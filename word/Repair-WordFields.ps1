<#
.SYNOPSIS
    Repairs broken Word fields: dead REF references, empty SEQ husks, and
    fields carrying no instruction at all.

.DESCRIPTION
    Three faults, all invisible in the document until Word recalculates, and
    all found in the NWIRP 2-1 author chapters.

    DEAD REF
      A REF field whose bookmark is no longer in the document. It renders
      "Error! Reference source not found." the moment anyone presses F9.
      The field is UNWRAPPED, not deleted: its cached result is kept as plain
      text and only the field machinery goes. Both dead REFs in NWIRP 2-1
      cached a real cross-reference - "Figure 23" and "Figure 26" - inside
      sentences that read "reflected in Figure 23 that shows which segments".
      Deleting the field would have left "reflected in  that shows". A field
      whose cache is empty, or is an error string, is removed outright because
      there is nothing to keep.

    EMPTY SEQ HUSK
      A SEQ field with no result, sitting beside a number the author typed by
      hand. It contributes nothing now and injects a second number on the next
      F9: "Figure 26.  1 For three road types". Twelve of these were spread
      across two chapters, every one in a caption that already read correctly.

    NO INSTRUCTION
      begin / separate / end with no w:instrText between them. Word has no
      instruction to evaluate, so it shows the object-replacement glyph. One
      caption read "Figure 3<FFFC>. FCC cell site outage reports".

    EMPTY SDT
      A structured document tag (content control) whose w:sdtContent is empty,
      so it wraps nothing. Google Docs leaves one behind at every point it
      tracked a suggestion, tagged goog_rdk_0, goog_rdk_141 and so on. They
      display nothing and do nothing. An SDT with content is NEVER touched -
      Mendeley stores each citation in one (tag MENDELEY_CITATION_v3_*) and the
      bibliography in another, and those are live.

    GOOGLE RESIDUE
      Two leftovers from a Google Docs export. A content control marked
      w:showingPlcHdr whose only content is spaces - Google writes one wherever
      it had a suggestion - and the navigation bookmarks it names
      "_heading=h.hxaqb7mcjl71", "gjdgxs" and "_GoBack". Nothing REFs any of
      them; they are position markers, like Word's own inert _Hlk bookmarks.
      Both start AND matching end go, by w:id, so the pairing stays intact. A
      bookmark that any REF, PAGEREF or NOTEREF field names is kept whatever it
      is called.

    STRAY MARKUP
      A loose "/>" sitting in the markup as character data, as in
      <w:pStyle w:val="CaptionFigure"/>/>. It gets there whenever a script
      matches an element with a pattern that stops at the last attribute quote
      and substitutes a whole element for it, leaving the original's closing
      "/>" behind. Those two characters are legal character data, so the
      document is WELL-FORMED and every XmlDocument check passes it - only a
      search for "/>" followed by "/>" finds it. Word's reaction is not
      well-defined, and it is pure corruption whatever it does.

      Matches inside w:t, w:instrText and w:delText are left alone: there the
      characters really are text.

    NEVER TOUCHED: anything carrying w:delInstrText or w:delText. Those are
    tracked DELETIONS of real fields - an author deleting a citation - and
    they are the authors' edits, not rot. Two Zotero citations and one SEQ
    field in NWIRP 2-1 are in exactly that state; an earlier scan that looked
    only for w:instrText reported all three as "empty fields" and would have
    destroyed the revision history.

.PARAMETER Only
    Run just these repairs: DeadRef, SeqHusk, NoInstruction. Default is all.

.PARAMETER ListOnly
    Report and write nothing.

.EXAMPLE
    .\Repair-WordFields.ps1 -Folder "C:\...\HM-261109-v2" -ListOnly

.EXAMPLE
    .\Repair-WordFields.ps1 -Folder "C:\...\HM-261109-v2" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [ValidateSet('DeadRef','SeqHusk','NoInstruction','EmptySdt','GoogleResidue','StrayMarkup')]
    [string[]]$Only,

    [switch]$ListOnly,
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$rules = if ($Only) { $Only } else { @('DeadRef','SeqHusk','NoInstruction','EmptySdt','GoogleResidue','StrayMarkup') }

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


Write-Host "Repair Word Fields" -ForegroundColor Cyan
Write-Host ("  Repairs: " + ($rules -join ', ')) -ForegroundColor Cyan
Write-Host "  Tracked deletions (w:delInstrText / w:delText) are never touched." -ForegroundColor Cyan
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

# Parse the complex fields: a stack over the runs that carry fldChar. Returns
# one object per field with the XML spans of its parts, innermost first so a
# nested field is repaired before its parent's spans shift.
function Get-ComplexFields {
    param([string]$Xml)
    $fields = New-Object System.Collections.ArrayList
    $stack  = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Xml, '(?s)<w:r\b(?:[^>]*)?>.*?</w:r>|<w:r\b[^>]*/>')) {
        $r = $m.Value
        if ($r -notmatch '<w:fldChar\b|<w:instrText\b|<w:delInstrText\b') { continue }
        if ($r -match '<w:fldChar\b[^>]*w:fldCharType="begin"') {
            [void]$stack.Add([pscustomobject]@{
                BeginStart = $m.Index; BeginEnd = $m.Index + $m.Length
                SepStart = -1; SepEnd = -1; EndStart = -1; EndEnd = -1
                Instr = ''; HasDel = $false
            })
            continue
        }
        if ($stack.Count -eq 0) { continue }
        $cur = $stack[$stack.Count - 1]
        if ($r -match '<w:delInstrText\b|<w:delText\b') { $cur.HasDel = $true }
        $im = [regex]::Match($r, '(?s)<w:instrText(?:\s[^>]*)?>(.*?)</w:instrText>')
        if ($im.Success) { $cur.Instr = $cur.Instr + $im.Groups[1].Value }
        if ($r -match '<w:fldChar\b[^>]*w:fldCharType="separate"') { $cur.SepStart = $m.Index; $cur.SepEnd = $m.Index + $m.Length }
        if ($r -match '<w:fldChar\b[^>]*w:fldCharType="end"') {
            $cur.EndStart = $m.Index; $cur.EndEnd = $m.Index + $m.Length
            [void]$stack.RemoveAt($stack.Count - 1)
            [void]$fields.Add($cur)
        }
    }
    # innermost / latest first
    return ,@($fields | Sort-Object -Property BeginStart -Descending)
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

    $bookmarks = @{}
    foreach ($m in [regex]::Matches($raw, '<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"')) { $bookmarks[$m.Groups[1].Value] = $true }

    $doc = $raw
    $tally = [ordered]@{}
    $notes = New-Object System.Collections.ArrayList
    $seqCacheDelta = 0
    $wsOnlyRemoved = 0

    # --- empty SEQ husks (the simple field form) -------------------------
    if ($rules -contains 'SeqHusk') {
        $husks = @([regex]::Matches($doc, '(?s)<w:fldSimple\b[^>]*w:instr="[^"]*\bSEQ\b[^"]*"[^>]*>.*?</w:fldSimple>|<w:fldSimple\b[^>]*w:instr="[^"]*\bSEQ\b[^"]*"\s*/>') |
                   Where-Object {
                       $inner = $_.Value
                       if ($inner -match '<w:delText\b|<w:delInstrText\b') { return $false }
                       -not [regex]::IsMatch($inner, '<w:t(?:\s[^>]*)?>[^<]')
                   })
        for ($k = $husks.Count - 1; $k -ge 0; $k--) {
            $doc = $doc.Remove($husks[$k].Index, $husks[$k].Length)
        }
        if ($husks.Count -gt 0) { $tally['empty SEQ husks removed'] = $husks.Count }

        # A SEQ field with \* ARABIC can only ever render digits. A cached
        # result holding anything else is stale or damaged, and Word will show
        # it until the next F9 - so it is worth putting right rather than
        # waiting for a field update that may not happen before delivery.
        # Reset it to the digits it contains.
        $fixedCache = 0
        $cands = @([regex]::Matches($doc, '(?s)<w:fldSimple\b[^>]*w:instr="[^"]*\bSEQ\b[^"]*\\\*\s*ARABIC[^"]*"[^>]*>.*?</w:fldSimple>') |
                   Where-Object { $_.Value -notmatch '<w:delText\b|<w:delInstrText\b' } |
                   Sort-Object -Property Index -Descending)
        foreach ($cm in $cands) {
            $inner = $cm.Value
            $cached = Get-Text -Block $inner
            if ($cached -match '^\d+$') { continue }
            $digits = ($cached -replace '[^\d]', '')
            if (-not $digits) { continue }
            # A plain loop, back to front. NOT a MatchEvaluator: the first w:t
            # has to be told apart from the rest, and a flag carried in a
            # $script: variable inside an evaluator never comes back - the
            # evaluator runs in its own module scope, so the work happens and
            # the flag reads unchanged.
            $tms = @([regex]::Matches($inner, '(?s)(<w:t)((?:\s[^>]*)?)(>)(.*?)(</w:t>)'))
            $rebuilt = $inner
            for ($q = $tms.Count - 1; $q -ge 0; $q--) {
                $tm = $tms[$q]
                $fill = if ($q -eq 0) { $digits } else { '' }
                $repl = $tm.Groups[1].Value + $tm.Groups[2].Value + $tm.Groups[3].Value + $fill + $tm.Groups[5].Value
                $rebuilt = $rebuilt.Remove($tm.Index, $tm.Length).Insert($tm.Index, $repl)
            }
            if ($rebuilt -ne $inner) {
                $doc = $doc.Remove($cm.Index, $cm.Length).Insert($cm.Index, $rebuilt)
                $fixedCache++
                $seqCacheDelta += ($cached.Length - $digits.Length)
                [void]$notes.Add("SEQ cached result '$cached' -> '$digits'")
            }
        }
        if ($fixedCache -gt 0) { $tally['SEQ cached results corrected'] = $fixedCache }
    }

    # --- stray markup ---------------------------------------------------
    if ($rules -contains 'StrayMarkup') {
        # Spans where "/>" is genuine text and must not be touched.
        $textSpans = @([regex]::Matches($doc, '(?s)<w:t(?:\s[^>]*)?>.*?</w:t>|<w:instrText(?:\s[^>]*)?>.*?</w:instrText>|<w:delText(?:\s[^>]*)?>.*?</w:delText>|<w:delInstrText(?:\s[^>]*)?>.*?</w:delInstrText>'))
        $hits = @([regex]::Matches($doc, '/>(\s*/>)') |
                  Where-Object {
                      $at = $_.Groups[1].Index
                      $inText = $false
                      foreach ($sp in $textSpans) { if ($at -ge $sp.Index -and $at -lt ($sp.Index + $sp.Length)) { $inText = $true; break } }
                      -not $inText
                  } | Sort-Object -Property Index -Descending)
        foreach ($h in $hits) {
            $doc = $doc.Remove($h.Groups[1].Index, $h.Groups[1].Length)
        }
        if ($hits.Count -gt 0) { $tally['stray "/>" removed'] = $hits.Count }
    }

    # --- empty structured document tags ---------------------------------
    if ($rules -contains 'EmptySdt') {
        $sdts = @([regex]::Matches($doc, '(?s)<w:sdt>(?:(?!<w:sdt>).)*?</w:sdt>') |
                  Where-Object {
                      # empty content, either form, and nothing tracked inside
                      ($_.Value -match '<w:sdtContent\s*/>' -or
                       $_.Value -match '<w:sdtContent>\s*</w:sdtContent>') -and
                      $_.Value -notmatch '<w:delText\b|<w:delInstrText\b|<w:ins\b|<w:del\b'
                  })
        for ($k = $sdts.Count - 1; $k -ge 0; $k--) {
            $doc = $doc.Remove($sdts[$k].Index, $sdts[$k].Length)
        }
        if ($sdts.Count -gt 0) {
            $tags = @($sdts | ForEach-Object { [regex]::Match($_.Value, '<w:tag\b[^>]*w:val="([^"]*)"').Groups[1].Value } |
                     Where-Object { $_ } | Select-Object -Unique -First 4)
            $tally['empty content controls removed'] = $sdts.Count
            if ($tags) { [void]$notes.Add("empty content controls tagged: $($tags -join ', ')") }
        }
    }

    # --- Google Docs residue --------------------------------------------
    if ($rules -contains 'GoogleResidue') {
        # placeholder-showing controls whose content is only whitespace
        $plc = @([regex]::Matches($doc, '(?s)<w:sdt>(?:(?!<w:sdt>).)*?</w:sdt>') |
                 Where-Object {
                     $_.Value -match '<w:showingPlcHdr\s*/>' -and
                     $_.Value -notmatch '<w:delText\b|<w:delInstrText\b|<w:ins\b|<w:del\b' -and
                     ((Get-Text -Block $_.Value).Trim() -eq '')
                 })
        for ($k = $plc.Count - 1; $k -ge 0; $k--) {
            $wsOnlyRemoved += (Get-Text -Block $plc[$k].Value).Length
            $doc = $doc.Remove($plc[$k].Index, $plc[$k].Length)
        }
        if ($plc.Count -gt 0) { $tally['placeholder content controls removed'] = $plc.Count }

        # navigation bookmarks nothing points at
        $refd = @{}
        foreach ($m in [regex]::Matches($doc, '(?s)<w:instrText(?:\s[^>]*)?>(.*?)</w:instrText>')) {
            foreach ($r in [regex]::Matches($m.Groups[1].Value, '(?i)\b(?:REF|PAGEREF|NOTEREF)\s+(\S+)')) {
                $refd[$r.Groups[1].Value] = $true
            }
        }
        foreach ($m in [regex]::Matches($doc, '(?i)w:instr="[^"]*?\b(?:REF|PAGEREF|NOTEREF)\s+([^\s""\\]+)')) {
            $refd[$m.Groups[1].Value] = $true
        }
        $dropIds = @{}
        $names = @()
        foreach ($m in [regex]::Matches($doc, '<w:bookmarkStart\b[^>]*\bw:id="(\d+)"[^>]*\bw:name="([^"]*)"')) {
            $bid = $m.Groups[1].Value; $bname = $m.Groups[2].Value
            if ($refd.ContainsKey($bname)) { continue }
            if ($bname -notmatch '^(?:_heading=h\.|gjdgxs$|_GoBack$|_Hlk)') { continue }
            $dropIds[$bid] = $true; $names += $bname
        }
        if ($dropIds.Count -gt 0) {
            $kill = New-Object System.Collections.ArrayList
            foreach ($m in [regex]::Matches($doc, '<w:bookmarkStart\b[^>]*\bw:id="(\d+)"[^>]*/>')) {
                if ($dropIds.ContainsKey($m.Groups[1].Value)) { [void]$kill.Add($m) }
            }
            foreach ($m in [regex]::Matches($doc, '<w:bookmarkEnd\b[^>]*\bw:id="(\d+)"\s*/>')) {
                if ($dropIds.ContainsKey($m.Groups[1].Value)) { [void]$kill.Add($m) }
            }
            foreach ($m in @($kill | Sort-Object -Property Index -Descending)) {
                $doc = $doc.Remove($m.Index, $m.Length)
            }
            $tally['navigation bookmarks removed'] = $dropIds.Count
            [void]$notes.Add("bookmarks removed: $(($names | Select-Object -First 4) -join ', ')")
        }
    }

    # --- complex fields: dead REF and no-instruction ---------------------
    $fields = Get-ComplexFields -Xml $doc
    $deadRef = 0; $deadRefGone = 0; $noInstr = 0
    foreach ($f in $fields) {
        if ($f.HasDel) { continue }
        if ($f.EndEnd -lt 0) { continue }
        $instr = $f.Instr.Trim()

        $isNoInstr = ($instr -eq '')
        $refName = $null
        if ($instr -match '^(?i)REF\s+(\S+)') { $refName = $Matches[1] }
        $isDeadRef = ($refName -and -not $bookmarks.ContainsKey($refName))

        if ($isNoInstr -and -not ($rules -contains 'NoInstruction')) { continue }
        if ($isDeadRef -and -not ($rules -contains 'DeadRef')) { continue }
        if (-not $isNoInstr -and -not $isDeadRef) { continue }

        $cache = ''
        if ($f.SepEnd -gt 0 -and $f.EndStart -gt $f.SepEnd) {
            $cache = Get-Text -Block $doc.Substring($f.SepEnd, $f.EndStart - $f.SepEnd)
        }
        # U+FFFC is the object-replacement glyph a no-instruction field shows.
        # It is not text anyone typed, so it is not worth keeping.
        $keep = $cache -replace "`u{FFFC}", ''
        $looksBroken = ($keep.Trim() -eq '') -or ($keep -match '(?i)Error!\s*(Reference|Bookmark)')

        if ($looksBroken) {
            # remove the whole field
            $doc = $doc.Remove($f.BeginStart, $f.EndEnd - $f.BeginStart)
            if ($isDeadRef) { $deadRefGone++ } else { $noInstr++ }
            [void]$notes.Add(("removed {0} field (nothing to keep)" -f $(if ($isDeadRef) { "dead REF $refName" } else { 'no-instruction' })))
        } else {
            # unwrap: drop begin..separate and the end run, keep the result
            $doc = $doc.Remove($f.EndStart, $f.EndEnd - $f.EndStart)
            $cutTo = if ($f.SepEnd -gt 0) { $f.SepEnd } else { $f.BeginEnd }
            $doc = $doc.Remove($f.BeginStart, $cutTo - $f.BeginStart)
            if ($isDeadRef) { $deadRef++ } else { $noInstr++ }
            [void]$notes.Add(("unwrapped {0} to plain text '{1}'" -f $(if ($isDeadRef) { "dead REF $refName" } else { 'no-instruction field' }), $keep.Trim()))
        }
    }
    if ($deadRef -gt 0)     { $tally['dead REF fields unwrapped'] = $deadRef }
    if ($deadRefGone -gt 0) { $tally['dead REF fields removed'] = $deadRefGone }
    if ($noInstr -gt 0)     { $tally['no-instruction fields repaired'] = $noInstr }

    foreach ($n in $notes) { Write-Host "    $n" -ForegroundColor DarkGray }
    if ($tally.Count -eq 0) { Write-Host "  Nothing to repair." -ForegroundColor Green; continue }
    foreach ($k in $tally.Keys) { Write-Host ("  {0,-38} {1}" -f $k, $tally[$k]) -ForegroundColor DarkGray }

    # --- guards ----------------------------------------------------------
    try { $probe = New-Object System.Xml.XmlDocument; $probe.XmlResolver = $null; $probeXml = $doc; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml) }
    catch { throw "Repair produced malformed document.xml for ${name}: $($_.Exception.Message)" }

    $b1 = ([regex]::Matches($doc, '<w:fldChar\b[^>]*w:fldCharType="begin"')).Count
    $e1 = ([regex]::Matches($doc, '<w:fldChar\b[^>]*w:fldCharType="end"')).Count
    if ($b1 -ne $e1) { throw "Field characters no longer balance in ${name} ($b1 begin / $e1 end). Nothing written." }

    # Not a foreach over @(@('a','b')): PowerShell flattens that to the two
    # STRINGS, so $pair[0] is the character '<' and the guard counted every
    # angle bracket in the file (22,869 of them) instead of the bookmarks.
    # Bookmarks may only drop by the number the GoogleResidue rule removed, and
    # the starts and ends must drop by the SAME number or the pairing is broken.
    $expectDrop = [int]$tally['navigation bookmarks removed']
    $bsNew = ([regex]::Matches($doc, '<w:bookmarkStart\b')).Count
    $bsOld = ([regex]::Matches($raw, '<w:bookmarkStart\b')).Count
    $beNew = ([regex]::Matches($doc, '<w:bookmarkEnd\b')).Count
    $beOld = ([regex]::Matches($raw, '<w:bookmarkEnd\b')).Count
    if (($bsOld - $bsNew) -ne $expectDrop -or ($beOld - $beNew) -ne $expectDrop) {
        throw "Bookmark counts changed by more than the $expectDrop removed in ${name} ($bsOld/$beOld -> $bsNew/$beNew). Nothing written."
    }
    $unpaired = @()
    $endIds = @{}
    foreach ($m in [regex]::Matches($doc, '<w:bookmarkEnd\b[^>]*\bw:id="(\d+)"')) { $endIds[$m.Groups[1].Value] = $true }
    foreach ($m in [regex]::Matches($doc, '<w:bookmarkStart\b[^>]*\bw:id="(\d+)"')) {
        if (-not $endIds.ContainsKey($m.Groups[1].Value)) { $unpaired += $m.Groups[1].Value }
    }
    if ($unpaired.Count -gt 0) { throw "Bookmark start with no end in ${name}: id $(($unpaired | Select-Object -Unique) -join ', '). Nothing written." }
    # tracked changes must survive untouched
    foreach ($tag in @('<w:ins\b', '<w:del\b', '<w:delText\b', '<w:delInstrText\b', '<w:commentReference\b')) {
        $n = ([regex]::Matches($doc, $tag)).Count; $o = ([regex]::Matches($raw, $tag)).Count
        if ($n -ne $o) { throw "Tracked-change markup changed in ${name}: $tag $o -> $n. Nothing written." }
    }
    # Visible text may only lose the object-replacement glyph, and - when a
    # placeholder control went - the whitespace that control held. Compare with
    # whitespace stripped in that case, so no actual word can go missing.
    $before = (Get-Text -Block $raw)  -replace "`u{FFFC}", ''
    $after  = (Get-Text -Block $doc) -replace "`u{FFFC}", ''
    if ($wsOnlyRemoved -gt 0 -or $seqCacheDelta -gt 0) {
        # Two repairs are allowed to change the visible text, and only in the
        # ways accounted for here: a placeholder control took whitespace with
        # it, and a SEQ cached result lost the stray punctuation that was not a
        # digit. So the texts must match once whitespace and periods are out,
        # AND the length must drop by exactly the characters those two removed.
        # Either test alone is too weak - the first would let a period vanish
        # anywhere, the second would let a character be swapped for another.
        if (([regex]::Replace($after, '[\s.]+', '')) -cne ([regex]::Replace($before, '[\s.]+', ''))) {
            throw "Repair changed more than whitespace and caption punctuation in ${name}; nothing written."
        }
        $expected = $wsOnlyRemoved + $seqCacheDelta
        $actual = $before.Length - $after.Length
        if ($actual -ne $expected) {
            throw "Repair removed $actual character(s) in ${name} but accounted for $expected. Nothing written."
        }
    } elseif ($after -cne $before) {
        throw "Repair changed visible text in ${name}; nothing written."
    }

    if ($ListOnly) { continue }
    if (-not $NoBackup) {
        $bak = "$target.fields.bak"
        if ($Force -or -not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $target -Destination $bak -Force }
    }
    $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($hadBom) { [string][char]0xFEFF + $doc } else { $doc }))
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "fields-$([guid]::NewGuid()).zip")
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
