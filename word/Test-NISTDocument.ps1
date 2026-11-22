<#
.SYNOPSIS
    Checks a finished Word file against the NIST NCSTAR/NWIRP spec and reports
    pass, warn or fail per check. Reads only; never writes.

.DESCRIPTION
    Every check here exists because the thing it looks for has gone wrong
    silently at least once. Word does not complain about most of them:

      - A w:basedOn naming a style that is not in the file is ignored, not
        reported. The style falls back to Normal and then to w:docDefaults, so a
        document whose docDefaults say Arial renders six heading levels in Arial
        with nothing in any log to show it.
      - A w:numPr pointing at a numId that numbering.xml does not define shows
        no number at all.
      - A REF field whose bookmark is gone reads "Error! Reference source not
        found" only after a field update, which may be weeks later.
      - An empty SEQ fldSimple husk left behind by a caption edit still
        increments the sequence, so the next F9 renumbers every caption after it.

    Run it after Build-NISTDocuments.ps1. Exit code is the number of failures,
    so it can gate a build.

    FAIL means the document is wrong and a reader would see it. WARN means it
    departs from house style or carries review scaffolding, which is legitimate
    mid-review. Counts are reported either way.

.PARAMETER File
    A single .docx/.dotx to check.

.PARAMETER Folder
    Check every Word file in this folder.

.PARAMETER Quiet
    Print only failures and warnings, not passing checks.

.EXAMPLE
    .\Test-NISTDocument.ps1 -File "C:\Users\Artist\Documents\HM-261103\NWIRP_2-3_Chapter-06_Findings.docx"

.EXAMPLE
    .\Test-NISTDocument.ps1 -Folder "C:\Users\Artist\Documents\HM-261102" -Quiet
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-WordParts {
    # name -> text, for every xml part we care about.
    param([string]$Path)
    $map = @{}
    # FileShare::ReadWrite, not Read: a document the user has open in Word is
    # held by a writer, and FileShare::Read then throws "being used by another
    # process". Reading is safe - nothing here writes to the package.
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                # .rels too, not just .xml: header and footer parts are found
                # through word/_rels/document.xml.rels, and excluding it meant
                # every header lookup fell through to the filename fallback.
                if ($e.FullName -notmatch '\.(xml|rels)$') { continue }
                $r = New-Object System.IO.StreamReader($e.Open())
                try { $text = $r.ReadToEnd() } finally { $r.Dispose() }
                if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
                $map[$e.FullName] = $text
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
    return $map
}

function Get-StyleMap {
    # styleId -> body of the w:style element.
    # styleId is NOT assumed to be the last attribute: Word writes
    # <w:style w:type="paragraph" w:styleId="X" w:customStyle="1">.
    param([string]$Xml)
    $map = @{}
    if (-not $Xml) { return $map }
    foreach ($m in [regex]::Matches($Xml, '(?s)<w:style\b([^>]*)>(.*?)</w:style>')) {
        $id = [regex]::Match($m.Groups[1].Value, 'w:styleId="([^"]+)"')
        if ($id.Success) { $map[$id.Groups[1].Value] = $m.Groups[2].Value }
    }
    return $map
}

function Get-Blocks {
    # Body-level paragraphs and tables, in document order.
    param([string]$Xml)
    $body = $Xml
    $i = $Xml.IndexOf('<w:body>')
    $j = $Xml.LastIndexOf('</w:body>')
    if ($i -ge 0 -and $j -gt $i) { $body = $Xml.Substring($i + 8, $j - $i - 8) }
    return [regex]::Matches($body, '(?s)<w:p\b(?:[^>]*/>|.*?</w:p>)|<w:tbl>.*?</w:tbl>')
}

function Get-Text {
    param([string]$Block)
    $sb = New-Object System.Text.StringBuilder
    foreach ($m in [regex]::Matches($Block, '(?s)<w:t(?:\s[^>]*)?>([^<]*)</w:t>')) { [void]$sb.Append($m.Groups[1].Value) }
    return $sb.ToString()
}

function Get-Style {
    param([string]$Block)
    $m = [regex]::Match($Block, '<w:pStyle\b[^>]*w:val="([^"]+)"')
    if ($m.Success) { return $m.Groups[1].Value }
    if ($Block.StartsWith('<w:tbl')) { return 'TABLE' }
    return 'Normal'
}

# ---------------------------------------------------------------------------
# Target list
# ---------------------------------------------------------------------------

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets = @((Resolve-Path -LiteralPath $File).ProviderPath)
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $targets = @(Get-ChildItem -LiteralPath $Folder -File |
                 Where-Object { $_.Extension -match '^\.(docx|dotx|docm|dotm)$' -and $_.Name -notlike '~$*' } |
                 ForEach-Object { $_.FullName })
    if ($targets.Count -eq 0) { throw "No Word files in $Folder" }
}

$totalFail = 0

foreach ($target in $targets) {

    $results = New-Object System.Collections.ArrayList
    function Add-Result {
        param([string]$Group, [string]$Name, [string]$Status, [string]$Detail)
        [void]$results.Add([pscustomobject]@{ Group = $Group; Name = $Name; Status = $Status; Detail = $Detail })
    }

    try { $parts = Get-WordParts -Path $target }
    catch {
        # Report and carry on. Aborting the loop meant a single locked or
        # corrupt file silently skipped every file after it in the folder.
        Write-Host ""
        Write-Host ([System.IO.Path]::GetFileName($target)) -ForegroundColor White
        Write-Host ("    [FAIL] readable as a Word package              " + $_.Exception.Message.Split([char]10)[0]) -ForegroundColor Red
        $totalFail++
        continue
    }
    $doc     = $parts['word/document.xml']
    $stylesX = $parts['word/styles.xml']
    $numX    = $parts['word/numbering.xml']
    if (-not $doc)     { throw "No word/document.xml in $target" }
    if (-not $stylesX) { throw "No word/styles.xml in $target" }

    $styles = Get-StyleMap -Xml $stylesX
    $blocks = Get-Blocks -Xml $doc

    # --------------------------------------------------------------- STYLES
    # Used styles, from the body and from every header and footer story.
    $used = @{}
    foreach ($key in $parts.Keys) {
        if ($key -ne 'word/document.xml' -and $key -notmatch '^word/(header|footer|footnotes|endnotes)\d*\.xml$') { continue }
        foreach ($pat in @('<w:pStyle\b[^>]*w:val="([^"]+)"', '<w:rStyle\b[^>]*w:val="([^"]+)"', '<w:tblStyle\b[^>]*w:val="([^"]+)"')) {
            foreach ($m in [regex]::Matches($parts[$key], $pat)) { $used[$m.Groups[1].Value] = $true }
        }
    }
    $missing = @($used.Keys | Where-Object { -not $styles.ContainsKey($_) } | Sort-Object)
    if ($missing.Count -gt 0) {
        Add-Result 'STYLES' 'every used style is defined' 'FAIL' "$($missing.Count) missing: $($missing -join ', ')"
    } else {
        Add-Result 'STYLES' 'every used style is defined' 'PASS' "$($used.Count) used, 0 missing"
    }

    $dangling = @()
    foreach ($sid in $styles.Keys) {
        foreach ($rel in @('basedOn', 'link', 'next')) {
            $m = [regex]::Match($styles[$sid], '<w:' + $rel + '\b[^>]*w:val="([^"]+)"')
            if ($m.Success -and -not $styles.ContainsKey($m.Groups[1].Value)) {
                $dangling += "$sid ($rel -> $($m.Groups[1].Value))"
            }
        }
    }
    if ($dangling.Count -gt 0) {
        Add-Result 'STYLES' 'no dangling basedOn/link/next' 'FAIL' ("$($dangling.Count): " + (($dangling | Sort-Object | Select-Object -First 8) -join '; '))
    } else {
        Add-Result 'STYLES' 'no dangling basedOn/link/next' 'PASS' "$($styles.Count) styles checked"
    }

    # Three base styles, not four. xxgraphic never existed in the original
    # nist-scripts architecture; it was toolkit drift, carried a stray frame,
    # and was used by nothing - removed from the script and the documents.
    $base = @('xxbase', 'xxheader', 'xxtable')
    $noBase = @($base | Where-Object { -not $styles.ContainsKey($_) })
    if ($noBase.Count -gt 0) {
        Add-Result 'STYLES' 'xx* base family present' 'WARN' "absent: $($noBase -join ', '). Run Set-NISTStyleArchitecture.ps1 to install them."
    } else {
        Add-Result 'STYLES' 'xx* base family present' 'PASS' 'all three'
    }

    # docDefaults is the fallback for anything with no style of its own.
    $ddFont = ''
    $dd = [regex]::Match($stylesX, '(?s)<w:docDefaults>.*?</w:docDefaults>')
    if ($dd.Success) {
        $f = [regex]::Match($dd.Value, '<w:rFonts\b[^>]*w:ascii="([^"]+)"')
        if ($f.Success) { $ddFont = $f.Groups[1].Value }
    }
    $normalParas = @($blocks | Where-Object { (Get-Style -Block $_.Value) -eq 'Normal' -and (Get-Text -Block $_.Value).Trim() }).Count
    if ($ddFont -in @('Arial', 'Calibri', 'Cambria', 'Times New Roman') -and $normalParas -gt 0) {
        Add-Result 'STYLES' 'docDefaults font' 'FAIL' "$ddFont, and $normalParas paragraph(s) are on Normal, so they render in it"
    } elseif ($ddFont -eq 'Arial') {
        Add-Result 'STYLES' 'docDefaults font' 'WARN' "Arial, but no paragraph is on Normal so nothing renders in it"
    } else {
        Add-Result 'STYLES' 'docDefaults font' 'PASS' $ddFont
    }

    # Skip shape content. A chart's axis labels, date ticks and callouts live in
    # w:txbxContent inside the drawing, and their face and size are part of the
    # picture: Reapply-Styles.ps1 leaves them alone on purpose, so counting them
    # here reports an offence against a file that is correct. In NWIRP 2-1 Ch04
    # all fourteen remaining direct fonts were inside shapes.
    $shapeSpans = @([regex]::Matches($doc, '(?s)<w:txbxContent>.*?</w:txbxContent>'))
    $inShape = {
        param([int]$At)
        foreach ($sp in $shapeSpans) { if ($At -ge $sp.Index -and $At -lt ($sp.Index + $sp.Length)) { return $true } }
        return $false
    }
    $directFonts = @{}
    foreach ($m in [regex]::Matches($doc, '<w:rFonts\b[^>]*w:ascii="([^"]+)"')) {
        if (& $inShape $m.Index) { continue }
        $directFonts[$m.Groups[1].Value] = 1 + [int]$directFonts[$m.Groups[1].Value]
    }
    # Cambria Math is the correct face for MathType equations, not an offender.
    $offenders = @($directFonts.Keys | Where-Object { $_ -match '^(Arial|Calibri|Cambria)' -and $_ -ne 'Cambria Math' })
    if ($offenders.Count -gt 0) {
        $d = ($offenders | ForEach-Object { "$_ x$($directFonts[$_])" }) -join ', '
        Add-Result 'STYLES' 'no non-NIST fonts applied directly' 'WARN' "$d in document.xml. Styles cannot override direct formatting; Reapply-Styles.ps1 clears it."
    } else {
        Add-Result 'STYLES' 'no non-NIST fonts applied directly' 'PASS' "$($directFonts.Count) distinct direct font(s), none off-palette"
    }

    # ------------------------------------------------------------ NUMBERING
    $numIds = @{}
    if ($numX) { foreach ($m in [regex]::Matches($numX, '<w:num\b[^>]*w:numId="(\d+)"')) { $numIds[$m.Groups[1].Value] = $true } }
    $wanted = @{}
    foreach ($m in [regex]::Matches($doc, '(?s)<w:numPr>.*?</w:numPr>')) {
        $n = [regex]::Match($m.Value, '<w:numId\b[^>]*w:val="(\d+)"')
        if ($n.Success -and $n.Groups[1].Value -ne '0') { $wanted[$n.Groups[1].Value] = $true }
    }
    $orphan = @($wanted.Keys | Where-Object { -not $numIds.ContainsKey($_) } | Sort-Object)
    if ($orphan.Count -gt 0) {
        Add-Result 'NUMBERING' 'every numPr numId resolves' 'FAIL' "numId $($orphan -join ', ') used in the body but not defined in numbering.xml"
    } else {
        Add-Result 'NUMBERING' 'every numPr numId resolves' 'PASS' "$($wanted.Count) numId(s) in use, all defined"
    }

    # Heading numbering resolves through the HMList chain, or headings are
    # unnumbered on purpose. Follow numStyleLink before calling it broken.
    $h1 = $styles['Heading1']
    $chain = 'unnumbered'
    if ($h1) {
        $n = [regex]::Match($h1, '(?s)<w:numPr>.*?<w:numId\b[^>]*w:val="(\d+)"')
        if ($n.Success -and $n.Groups[1].Value -ne '0') {
            $nid = $n.Groups[1].Value
            if (-not $numIds.ContainsKey($nid)) { $chain = "broken: Heading1 -> numId $nid, which numbering.xml does not define" }
            else {
                $numEl = [regex]::Match($numX, '(?s)<w:num\b[^>]*w:numId="' + $nid + '".*?</w:num>')
                $abs = [regex]::Match($numEl.Value, '<w:abstractNumId\b[^>]*w:val="(\d+)"')
                if (-not $abs.Success) { $chain = "broken: num $nid has no abstractNumId" }
                else {
                    $absEl = [regex]::Match($numX, '(?s)<w:abstractNum\b[^>]*w:abstractNumId="' + $abs.Groups[1].Value + '".*?</w:abstractNum>')
                    if (-not $absEl.Success) { $chain = "broken: abstractNum $($abs.Groups[1].Value) is absent" }
                    else {
                        $link = [regex]::Match($absEl.Value, '<w:numStyleLink\b[^>]*w:val="([^"]+)"')
                        if ($link.Success -and -not $styles.ContainsKey($link.Groups[1].Value)) {
                            $chain = "broken: numStyleLink -> $($link.Groups[1].Value), which is not a style in this file"
                        } else { $chain = 'ok' }
                    }
                }
            }
        }
    }
    if ($chain -eq 'ok')                { Add-Result 'NUMBERING' 'heading numbering chain' 'PASS' 'Heading1 resolves to an abstractNum' }
    elseif ($chain -eq 'unnumbered')    { Add-Result 'NUMBERING' 'heading numbering chain' 'WARN' 'Heading1 carries no numbering; heading numbers must be typed' }
    else                                { Add-Result 'NUMBERING' 'heading numbering chain' 'FAIL' $chain }

    # -------------------------------------------------------------- CAPTIONS
    $capShape = '^\s*(Table|Figure)[\s\u00a0]+[0-9A-Za-z]+\s?[-\u2010-\u2015]\s?\d+'
    $wrongStyle = @(); $wrongPlace = @()
    for ($i = 0; $i -lt $blocks.Count; $i++) {
        $b = $blocks[$i].Value
        $sty = Get-Style -Block $b
        $txt = Get-Text -Block $b
        # A paragraph already on a caption style only has to start with the
        # number. One that is not has to carry the caption's period too, or
        # every sentence opening "Figure 4-5 includes ..." is read as a caption.
        $onCapStyle = $sty -in @('CaptionTable', 'CaptionFigure', 'CaptionFigure2')
        $pattern = if ($onCapStyle) { $capShape } else { $capShape + '\s*(\[was[^\]]*\])?\s*[.)]' }
        $m = [regex]::Match($txt, $pattern)
        if (-not $m.Success) { continue }
        $label = $m.Groups[1].Value
        $want = if ($label -eq 'Table') { 'CaptionTable' } else { 'CaptionFigure' }
        if ($sty -ne $want) { $wrongStyle += "$($txt.Substring(0, [Math]::Min(44, $txt.Length))) [on $sty, expected $want]" }
        # table captions sit above their table, figure captions below their figure
        if ($sty -eq 'CaptionTable') {
            $next = if ($i + 1 -lt $blocks.Count) { $blocks[$i+1].Value } else { '' }
            if (-not $next.StartsWith('<w:tbl')) { $wrongPlace += "$label caption not immediately above a table: $($txt.Substring(0, [Math]::Min(40, $txt.Length)))" }
        }
    }
    if ($wrongStyle.Count -gt 0) {
        Add-Result 'CAPTIONS' 'caption on the matching style' 'WARN' ("$($wrongStyle.Count): " + (($wrongStyle | Select-Object -First 5) -join '; '))
    } else { Add-Result 'CAPTIONS' 'caption on the matching style' 'PASS' 'all captions matched their label' }
    if ($wrongPlace.Count -gt 0) {
        Add-Result 'CAPTIONS' 'table caption sits above its table' 'WARN' ("$($wrongPlace.Count): " + (($wrongPlace | Select-Object -First 5) -join '; '))
    } else { Add-Result 'CAPTIONS' 'table caption sits above its table' 'PASS' 'ok' }

    # ---------------------------------------------------------------- TABLES
    # Content tables: every table except the page-anchored banner, which is
    # page furniture and carries its own colour bands.
    $bannerRxT = '<w:tblpPr[^>]*w:vertAnchor="page"[^>]*w:horzAnchor="page"|<w:tblpPr[^>]*w:horzAnchor="page"[^>]*w:vertAnchor="page"'
    $tblAllForData = @([regex]::Matches($doc, '(?s)<w:tbl>.*?</w:tbl>') |
                       Where-Object { -not [regex]::IsMatch($_.Value, $bannerRxT) })
    $nistTables = @($blocks | Where-Object { $_.Value.StartsWith('<w:tbl') -and $_.Value -match '<w:tblStyle\b[^>]*w:val="NIST-Data"' })
    $badLook = 0; $noHeader = 0; $barGaps = 0; $badNote = 0; $withNote = 0; $barExtra = 0; $popTables = 0
    foreach ($t in $nistTables) {
        $tx = $t.Value
        if ($tx -notmatch '<w:tblLook\b[^>]*w:val="0420"') { $badLook++ }
        $rows = [regex]::Matches($tx, '(?s)<w:tr\b.*?</w:tr>')
        if ($rows.Count -eq 0) { continue }
        if ($rows[0].Value -notmatch '<w:tblHeader\s*/>') { $noHeader++ }
        $lastIsNote = $rows[$rows.Count - 1].Value -match 'w:val="TableandFigureFootnotes"'
        if ($lastIsNote) {
            $withNote++
            $cell = [regex]::Match($rows[$rows.Count - 1].Value, '(?s)<w:tcPr>.*?</w:tcPr>').Value
            if ($cell -notmatch 'w:fill="auto"' -or
                $cell -notmatch '<w:left w:val="nil"\s*/>' -or
                $cell -notmatch '<w:bottom w:val="nil"\s*/>' -or
                $cell -notmatch '<w:right w:val="nil"\s*/>' -or
                $cell -notmatch 'w:val="000000000000"') { $badNote++ }
        }
        # Whether a table SHOULD carry the bar is an editorial call - it marks the
        # schools/hospitals population, so methodology, funding and findings
        # tables do not get one. That cannot be read off the markup, so do not
        # guess. What is checkable: a bar, once present, must run every data row.
        $dataRows = if ($lastIsNote) { $rows.Count - 1 } else { $rows.Count }
        $withBar = 0
        for ($r = 0; $r -lt $dataRows; $r++) {
            $firstCell = [regex]::Match($rows[$r].Value, '(?s)<w:tc>.*?</w:tc>').Value
            if ($firstCell -match '<w:left w:val="single"[^>]*w:sz="32"') { $withBar++ }
        }
        if ($withBar -eq 0) { continue }
        $popTables++
        if ($withBar -lt $dataRows) { $barGaps += ($dataRows - $withBar); $barExtra++ }
    }
    if ($nistTables.Count -eq 0 -and $tblAllForData.Count -eq 0) {
        # No content tables at all. The Preface, the Introduction and the
        # Findings chapter each carry only the page-anchored banner, and
        # warning that none of their tables is on NIST-Data reads as a defect
        # in a file that has nothing to put there.
        Add-Result 'TABLES' 'NIST-Data tables' 'PASS' 'no content tables'
    } elseif ($nistTables.Count -eq 0) {
        Add-Result 'TABLES' 'NIST-Data tables' 'WARN' 'no table is on NIST-Data'
    } else {
        if ($badLook -gt 0) { Add-Result 'TABLES' 'tblLook 0420' 'WARN' "$badLook of $($nistTables.Count) table(s) differ" }
        else                { Add-Result 'TABLES' 'tblLook 0420' 'PASS' "$($nistTables.Count) table(s)" }
        if ($noHeader -gt 0) { Add-Result 'TABLES' 'header row repeats' 'WARN' "$noHeader of $($nistTables.Count) table(s) have no w:tblHeader on row 1" }
        else                 { Add-Result 'TABLES' 'header row repeats' 'PASS' "$($nistTables.Count) table(s)" }
        if ($barExtra -gt 0) {
            $bits = @()
            $bits += "$barExtra table(s) have a partial bar - $barGaps row(s) missing it"
            Add-Result 'TABLES' 'accent bar runs every row where present' 'WARN' ($bits -join '; ')
        } else {
            Add-Result 'TABLES' 'accent bar runs every row where present' 'PASS' "$popTables barred table(s)"
        }
        if ($badNote -gt 0) { Add-Result 'TABLES' 'footnote row unshaded, no side or bottom rule' 'WARN' "$badNote of $withNote footnote row(s) wrong" }
        else                { Add-Result 'TABLES' 'footnote row unshaded, no side or bottom rule' 'PASS' "$withNote footnote row(s)" }
    }

    # ---------------------------------------------------------------- FIELDS
    $bookmarks = @{}
    foreach ($m in [regex]::Matches($doc, '<w:bookmarkStart\b[^>]*w:name="([^"]+)"')) { $bookmarks[$m.Groups[1].Value] = $true }
    $badRef = @()
    foreach ($m in [regex]::Matches($doc, 'REF\s+(_Ref\d+|[A-Za-z_][\w]*)')) {
        if (-not $bookmarks.ContainsKey($m.Groups[1].Value)) { $badRef += $m.Groups[1].Value }
    }
    $badRef = @($badRef | Sort-Object -Unique)
    if ($badRef.Count -gt 0) {
        Add-Result 'FIELDS' 'every REF target has a bookmark' 'FAIL' "$($badRef.Count) missing: $(($badRef | Select-Object -First 8) -join ', ')"
    } else {
        Add-Result 'FIELDS' 'every REF target has a bookmark' 'PASS' "$(([regex]::Matches($doc, 'REF\s+_Ref\d+')).Count) REF field(s)"
    }

    # An empty SEQ fldSimple still evaluates and increments the sequence.
    $husks = 0
    foreach ($m in [regex]::Matches($doc, '(?s)<w:fldSimple\b[^>]*w:instr="[^"]*SEQ[^"]*"[^>]*>(.*?)</w:fldSimple>')) {
        if (-not [regex]::IsMatch($m.Groups[1].Value, '<w:t(?:\s[^>]*)?>[^<]')) { $husks++ }
    }
    foreach ($m in [regex]::Matches($doc, '<w:fldSimple\b[^>]*w:instr="[^"]*SEQ[^"]*"\s*/>')) { $husks++ }
    if ($husks -gt 0) {
        Add-Result 'FIELDS' 'no empty SEQ husks' 'FAIL' "$husks empty SEQ field(s); each one still increments the sequence on the next F9"
    } else {
        Add-Result 'FIELDS' 'no empty SEQ husks' 'PASS' 'none'
    }

    # Highlighter is review markup, not NIST formatting - it has no place in a
    # delivered chapter and prints as a grey block. Reported, never stripped:
    # highlight is emphasis, and emphasis is the author's to remove.
    $hl = @{}
    foreach ($m in [regex]::Matches($doc, '<w:highlight\s+w:val="([^"]+)"')) {
        $k = $m.Groups[1].Value; $hl[$k] = 1 + [int]$hl[$k]
    }
    $hlTotal = 0; foreach ($k in $hl.Keys) { $hlTotal += $hl[$k] }
    if ($hlTotal -gt 0) {
        Add-Result 'CONTENT' 'no highlighter left' 'WARN' ("$hlTotal run(s): " + (($hl.Keys | Sort-Object | ForEach-Object { "$_ x$($hl[$_])" }) -join ', '))
    } else {
        Add-Result 'CONTENT' 'no highlighter left' 'PASS' 'none'
    }

    # Direct formatting on BodyText that fights the style. A direct w:sz always
    # beats the style, so a correct 10pt BodyText still renders at 11 or 12 -
    # NWIRP 2-1 Ch00 had a hard size on 100% of its body runs, and Ch04 had
    # genuine body paragraphs at 14pt. Remove-DirectRunSize.ps1 strips only
    # w:sz/w:szCs and leaves bold, italic and color alone.
    $bodyPara = 0; $bodySz = @{}; $bodySpac = 0; $bodyInd = 0
    foreach ($pm in [regex]::Matches($doc, '(?s)<w:p\b[^>]*>.*?</w:p>')) {
        $pp = $pm.Value
        $stm = [regex]::Match($pp, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        if (-not $stm.Success -or $stm.Groups[1].Value -ne 'BodyText') { continue }
        # Shape text often lands on Body Text because that is the document's
        # default, but a chart callout at 14pt or 18pt is sized to the picture,
        # not fighting the body style. Ch04's seventeen "body paragraphs at
        # 14pt and 18pt" were all inside drawings.
        if (& $inShape $pm.Index) { continue }
        $bodyPara++
        foreach ($zm in [regex]::Matches($pp, '<w:sz\s+w:val="(\d+)"\s*/>')) {
            $pt = [int]$zm.Groups[1].Value / 2
            $bodySz["$pt"] = 1 + [int]$bodySz["$pt"]
        }
        $pprm = [regex]::Match($pp, '(?s)<w:pPr>.*?</w:pPr>')
        if ($pprm.Success) {
            if ($pprm.Value -match '<w:spacing\b') { $bodySpac++ }
            if ($pprm.Value -match '<w:ind\b')     { $bodyInd++ }
        }
    }
    # 10pt matches what xxbase gives BodyText, so a direct 10pt is redundant,
    # not wrong; the certified set carries plenty. Only a DIFFERENT size fights.
    $wrongSz = 0; $wrongList = @()
    foreach ($k in $bodySz.Keys) { if ($k -ne '10') { $wrongSz += $bodySz[$k]; $wrongList += "${k}pt x$($bodySz[$k])" } }
    if ($wrongSz -gt 0) {
        Add-Result 'STYLES' 'no direct size fighting BodyText' 'WARN' (($wrongList | Sort-Object) -join ', ')
    } else {
        Add-Result 'STYLES' 'no direct size fighting BodyText' 'PASS' "$bodyPara BodyText paragraph(s)"
    }
    if ($bodyPara -gt 0 -and ($bodySpac + $bodyInd) -gt 0) {
        Add-Result 'STYLES' 'no direct spacing/indent on BodyText' 'WARN' "$bodySpac with spacing, $bodyInd with indent, of $bodyPara"
    } else {
        Add-Result 'STYLES' 'no direct spacing/indent on BodyText' 'PASS' 'none'
    }

    # Direct LINE spacing is a different animal from direct before/after.
    # xxbase gives w:line="264" auto (1.1), and the certified set carries almost
    # no direct w:line outside table cells - 673 of its 793 are TableText at
    # 259 and BodyText has exactly one. NWIRP 2-1 arrived with 522 overrides,
    # 365 of them w:line="480": whole chapters drafted double-spaced. Several
    # others are Google Docs defaults that rode in with the export (276 is
    # Google's 1.15; 257 and 259 are variants of it).
    #
    # before/after are NOT flagged: the certified set carries hundreds, so they
    # are the house style. A w:pPrChange is skipped - it records what a tracked
    # change USED to look like, not live formatting.
    $lineOv = @{}
    foreach ($pm in [regex]::Matches($doc, '(?s)<w:p\b[^>]*>.*?</w:p>')) {
        $pprm = [regex]::Match($pm.Value, '(?s)<w:pPr>.*?</w:pPr>')
        if (-not $pprm.Success) { continue }
        $whole = $pprm.Value
        $cut = $whole.IndexOf('<w:pPrChange')
        $headp = if ($cut -lt 0) { $whole } else { $whole.Substring(0, $cut) }
        $spm = [regex]::Match($headp, '<w:spacing\b([^>]*)/>')
        if (-not $spm.Success -or $spm.Groups[1].Value -notmatch 'w:line="(\d+)"') { continue }
        $lv = [regex]::Match($spm.Groups[1].Value, 'w:line="(\d+)"').Groups[1].Value
        $sty = [regex]::Match($headp, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        $sid = if ($sty.Success) { $sty.Groups[1].Value } else { '(none)' }
        # TableText legitimately carries tight line spacing; 480 never does.
        if ($sid -eq 'TableText' -and $lv -ne '480') { continue }
        $k = "$sid@$lv"; $lineOv[$k] = 1 + [int]$lineOv[$k]
    }
    $lineTotal = 0; foreach ($k in $lineOv.Keys) { $lineTotal += $lineOv[$k] }
    if ($lineTotal -gt 0) {
        $top = ($lineOv.Keys | Sort-Object { -$lineOv[$_] } | Select-Object -First 4 |
                ForEach-Object { "$_ x$($lineOv[$_])" }) -join ', '
        Add-Result 'STYLES' 'no direct line spacing' 'WARN' "$lineTotal override(s): $top"
    } else {
        Add-Result 'STYLES' 'no direct line spacing' 'PASS' 'none outside table cells'
    }

    # The anchored NWIRP banner table belongs on the first page of EVERY
    # chapter, before the Heading1. Identify it by PAGE anchoring
    # (vertAnchor=page + horzAnchor=page), not by being floated at all: authors
    # also float content tables with vertAnchor="text", and treating those as
    # banners skipped 2 real data tables in NWIRP 2-1 Ch03.
    $bannerRx = '<w:tblpPr[^>]*w:vertAnchor="page"[^>]*w:horzAnchor="page"|<w:tblpPr[^>]*w:horzAnchor="page"[^>]*w:vertAnchor="page"'
    $h1At = $doc.IndexOf('w:val="Heading1"')
    $head = if ($h1At -gt 0) { $doc.Substring(0, $h1At) } else { $doc }
    $bannerCount = ([regex]::Matches($head, $bannerRx)).Count
    if ($bannerCount -eq 1) {
        Add-Result 'LAYOUT' 'banner table on the first chapter page' 'PASS' 'one, before the Heading 1'
    } elseif ($bannerCount -eq 0) {
        Add-Result 'LAYOUT' 'banner table on the first chapter page' 'FAIL' 'none before the Heading 1'
    } else {
        Add-Result 'LAYOUT' 'banner table on the first chapter page' 'WARN' "$bannerCount page-anchored tables before the Heading 1"
    }

    # ListParagraph is Word's generic list style and carries no glyph of its
    # own. Items belong on List Bullet / List Number, whose style supplies the
    # numbering - the certified set's items carry NO direct numPr.
    $lpCount = ([regex]::Matches($doc, '<w:pStyle w:val="ListParagraph"\s*/>')).Count
    if ($lpCount -gt 0) {
        Add-Result 'LISTS' 'no ListParagraph' 'WARN' "$lpCount on ListParagraph; List Bullet / List Number carry the glyph"
    } else {
        Add-Result 'LISTS' 'no ListParagraph' 'PASS' 'none'
    }

    # Content tables go on NIST-Data; only the page-anchored banner stays on
    # TableGrid. A content table left on TableGrid draws black box rules instead
    # of the 808080 themeShade grid.
    $tblAll = @([regex]::Matches($doc, '(?s)<w:tbl>.*?</w:tbl>'))
    $offStyle = 0; $contentTbl = 0
    foreach ($t in $tblAll) {
        if ([regex]::IsMatch($t.Value, $bannerRx)) { continue }
        $contentTbl++
        if ($t.Value -notmatch '<w:tblStyle w:val="(NIST-Data|HM-Data|SidebarTable)"\s*/>') { $offStyle++ }
    }
    if ($contentTbl -eq 0) {
        Add-Result 'TABLES' 'content tables on NIST-Data' 'PASS' 'no content tables'
    } elseif ($offStyle -gt 0) {
        Add-Result 'TABLES' 'content tables on NIST-Data' 'WARN' "$offStyle of $contentTbl not on NIST-Data/HM-Data/SidebarTable"
    } else {
        Add-Result 'TABLES' 'content tables on NIST-Data' 'PASS' "$contentTbl table(s)"
    }

    # Tracked changes and comments are REPORTED, not failed. Authors work in
    # these documents with track changes on, so revisions and comments are the
    # normal mid-work state - treating them as a defect made every actively
    # edited chapter look broken. What they do affect is what a formatting pass
    # can safely touch: a citation field whose instrText sits inside a w:del
    # cannot be unlinked without falsifying the revision, and the rendered text
    # (so page counts and any PDF) depends on what gets accepted. Worth knowing
    # before a final production run, not worth blocking on.
    $insCount = ([regex]::Matches($doc, '<w:ins\b')).Count
    $delCount = ([regex]::Matches($doc, '<w:del\b')).Count
    $pprCount = ([regex]::Matches($doc, '<w:pPrChange\b')).Count
    $cmtCount = 0
    if ($parts.ContainsKey('word/comments.xml')) {
        $cmtCount = ([regex]::Matches($parts['word/comments.xml'], '<w:comment\b')).Count
    }
    if ($insCount + $delCount + $cmtCount -gt 0) {
        $bits = @()
        if ($insCount) { $bits += "$insCount insertion(s)" }
        if ($delCount) { $bits += "$delCount deletion(s)" }
        if ($pprCount) { $bits += "$pprCount format revision(s)" }
        if ($cmtCount) { $bits += "$cmtCount comment(s)" }
        Add-Result 'CONTENT' 'tracked changes and comments (informational)' 'WARN' ($bits -join ', ')
    } else {
        Add-Result 'CONTENT' 'tracked changes and comments (informational)' 'PASS' 'none'
    }

    # Images with no caption anywhere in the document. The 'caption on the
    # matching style' check only inspects captions that EXIST, so a chapter
    # whose figure captions were never styled passes it vacuously - NWIRP 2-1
    # Ch02 had 12 images and zero caption paragraphs and looked clean.
    $imgCount = ([regex]::Matches($doc, '<a:blip\b')).Count
    $capCount = ([regex]::Matches($doc, '<w:pStyle w:val="(?:Caption|CaptionFigure|CaptionTable)"\s*/>')).Count
    if ($imgCount -gt 0 -and $capCount -eq 0) {
        Add-Result 'CAPTIONS' 'figures have captions at all' 'FAIL' "$imgCount image(s), no caption paragraph anywhere"
    } elseif ($imgCount -gt 0 -and $capCount -lt $imgCount) {
        Add-Result 'CAPTIONS' 'figures have captions at all' 'WARN' "$imgCount image(s) but only $capCount caption paragraph(s)"
    } else {
        Add-Result 'CAPTIONS' 'figures have captions at all' 'PASS' "$imgCount image(s), $capCount caption(s)"
    }

    # --------------------------------------------------------------- CONTENT
    $markers = @{}
    # Allow descriptive text after the keyword: a useful placeholder says what
    # it is a placeholder FOR, e.g. '[PLACEHOLDER - figure title]'. Requiring
    # the brackets to hold nothing but the keyword meant the 39 placeholder
    # captions this suite's own tooling inserts were reported as 'none'.
    foreach ($m in [regex]::Matches($doc, '(?i)\[\s*(PENDING\s+FORMAT|FORMAT\s+PENDING|PLACEHOLDER|TBD|TK)\b[^\]]{0,60}\]')) { $markers[$m.Value] = 1 + [int]$markers[$m.Value] }
    $markerCount = 0; foreach ($k in $markers.Keys) { $markerCount += $markers[$k] }
    if ($markerCount -gt 0) {
        Add-Result 'CONTENT' 'no review markers left' 'WARN' "${markerCount}: $(($markers.Keys | Sort-Object) -join ', ')"
    } else {
        Add-Result 'CONTENT' 'no review markers left' 'PASS' 'none'
    }

    $core = $parts['docProps/core.xml']
    $title = if ($core) { [regex]::Match($core, '(?s)<dc:title>(.*?)</dc:title>').Groups[1].Value } else { '' }
    if (-not $title.Trim()) {
        Add-Result 'CONTENT' 'dc:title set' 'WARN' 'empty. The footer TITLE field reads this; run Set-DocTitle.ps1 before Set-NISTHeaderFooterContent.ps1.'
    } else {
        Add-Result 'CONTENT' 'dc:title set' 'PASS' $title
    }


    # ------------------------------------------------- CAPTIONS, extra checks
    # Every check below exists because a reviewer found it by hand after this
    # script reported "no failures".

    # Caption must not be italic. Latin italic is w:i; w:iCs is complex-script
    # only and is present in every file including the NCSTAR reference, so it is
    # not a defect.
    $capItal = @()
    foreach ($sid in @('Caption', 'CaptionTable', 'CaptionFigure', 'CaptionFigure2')) {
        if (-not $styles.ContainsKey($sid)) { continue }
        if ([regex]::IsMatch($styles[$sid], '<w:i/>|<w:i w:val="(?:1|true|on)"')) { $capItal += $sid }
    }
    if ($capItal.Count -gt 0) {
        Add-Result 'CAPTIONS' 'caption styles not italic' 'FAIL' "italic on: $($capItal -join ', ')"
    } else { Add-Result 'CAPTIONS' 'caption styles not italic' 'PASS' 'none italic' }

    # Figure captions go BELOW the figure, table captions ABOVE the table.
    $figAbove = @()
    for ($i = 0; $i -lt $blocks.Count; $i++) {
        $b = $blocks[$i].Value
        if ((Get-Style -Block $b) -ne 'CaptionFigure') { continue }
        if (-not (Get-Text -Block $b).Trim()) { continue }
        # A caption paragraph that holds its OWN image is not misplaced: some
        # authors put the figure and its caption in a single paragraph, so there
        # is no figure above to find and the next figure below is the NEXT one.
        if ($b -match '<w:drawing>|<w:pict>') { continue }
        $hasImgUp = $false; $hasImgDn = $false
        for ($j = $i-1; $j -ge [Math]::Max(0,$i-4); $j--) {
            if ($blocks[$j].Value -match '<w:drawing>|<w:pict>') { $hasImgUp = $true; break }
            # Skip a short panel label - '(b)', 'Observed NDWE', 'Exposure B' -
            # which sits between a multi-panel figure and its caption. Stopping
            # at the first non-empty block read those as "no figure above".
            $tj = (Get-Text -Block $blocks[$j].Value).Trim()
            if ($tj -and $tj.Length -gt 30) { break }
        }
        for ($j = $i+1; $j -lt [Math]::Min($blocks.Count,$i+4); $j++) {
            if ($blocks[$j].Value -match '<w:drawing>|<w:pict>') { $hasImgDn = $true; break }
            if ((Get-Text -Block $blocks[$j].Value).Trim()) { break }
        }
        if (-not $hasImgUp -and $hasImgDn) { $figAbove += (Get-Text -Block $b).Substring(0,[Math]::Min(44,(Get-Text -Block $b).Length)) }
    }
    if ($figAbove.Count -gt 0) {
        Add-Result 'CAPTIONS' 'figure caption sits below its figure' 'FAIL' ("$($figAbove.Count): " + ($figAbove -join '; '))
    } else { Add-Result 'CAPTIONS' 'figure caption sits below its figure' 'PASS' 'ok' }

    # Caption numbers should be fields, not typed. A typed number does not
    # renumber when a figure moves.
    $typed = 0; $fielded = 0
    foreach ($bm in $blocks) {
        $b = $bm.Value
        if ((Get-Style -Block $b) -notmatch '^Caption(Table|Figure)$') { continue }
        if (-not (Get-Text -Block $b).Trim()) { continue }
        if ($b -match '<w:instrText|<w:fldSimple') { $fielded++ } else { $typed++ }
    }
    if ($typed -gt 0) {
        Add-Result 'CAPTIONS' 'caption numbers are fields' 'WARN' "$typed typed, $fielded fielded. Typed numbers do not renumber."
    } elseif ($fielded -gt 0) {
        Add-Result 'CAPTIONS' 'caption numbers are fields' 'PASS' "$fielded field(s)"
    }

    # ------------------------------------------------- HEADERS AND FOOTERS
    $hfBad = @()
    foreach ($key in $parts.Keys) {
        if ($key -notmatch '^word/(header|footer)\d*\.xml$') { continue }
        $x = $parts[$key]
        if (-not [regex]::IsMatch($x, '<w:t(?:\s[^>]*)?>[^<]') -and $x -notmatch '<w:instrText') { continue }
        $want = if ($key -match 'header') { '^(HeaderOdd|HeaderEven|Headerch|Header)$' } else { '^(FooterOdd|FooterEven|Footer)$' }
        $found = @([regex]::Matches($x, '<w:pStyle\b[^>]*w:val="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        if ($found.Count -eq 0) { $hfBad += "$(Split-Path -Leaf $key) on Normal (no pStyle)" }
        else {
            foreach ($st in $found) { if ($st -notmatch $want) { $hfBad += "$(Split-Path -Leaf $key) on $st" } }
        }
    }
    if ($hfBad.Count -gt 0) {
        Add-Result 'HEADER/FOOTER' 'header and footer paragraphs styled' 'FAIL' ($hfBad -join '; ')
    } else { Add-Result 'HEADER/FOOTER' 'header and footer paragraphs styled' 'PASS' 'ok' }

    # ------------------------------------------------- LISTS
    $listBad = @()
    foreach ($sid in @('ListBullet','ListBullet2','ListBullet3','ListBullet4','ListBullet5',
                       'ListNumber','ListNumber2','ListNumber3','ListNumber4','ListNumber5')) {
        if (-not $styles.ContainsKey($sid)) { continue }
        $used = ([regex]::Matches($doc, '<w:pStyle\b[^>]*w:val="' + $sid + '"')).Count
        $np = [regex]::Match($styles[$sid], '(?s)<w:numPr>.*?<w:numId\b[^>]*w:val="(\d+)"')
        if (-not $np.Success -or $np.Groups[1].Value -eq '0') {
            if ($used -gt 0) { $listBad += "$sid has no numbering ($used paragraph(s) use it)" }
            else { $listBad += "$sid has no numbering (unused)" }
        } elseif (-not $numIds.ContainsKey($np.Groups[1].Value)) {
            $listBad += "$sid -> numId $($np.Groups[1].Value) not in numbering.xml"
        }
    }
    if ($listBad.Count -gt 0) {
        $hard = @($listBad | Where-Object { $_ -notmatch 'unused' })
        Add-Result 'LISTS' 'list styles have a bullet or number' $(if ($hard.Count) { 'FAIL' } else { 'WARN' }) ($listBad -join '; ')
    } else { Add-Result 'LISTS' 'list styles have a bullet or number' 'PASS' 'ok' }

    # ------------------------------------------------- REFERENCES
    $refHead = $null
    for ($i = 0; $i -lt $blocks.Count; $i++) {
        if ((Get-Text -Block $blocks[$i].Value).Trim().ToLower() -eq 'references') { $refHead = $i }
    }
    if ($refHead -ne $null) {
        $refEnts = @(); $j = $refHead + 1
        while ($j -lt $blocks.Count) {
            $st = Get-Style -Block $blocks[$j].Value
            # Stop at the next section, not just at a Heading: the reference list
            # is followed by AppendixHead / Front_matter_head in these chapters,
            # and running past it counts appendix placeholders as references.
            #
            # An EMPTY heading is a boundary too. Requiring text here made the
            # scan run past both empty AppendixHead paragraphs at the end of
            # NWIRP 2-1 Ch06 - one is the closing paragraph of a multi-paragraph
            # bibliography field, the other is Appendix B's untitled heading -
            # and on into that chapter's mathematical nomenclature, counting
            # "I = Set of nodes" and "Parameters" as reference entries that
            # were not on the References style. 71 real entries reported as
            # 63 of 97.
            if ($st -match '^(Heading[1-9]|AppendixHead|Appendix_Head\d*|Frontmatterhead\d*|Front_matter_head\d*|Heading1NoNumber)$') { break }
            if ((Get-Text -Block $blocks[$j].Value).Trim()) { $refEnts += $blocks[$j].Value }
            $j++
        }
        $onRef = @($refEnts | Where-Object { (Get-Style -Block $_) -eq 'References' }).Count
        if ($onRef -lt $refEnts.Count) {
            Add-Result 'REFERENCES' 'entries on the References style' 'FAIL' "$onRef of $($refEnts.Count) on References"
        } else { Add-Result 'REFERENCES' 'entries on the References style' 'PASS' "$onRef entries" }
        $refItal = 0
        foreach ($e in $refEnts) { $refItal += ([regex]::Matches($e, '<w:i/>')).Count }
        if ($refEnts.Count -gt 0 -and $refItal -eq 0) {
            Add-Result 'REFERENCES' 'journal and book titles italic' 'WARN' "$($refEnts.Count) entries, 0 italic runs. APA italicises the source title; a conversion flattens it."
        } elseif ($refEnts.Count -gt 0) {
            Add-Result 'REFERENCES' 'journal and book titles italic' 'PASS' "$refItal italic run(s) across $($refEnts.Count) entries"
        }
    }

    # ------------------------------------------------- CITATION MANAGER
    $cm = 0
    foreach ($key in $parts.Keys) {
        if ($key -eq 'word/comments.xml') { continue }
        $cm += ([regex]::Matches($parts[$key], 'ZOTERO|MENDELEY|EN\.CITE|ADDIN CSL')).Count
    }
    if ($cm -gt 0) {
        Add-Result 'FIELDS' 'no citation-manager linking' 'FAIL' "$cm reference(s) to Zotero/Mendeley/EndNote outside comments"
    } else { Add-Result 'FIELDS' 'no citation-manager linking' 'PASS' 'none' }

    # --------------------------------------------------- INTEGRITY
    # Every check below came out of a real failure found in NWIRP 2-3.

    # A bare "REF bookmark \h" pastes the ENTIRE bookmark span, not a number.
    # Six caption bookmarks in Chapter 4 never closed at their caption - their
    # bookmarkEnd sat thousands of paragraphs downstream - so each span ran to
    # ~1.75 MB and every F9 injected all of it back into the document. One
    # update took that chapter from 6,257 paragraphs to 22,958.
    $bmStart = @{}
    foreach ($m in [regex]::Matches($doc, '<w:bookmarkStart\b[^>]*w:id="(\d+)"[^>]*w:name="([^"]+)"')) {
        $bmStart[$m.Groups[2].Value] = @{ Pos = $m.Index + $m.Length; Id = $m.Groups[1].Value }
    }
    $bmEnd = @{}
    foreach ($m in [regex]::Matches($doc, '<w:bookmarkEnd\b[^>]*w:id="(\d+)"[^>]*/>')) {
        if (-not $bmEnd.ContainsKey($m.Groups[1].Value)) { $bmEnd[$m.Groups[1].Value] = $m.Index }
    }
    $runaway = @()
    foreach ($m in [regex]::Matches($doc, '(?s)<w:instrText(?:\s[^>]*)?>\s*(?:PAGE|NOTE)?REF\s+(\S+)([^<]*)</w:instrText>')) {
        if ($m.Groups[2].Value -match '\\r') { continue }   # \r inserts a number, not the span
        $nm = $m.Groups[1].Value
        if (-not $bmStart.ContainsKey($nm)) { continue }
        $id = $bmStart[$nm].Id
        if (-not $bmEnd.ContainsKey($id)) { continue }
        $span = $bmEnd[$id] - $bmStart[$nm].Pos
        if ($span -gt 20000) { $runaway += ('{0} ({1:N0} chars)' -f $nm, $span) }
    }
    if ($runaway.Count -gt 0) {
        Add-Result 'INTEGRITY' 'no REF field pastes a runaway bookmark' 'FAIL' (
            "$($runaway.Count): " + (($runaway | Select-Object -First 4) -join '; ') +
            '. Updating fields will duplicate that span into the document.')
    } else {
        Add-Result 'INTEGRITY' 'no REF field pastes a runaway bookmark' 'PASS' 'none over 20,000 chars'
    }

    # Bookmarks must pair up. A duplicate bookmarkEnd or an end with no start is
    # how a hand-patched caption breaks Word's revision and reference tracking.
    $sCount = ([regex]::Matches($doc, '<w:bookmarkStart\b')).Count
    $eCount = ([regex]::Matches($doc, '<w:bookmarkEnd\b')).Count
    $endIds = @{}
    foreach ($m in [regex]::Matches($doc, '<w:bookmarkEnd\b[^>]*w:id="(\d+)"')) {
        $k = $m.Groups[1].Value
        $endIds[$k] = 1 + $(if ($endIds.ContainsKey($k)) { $endIds[$k] } else { 0 })
    }
    $dupEnds = @($endIds.GetEnumerator() | Where-Object { $_.Value -gt 1 })
    $startIds = @{}
    foreach ($m in [regex]::Matches($doc, '<w:bookmarkStart\b[^>]*w:id="(\d+)"')) { $startIds[$m.Groups[1].Value] = $true }
    $orphanEnds = @($endIds.Keys | Where-Object { -not $startIds.ContainsKey($_) })
    $startCounts = @{}
    foreach ($m in [regex]::Matches($doc, '<w:bookmarkStart\b[^>]*w:id="(\d+)"')) {
        $k = $m.Groups[1].Value
        $startCounts[$k] = 1 + $(if ($startCounts.ContainsKey($k)) { $startCounts[$k] } else { 0 })
    }
    $dupStarts = @($startCounts.GetEnumerator() | Where-Object { $_.Value -gt 1 })
    if ($sCount -ne $eCount -or $dupEnds.Count -gt 0 -or $orphanEnds.Count -gt 0 -or $dupStarts.Count -gt 0) {
        $detail = "$sCount start / $eCount end"
        if ($dupStarts.Count -gt 0) { $detail += "; $($dupStarts.Count) duplicate start id(s): " + (($dupStarts | Select-Object -First 6 | ForEach-Object { $_.Key }) -join ', ') }
        if ($dupEnds.Count   -gt 0) { $detail += "; $($dupEnds.Count) duplicate end id(s)" }
        if ($orphanEnds.Count -gt 0) { $detail += "; $($orphanEnds.Count) end(s) with no start" }
        Add-Result 'INTEGRITY' 'bookmarks pair up' 'FAIL' $detail
    } else {
        Add-Result 'INTEGRITY' 'bookmarks pair up' 'PASS' "$sCount pair(s)"
    }

    # An unterminated field swallows everything after it: Word treats the rest of
    # the document as field content. Four of these sat in Chapter 4.
    $depth = 0; $wentNegative = $false
    foreach ($m in [regex]::Matches($doc, 'w:fldCharType="(begin|end)"')) {
        if ($m.Groups[1].Value -eq 'begin') { $depth++ } else { $depth-- }
        if ($depth -lt 0) { $wentNegative = $true }
    }
    if ($depth -ne 0 -or $wentNegative) {
        Add-Result 'INTEGRITY' 'field characters balance' 'FAIL' (
            "$depth unterminated field(s)" + $(if ($wentNegative) { ', and an end appeared before its begin' } else { '' }))
    } else {
        Add-Result 'INTEGRITY' 'field characters balance' 'PASS' 'balanced'
    }

    # Google Docs export residue. The goog_rdk content controls nest paragraphs
    # inside w:sdt, which breaks any tool that walks w:p directly.
    $goog = ([regex]::Matches($doc, 'goog_rdk')).Count
    $cxml = ([regex]::Matches($doc, '<w:customXml(?:Ins|Del)Range')).Count
    $gbm  = ([regex]::Matches($doc, 'w:name="(?:_heading=)?h\.|w:name="bookmark=id\.')).Count
    $glink = ([regex]::Matches($doc, '1155CC')).Count
    $gTotal = $goog + $cxml + $gbm + $glink
    if ($gTotal -gt 0) {
        Add-Result 'INTEGRITY' 'no Google Docs export residue' 'WARN' (
            "$goog goog_rdk control(s), $cxml customXml range marker(s), $gbm Google bookmark(s), $glink link-blue run(s)")
    } else {
        Add-Result 'INTEGRITY' 'no Google Docs export residue' 'PASS' 'none'
    }

    # Captions are 10pt (sz=20) and the style owns it. Chapter 6's Caption
    # declared only szCs, so it inherited 12pt from xxheader and every caption
    # in the chapter was a size too big.
    $capBlock = [regex]::Match($stylesX, '(?s)<w:style\b[^>]*w:styleId="Caption"[^>]*>.*?</w:style>')
    $capSz = if ($capBlock.Success) { [regex]::Match($capBlock.Value, '<w:sz\b[^>]*w:val="(\d+)"') } else { $null }
    $directCap = 0
    foreach ($m in [regex]::Matches($doc, '(?s)<w:p(?:\s[^>]*)?>.*?</w:p>')) {
        if ($m.Value -notmatch '<w:pStyle\b[^>]*w:val="Caption[A-Za-z]*"') { continue }
        $directCap += ([regex]::Matches($m.Value, '<w:sz\b[^>]*w:val="\d+"')).Count
    }
    $capParas = ([regex]::Matches($doc, '<w:pStyle\b[^>]*w:val="Caption[A-Za-z]*"')).Count
    if (-not $capBlock.Success -and $capParas -eq 0) {
        # nothing to check: no Caption style and nothing using one
    } elseif (-not $capSz -or -not $capSz.Success) {
        Add-Result 'CAPTIONS' 'caption size is 10pt and style-owned' 'FAIL' 'Caption declares no w:sz, so it inherits (xxheader is 12pt)'
    } elseif ($capSz.Groups[1].Value -ne '20') {
        Add-Result 'CAPTIONS' 'caption size is 10pt and style-owned' 'FAIL' ("Caption is {0}pt, expected 10pt" -f ([int]$capSz.Groups[1].Value / 2))
    } elseif ($directCap -gt 0) {
        Add-Result 'CAPTIONS' 'caption size is 10pt and style-owned' 'WARN' "$directCap direct size(s) on caption paragraphs override the style"
    } else {
        Add-Result 'CAPTIONS' 'caption size is 10pt and style-owned' 'PASS' '10pt, no direct overrides'
    }

    # Near-miss palette colors. 002952 was on 345 caption runs across the two
    # Chapter 3 copies - two shades off the navy every other chapter used.
    $nearMiss = @{}
    foreach ($bad in @('002952','02257A','002060','366091')) {
        $c = ([regex]::Matches($doc + $stylesX, 'w:(?:val|fill|color)="' + $bad + '"', 'IgnoreCase')).Count
        if ($c -gt 0) { $nearMiss[$bad] = $c }
    }
    if ($nearMiss.Count -gt 0) {
        Add-Result 'STYLES' 'palette colors are canonical' 'WARN' (
            (($nearMiss.GetEnumerator() | ForEach-Object { "$($_.Key) x$($_.Value)" }) -join ', ') +
            '. Canonical navy is 00257A; mid-blue is 355195.')
    } else {
        Add-Result 'STYLES' 'palette colors are canonical' 'PASS' 'no near-miss values'
    }

    # Alt text. Chapter 6 shipped 9 of 10 images with none.
    # Evaluate the ACCEPTED document: an image inside a tracked deletion is on
    # its way out, so it is not an accessibility gap.
    $accepted = $doc
    while ($true) {
        $next = [regex]::Replace($accepted, '(?s)<w:del\b[^>]*[^/]>.*?</w:del>', '')
        if ($next -eq $accepted) { break }
        $accepted = $next
    }
    $imgs = 0; $noAlt = 0
    foreach ($m in [regex]::Matches($accepted, '(?s)<w:drawing>.*?</w:drawing>')) {
        $imgs++
        $dp = [regex]::Match($m.Value, '<wp:docPr\b[^>]*>')
        if (-not $dp.Success -or $dp.Value -notmatch 'descr="[^"]' ) { $noAlt++ }
    }
    if ($noAlt -gt 0) {
        Add-Result 'CONTENT' 'every image has alt text' 'FAIL' "$noAlt of $imgs image(s) have none (508 requirement)"
    } elseif ($imgs -gt 0) {
        Add-Result 'CONTENT' 'every image has alt text' 'PASS' "$imgs image(s)"
    }

    # Auto-width tables re-derive their widths from content, so tables sharing a
    # grid render at different sizes.
    $tbls = [regex]::Matches($doc, '(?s)<w:tbl>.*?</w:tbl>')
    $auto = 0
    foreach ($t in $tbls) { if ($t.Value -match '<w:tblW\b[^>]*w:type="auto"') { $auto++ } }
    if ($auto -gt 0) {
        Add-Result 'TABLES' 'tables have a fixed width' 'WARN' "$auto of $($tbls.Count) on auto width; Set-TableLook.ps1 -FixedWidth pins them"
    } elseif ($tbls.Count -gt 0) {
        Add-Result 'TABLES' 'tables have a fixed width' 'PASS' "$($tbls.Count) table(s)"
    }

    # ------------------------------------------------- LAYOUT / HYGIENE
    # Each of these came out of a defect found in NWIRP 2-3.

    # Frames. Never used in these reports, and they sneak in on Title, where
    # they anchor it to the page and float it out of the text flow.
    $frameStyles = @()
    foreach ($m in [regex]::Matches($stylesX, '(?s)<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>.*?</w:style>')) {
        if ($m.Value -match '<w:framePr\b') { $frameStyles += $m.Groups[1].Value }
    }
    $frameDirect = ([regex]::Matches($doc, '<w:framePr\b')).Count
    if ($frameStyles.Count -gt 0 -or $frameDirect -gt 0) {
        Add-Result 'LAYOUT' 'no text frames' 'WARN' (
            "$($frameStyles.Count) style(s): $(($frameStyles | Select-Object -First 6) -join ', '); $frameDirect direct on paragraphs")
    } else { Add-Result 'LAYOUT' 'no text frames' 'PASS' 'none' }

    # xxgraphic is not part of the base architecture - it never appears in the
    # original nist-scripts and nothing inherits it.
    if ($stylesX -match '<w:style\b[^>]*w:styleId="xxgraphic"') {
        Add-Result 'STYLES' 'no xxgraphic style' 'WARN' 'present; it is toolkit drift, used by nothing'
    } else { Add-Result 'STYLES' 'no xxgraphic style' 'PASS' 'absent' }

    # Page setup: Letter, 1 inch all round, header and footer 0.3 inch.
    $pgBad = @()
    foreach ($m in [regex]::Matches($doc, '(?s)<w:sectPr\b[^>]*>.*?</w:sectPr>')) {
        $sz = [regex]::Match($m.Value, '<w:pgSz\b[^>]*w:w="(\d+)"[^>]*w:h="(\d+)"')
        $mg = [regex]::Match($m.Value, '<w:pgMar\b[^>]*w:top="(-?\d+)"[^>]*w:right="(\d+)"[^>]*w:bottom="(-?\d+)"[^>]*w:left="(\d+)"')
        $hd = [regex]::Match($m.Value, 'w:header="(\d+)"'); $ft = [regex]::Match($m.Value, 'w:footer="(\d+)"')
        if (-not ($sz.Success -and $mg.Success)) { continue }
        $bad = @()
        if ($sz.Groups[1].Value -ne '12240' -or $sz.Groups[2].Value -ne '15840') { $bad += 'page size' }
        foreach ($g in 1..4) { if ($mg.Groups[$g].Value -ne '1440') { $bad += 'margin'; break } }
        if ($hd.Success -and $hd.Groups[1].Value -ne '432') { $bad += "header $([int]$hd.Groups[1].Value/1440)in" }
        if ($ft.Success -and $ft.Groups[1].Value -ne '432') { $bad += "footer $([int]$ft.Groups[1].Value/1440)in" }
        if ($bad.Count) { $pgBad += ($bad -join '+') }
    }
    if ($pgBad.Count) {
        Add-Result 'LAYOUT' 'page setup 1in margins, 0.3in header/footer' 'FAIL' ("$($pgBad.Count) section(s): " + ($pgBad -join '; '))
    } else { Add-Result 'LAYOUT' 'page setup 1in margins, 0.3in header/footer' 'PASS' 'Letter, 1in, 0.3/0.3' }

    # An unstyled empty paragraph takes its height from Normal, which is Times
    # New Roman 12pt in most of these files - a taller gap than it looks.
    $unstyled = 0
    foreach ($m in [regex]::Matches($doc, '(?s)<w:p(?:\s[^>]*)?>.*?</w:p>')) {
        $v = $m.Value
        if ($v -match '<w:pStyle\b' -or $v -match '<w:drawing' -or $v -match '<w:pict') { continue }
        if ($v -match '<w:br w:type="page"' -or $v -match '<w:sectPr') { continue }
        if ([regex]::Match($v, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>').Groups[1].Value.Trim()) { continue }
        $unstyled++
    }
    # INFORMATIONAL, not a defect. Measured against the last certified chapter
    # set, 128 of its 149 empty paragraphs carry no style at all, so leaving
    # them unstyled is the house practice rather than an oversight - the
    # classifier in Get-UnstyledContent.ps1 maps "empty" to "leave on Normal"
    # for that reason. What DID matter was the direct formatting on them:
    # seventy-one empty paragraphs in NWIRP 2-1 carried their own double line
    # spacing, so every blank line was a different height, and that is covered
    # by 'no direct line spacing' instead.
    if ($unstyled -gt 0) {
        Add-Result 'LAYOUT' 'empty paragraphs carry a style' 'PASS' "$unstyled unstyled (informational; the certified set leaves them that way)"
    } else { Add-Result 'LAYOUT' 'empty paragraphs carry a style' 'PASS' 'none unstyled' }

    # A figure outside a table belongs on the graphic style - unless the
    # paragraph is a CAPTION that holds its own image. Two figures in NWIRP 2-1
    # Ch05 are built that way, with the picture and the caption text in one
    # paragraph, and of the two descriptions the caption is the more specific:
    # moving it to graphic would strip the caption styling off a real caption.
    # Text boxes are skipped for the same reason as everywhere else - an image
    # inside a shape is part of the drawing, not a figure in the text flow.
    $imgBad = 0
    $txbxSpansImg = @([regex]::Matches($doc, '(?s)<w:txbxContent>.*?</w:txbxContent>'))
    foreach ($m in [regex]::Matches($doc, '(?s)<w:p(?:\s[^>]*)?>.*?</w:p>')) {
        if ($m.Value -notmatch '<w:drawing' -and $m.Value -notmatch '<w:pict') { continue }
        $before = $doc.Substring(0, $m.Index)
        $open = ([regex]::Matches($before, '<w:tbl>')).Count - ([regex]::Matches($before, '</w:tbl>')).Count
        if ($open -gt 0) { continue }
        $inShape = $false
        foreach ($sp in $txbxSpansImg) { if ($m.Index -ge $sp.Index -and $m.Index -lt ($sp.Index + $sp.Length)) { $inShape = $true; break } }
        if ($inShape) { continue }
        if ($m.Value -match '<w:pStyle\b[^>]*w:val="Caption\w*"') { continue }
        if ($m.Value -notmatch '<w:pStyle\b[^>]*w:val="graphic"') { $imgBad++ }
    }
    if ($imgBad -gt 0) { Add-Result 'LAYOUT' 'figures on the graphic style' 'WARN' "$imgBad outside a table are not on graphic" }
    else { Add-Result 'LAYOUT' 'figures on the graphic style' 'PASS' 'ok' }

    # These chapters carry no table of contents; the report builds one.
    $tocField = 0
    foreach ($m in [regex]::Matches($doc, '(?s)<w:instrText(?:\s[^>]*)?>(.*?)</w:instrText>')) {
        if ($m.Groups[1].Value -match '\bTOC\b') { $tocField++ }
    }
    $tocPara = ([regex]::Matches($doc, '<w:pStyle\b[^>]*w:val="(?:TOC\d|TOCHeading)"')).Count
    if ($tocField -gt 0 -or $tocPara -gt 0) {
        Add-Result 'CONTENT' 'no table of contents in the chapter' 'WARN' "$tocField TOC field(s), $tocPara TOC paragraph(s)"
    } else { Add-Result 'CONTENT' 'no table of contents in the chapter' 'PASS' 'none' }

    # The running head should name this chapter.
    #
    # Find header parts through the RELATIONSHIPS, not the filename. Word's
    # default naming is header1.xml, but a template may name them anything:
    # NCSTAR 4-3 uses nistHeaderDefault.xml, and matching the filename pattern
    # found no parts at all and called a perfectly good header empty.
    #
    # And judge the FIELD, not its cached result. The right way to build a
    # running head is STYLEREF "Heading 1", which Word recomputes per page from
    # the heading actually above it. The literal text sitting between
    # fldCharType separate and end is only the last cached value, so a correct
    # header can read "Chapter Title" (NCSTAR's template default) or name some
    # other chapter entirely, and both recompute correctly on open. Only a
    # header with NO STYLEREF is worth testing against its own text.
    $hdrParts = @()
    $relsXml = $parts['word/_rels/document.xml.rels']
    if ($relsXml) {
        foreach ($m in [regex]::Matches($relsXml, '<Relationship\b[^>]*/header"[^>]*>')) {
            $t = [regex]::Match($m.Value, 'Target="([^"]+)"')
            if ($t.Success -and $t.Groups[1].Value -notmatch '^https?:') {
                $hdrParts += ('word/' + $t.Groups[1].Value.TrimStart('/'))
            }
        }
    }
    if (-not $hdrParts) { $hdrParts = @($parts.Keys | Where-Object { $_ -match '^word/header\d+\.xml$' }) }
    $hdrParts = @($hdrParts | Select-Object -Unique | Where-Object { $parts.ContainsKey($_) })

    # A first-page header is empty by NIST convention - no running head on the
    # opening page of a chapter - so it must not be counted as a missing one.
    $hdrStyleRef = $false
    $hdrText = ''
    $hdrBodies = 0
    foreach ($hp in $hdrParts) {
        $hx = $parts[$hp]
        if ($hp -match 'First') { continue }
        $hdrBodies++
        if ($hx -match '(?s)<w:instrText[^>]*>[^<]*STYLEREF[^<]*(Heading\s*1|Heading1)') { $hdrStyleRef = $true }
        foreach ($t in [regex]::Matches($hx, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')) { $hdrText += $t.Groups[1].Value + ' ' }
    }
    # The chapter number comes from the DOCUMENT, not its filename: HMList's
    # ilvl 0 w:start is the number the headings actually render, it survives a
    # rename, and it is what Set-NISTChapterNumbering.ps1 writes. Reading the
    # filename instead meant this check silently compared nothing whenever the
    # name was spelled differently - "Chapter_01" with an underscore missed the
    # old strict pattern entirely and passed by falling through.
    $chapterNo = $null
    if ($numX) {
        $hmStyle = [regex]::Match($stylesX, '(?s)<w:style\b[^>]*\bw:styleId="HMList".*?</w:style>')
        $hmNumId = if ($hmStyle.Success) { [regex]::Match($hmStyle.Value, '<w:numId\b[^>]*\bw:val="(\d+)"') } else { $null }
        $hmAbs = $null
        if ($hmNumId -and $hmNumId.Success) {
            $nm = [regex]::Match($numX, '(?s)<w:num\b[^>]*\bw:numId="' + $hmNumId.Groups[1].Value + '".*?<w:abstractNumId w:val="(\d+)"')
            if ($nm.Success) { $hmAbs = $nm.Groups[1].Value }
        }
        if (-not $hmAbs) {
            $sl = [regex]::Match($numX, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"[^>]*>(?:(?!</w:abstractNum>).)*?<w:styleLink\b[^>]*\bw:val="HMList"')
            if ($sl.Success) { $hmAbs = $sl.Groups[1].Value }
        }
        if ($hmAbs) {
            $ab = [regex]::Match($numX, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="' + $hmAbs + '"[^>]*>.*?</w:abstractNum>')
            if ($ab.Success) {
                $lv = [regex]::Match($ab.Value, '(?s)<w:lvl\b[^>]*\bw:ilvl="0"[^>]*>.*?</w:lvl>')
                if ($lv.Success) {
                    $st0 = [regex]::Match($lv.Value, '<w:start\b[^>]*\bw:val="(\d+)"')
                    if ($st0.Success -and [int]$st0.Groups[1].Value -gt 0) { $chapterNo = [int]$st0.Groups[1].Value }
                }
            }
        }
    }

    if ($hdrParts.Count -eq 0) {
        Add-Result 'HEADER/FOOTER' 'running head names the chapter' 'FAIL' 'no header part referenced by the document'
    } elseif ($hdrStyleRef) {
        Add-Result 'HEADER/FOOTER' 'running head names the chapter' 'PASS' 'STYLEREF "Heading 1" - computed per page'
    } elseif ($hdrBodies -eq 0) {
        Add-Result 'HEADER/FOOTER' 'running head names the chapter' 'FAIL' 'only a first-page header; no running head on later pages'
    } elseif (-not $hdrText.Trim()) {
        Add-Result 'HEADER/FOOTER' 'running head names the chapter' 'FAIL' 'header is empty'
    } elseif ($null -ne $chapterNo -and $hdrText -notmatch ("Chapter\s*" + $chapterNo)) {
        Add-Result 'HEADER/FOOTER' 'running head names the chapter' 'FAIL' ("literal header, expected 'Chapter $chapterNo', says: " + $hdrText.Trim().Substring(0,[Math]::Min(60,$hdrText.Trim().Length)))
    } else {
        Add-Result 'HEADER/FOOTER' 'running head names the chapter' 'PASS' ('literal: ' + $hdrText.Trim().Substring(0,[Math]::Min(40,$hdrText.Trim().Length)))
    }

    # A paragraph style based on Normal, or on nothing, inherits NOTHING from
    # the template - not the face, not the size, not the spacing - so text on
    # it renders as Normal however carefully it was applied. This is the gap
    # that let 21 large blocks of body prose through in NWIRP 2-1 Ch07: the
    # style was defined, applied consistently, and carried no direct
    # formatting, so every other check passed it. Nothing was looking at what
    # it inherited FROM.
    #
    # Normal itself and the xx* bases are the roots of the architecture and are
    # expected to have no template parent.
    $archRoots = @('Normal', 'xxbase', 'xxheader', 'xxtable', 'Title')
    $usedPara = @{}
    foreach ($m in [regex]::Matches($doc, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')) { $usedPara[$m.Groups[1].Value] = $true }
    $rootless = @()
    foreach ($m in [regex]::Matches($stylesX, '(?s)<w:style\b[^>]*w:type="paragraph"[^>]*w:styleId="([^"]+)"[^>]*>.*?</w:style>')) {
        $sid = $m.Groups[1].Value
        if ($archRoots -contains $sid) { continue }
        if (-not $usedPara.ContainsKey($sid)) { continue }
        $bo = [regex]::Match($m.Value, '<w:basedOn\b[^>]*\bw:val="([^"]+)"')
        if (-not $bo.Success) { $rootless += "$sid (no parent)"; continue }
        if ($bo.Groups[1].Value -eq 'Normal') { $rootless += "$sid (basedOn Normal)" }
    }
    if ($rootless.Count -gt 0) {
        Add-Result 'STYLES' 'styles in use descend from the template' 'FAIL' (($rootless | Select-Object -First 6) -join '; ')
    } else {
        Add-Result 'STYLES' 'styles in use descend from the template' 'PASS' "$($usedPara.Count) style(s) in use"
    }

    # Character styles nothing points at - mostly pandoc syntax-highlight junk.
    #
    # Word's own machinery styles are EXEMPT. Word applies FootnoteReference,
    # CommentReference, PageNumber, Hyperlink and the rest the moment someone
    # inserts a footnote, a comment, a page number or a link, so a chapter that
    # has none yet still needs them for the first one anyone adds, and they
    # ship semiHidden so they are not in the Styles pane either way.
    # Set-NISTStyleArchitecture.ps1 keeps them for that reason; counting them
    # here reported 11 to 13 "orphans" per file that it is correct to keep, and
    # buried the ones that are real drift ("Serif-Cambria", "relevance").
    $charMachinery = @(
        'DefaultParagraphFont', 'FootnoteReference', 'EndnoteReference',
        'CommentReference', 'PageNumber', 'Hyperlink', 'FollowedHyperlink',
        'LineNumber', 'PlaceholderText', 'UnresolvedMention', 'Mention'
    )
    $charIds = @()
    foreach ($m in [regex]::Matches($stylesX, '<w:style\b[^>]*w:type="character"[^>]*w:styleId="([^"]+)"')) { $charIds += $m.Groups[1].Value }
    $rUsed = @{}; foreach ($m in [regex]::Matches($doc, '<w:rStyle\b[^>]*w:val="([^"]+)"')) { $rUsed[$m.Groups[1].Value]=$true }
    $linkTo = @{}; foreach ($m in [regex]::Matches($stylesX, '<w:link\b[^>]*w:val="([^"]+)"')) { $linkTo[$m.Groups[1].Value]=$true }
    $basedOn = @{}; foreach ($m in [regex]::Matches($stylesX, '<w:basedOn\b[^>]*w:val="([^"]+)"')) { $basedOn[$m.Groups[1].Value]=$true }
    $orphanChar = @($charIds | Where-Object {
        -not $rUsed.ContainsKey($_) -and -not $linkTo.ContainsKey($_) -and
        -not $basedOn.ContainsKey($_) -and $charMachinery -notcontains $_ })
    if ($orphanChar.Count -gt 0) {
        Add-Result 'STYLES' 'no orphan character styles' 'WARN' "$($orphanChar.Count) of $($charIds.Count) referenced by nothing: $(($orphanChar | Select-Object -First 6) -join ', ')"
    } else { Add-Result 'STYLES' 'no orphan character styles' 'PASS' "$($charIds.Count) all referenced" }

    # Fonts are de-embedded before delivery.
    $embedded = @($parts.Keys | Where-Object { $_ -like 'word/fonts/*' }).Count
    if ($embedded -gt 0) { Add-Result 'CONTENT' 'fonts de-embedded' 'WARN' "$embedded embedded font part(s)" }
    else { Add-Result 'CONTENT' 'fonts de-embedded' 'PASS' 'none embedded' }

    # ----------------------------------------------------------------- print
    $fails = @($results | Where-Object { $_.Status -eq 'FAIL' }).Count
    $warns = @($results | Where-Object { $_.Status -eq 'WARN' }).Count
    $passes = @($results | Where-Object { $_.Status -eq 'PASS' }).Count
    $totalFail += $fails

    Write-Host ""
    Write-Host (Split-Path -Leaf $target) -ForegroundColor Cyan
    $group = ''
    foreach ($r in $results) {
        if ($Quiet -and $r.Status -eq 'PASS') { continue }
        if ($r.Group -ne $group) { $group = $r.Group; Write-Host "  $group" -ForegroundColor DarkGray }
        $color = switch ($r.Status) { 'FAIL' { 'Red' } 'WARN' { 'Yellow' } default { 'Green' } }
        Write-Host ("    [{0}] {1,-44} {2}" -f $r.Status, $r.Name, $r.Detail) -ForegroundColor $color
    }
    Write-Host ("  {0} failed, {1} warning(s), {2} passed" -f $fails, $warns, $passes) -ForegroundColor $(if ($fails) { 'Red' } elseif ($warns) { 'Yellow' } else { 'Green' })
}

Write-Host ""
if ($totalFail -gt 0) { Write-Host "$totalFail check(s) failed." -ForegroundColor Red }
else { Write-Host "No failures." -ForegroundColor Green }
exit $totalFail
