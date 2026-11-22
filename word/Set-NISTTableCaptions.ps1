<#
.SYNOPSIS
    Turns the paragraphs that sit above tables into NIST table captions, with
    an auto-numbering SEQ field.

.DESCRIPTION
    Converted documents often carry a heading above each table where a caption
    belongs. This finds paragraphs on -FromStyle that are immediately followed
    by a table, moves them onto -Style (Caption_Table), and prepends the NIST
    caption number, keeping the existing title text and its formatting.

    -FromStyle takes more than one style, which converted HTML usually needs:
    an <h2> above one table and an <h3> above the next both land on their own
    heading style, and an HTML <caption> inside a table comes through pandoc on
    Caption. Numbering runs in document order across the whole set, so mixing
    the levels does not disturb the sequence.

    The result matches the production NCSTAR chapters exactly:

        Table 4-1.  Income Data Used for Weighting

    which is, character by character: the word Table, a non-breaking space, the
    chapter number, an en dash, a SEQ field, a period, and two spaces.

    Only the table number is a field. The chapter number is written as literal
    text from -Chapter, because a converted extract has no Heading 1 for a
    STYLEREF field to read. Pass -UseStyleRef once the content is merged into a
    real chapter, and the chapter number tracks the heading it lands under.

    The SEQ field is written with its computed value cached between the
    separate and end marks, so the numbers read correctly before Word has
    updated fields. They renumber on Ctrl+A then F9, or on print.

    A paragraph counts as a caption when the next body-level element is a
    table, ignoring empty paragraphs in between. A heading that introduces
    prose rather than a table is left alone, so section headings survive.

    Captions go ABOVE the table, which is the NIST convention for tables and
    the opposite of figures.

    Idempotent in practice: once converted, the paragraphs are no longer on
    -FromStyle, so a second run finds nothing.

    Backs up to <name>.captions.bak before writing.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Folder of Word files. All .docx/.dotx/.docm/.dotm in it are processed.

.PARAMETER FromStyle
    StyleIds of the paragraphs to convert. Defaults to Heading1. Takes a list:
    -FromStyle Heading1,Heading2,Caption . A paragraph on any of them that sits
    directly above a table becomes a caption.

.PARAMETER Style
    StyleId of the caption style to apply. Defaults to CaptionTable, whose
    style name in the NIST template is Caption_Table.

.PARAMETER Chapter
    Chapter number written before the dash. Defaults to 1. Digits, or a single
    letter for an appendix. Ignored when -UseStyleRef is passed.

.PARAMETER UseStyleRef
    Emit a STYLEREF 1 \s field for the chapter number instead of literal text,
    matching the production chapters. Only useful once the content sits under a
    real Heading 1.

.PARAMETER Label
    The word before the number. Defaults to Table.

.PARAMETER SeqName
    The SEQ sequence name. Defaults to Table. Keep it aligned with the rest of
    the document so numbering is continuous.

.PARAMETER StripLeadingNumber
    Drop the number the source already put at the front of the title, so the
    caption does not read "Table 5-1.  1 Sample Characteristics". Converted
    HTML normally carries one, because the generator numbered its own tables.

    Only stripped when the leading number equals the number this caption is
    being given. A title like "2020 Enrollment Data" landing on table 3 keeps
    its 2020, because 2020 is not 3. If the source numbering is offset from the
    caption sequence nothing is stripped, which shows up as a visible doubled
    number rather than as silently deleted text.

.PARAMETER ListOnly
    Report which paragraphs would become captions and write nothing. Worth a
    first pass on any document whose headings are a real section hierarchy: a
    section heading that happens to sit above a table matches the same rule as
    a caption, and only the title text tells them apart.

.PARAMETER Force
    Skip the confirmation prompt.

.PARAMETER NoBackup
    Skip the .bak copy and write the file in place.

.EXAMPLE
    # Always look before writing.
    .\Set-NISTTableCaptions.ps1 -File "descriptivetabs.docx" -ListOnly

.EXAMPLE
    .\Set-NISTTableCaptions.ps1 -File "descriptivetabs.docx" -Chapter 1

.EXAMPLE
    # Converted HTML, where the headers above the tables are not all one level
    # and pandoc put any <caption> elements on Caption.
    .\Set-NISTTableCaptions.ps1 -File "tables.docx" -FromStyle Heading1,Heading2,Caption -Chapter 5

.EXAMPLE
    # Once merged into a real chapter, let the chapter number track the heading.
    .\Set-NISTTableCaptions.ps1 -File "Chapter-05.docx" -UseStyleRef
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string[]]$FromStyle = @('Heading1'),

    [string]$Style = 'CaptionTable',

    # Digits, or a single letter for an appendix (Table A-1). Anything else is
    # a mistake that would otherwise be written into every caption in the file.
    [ValidatePattern('^([0-9]+|[A-Za-z])$')]
    [string]$Chapter = '1',

    [switch]$UseStyleRef,

    [string]$Label = 'Table',

    [string]$SeqName = 'Table',

    [switch]$StripLeadingNumber,

    # Drops a whole leading caption prefix ("Table 1-1.  ") whatever its
    # numbers, which is what a chapter carries after it has been renumbered.
    # -StripLeadingNumber only drops a number that already matches.
    # Figure captions sit BELOW the graphic, table captions ABOVE the table, so
    # the adjacency test has to flip. With this the same caption machinery does
    # figures: -Label Figure -SeqName Figure -AfterGraphic.
    [switch]$AfterGraphic,

    [switch]$StripSourceCaption,

    # Trust the caption STYLE instead of adjacency. The adjacency rules exist to
    # decide which HEADING becomes a caption; a paragraph already on
    # Caption_Figure is known to be one, and testing adjacency then only loses
    # captions - in NWIRP 2-1 Ch04 four figures are rendered as tables, so
    # -AfterGraphic walks back, meets a w:tbl and skips them.
    [switch]$InPlace,

    [switch]$ListOnly,

    [switch]$Force,

    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# -InPlace means the paragraph is ALREADY a caption, so the style to look on is
# the caption style itself, not a heading. Defaulting -FromStyle to Heading1 and
# requiring the caller to remember -FromStyle CaptionFigure as well made the
# documented step 3e2 invocation a silent no-op: no Heading1 paragraph begins
# with "Figure", so it matched nothing, printed "Nothing to convert" for all 8
# chapters, and exited 0. Only override a default the caller did not set.
if ($InPlace -and -not $PSBoundParameters.ContainsKey('FromStyle')) {
    $FromStyle = @($Style)
}

# One alternation over every -FromStyle, so a mixed set of heading levels is
# found in a single document-order pass and the numbering stays sequential.
$fromAlt      = ($FromStyle | ForEach-Object { [regex]::Escape($_) }) -join '|'
$fromStyleRx  = '<w:pStyle\s[^>]*w:val="(?:' + $fromAlt + ')"[^>]*/>'
$fromStyleTxt = $FromStyle -join ', '

$NBSP   = [char]0x00A0
$ENDASH = [char]0x2013

# -------- Target list --------
$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|dotx|docm|dotm)$' -and $_.Name -notlike '~$*' } |
        Sort-Object Name |
        ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }

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


# -------- Confirm --------
$sample = if ($UseStyleRef) { "$Label$NBSP{STYLEREF 1 \s}$ENDASH{SEQ $SeqName}.  <title>" }
          else              { "$Label$NBSP$Chapter$ENDASH{SEQ $SeqName}.  <title>" }

Write-Host "Set NIST Table Captions" -ForegroundColor Cyan
if ($InPlace) {
    Write-Host "  Renumbers paragraphs already on $fromStyleTxt in place. Adjacency is not consulted." -ForegroundColor Cyan
} elseif ($AfterGraphic) {
    Write-Host "  Converts $fromStyleTxt paragraphs that sit directly BELOW a graphic into $Style." -ForegroundColor Cyan
} else {
    Write-Host "  Converts $fromStyleTxt paragraphs that sit directly above a table into $Style." -ForegroundColor Cyan
}
Write-Host "  Caption form: $sample" -ForegroundColor Cyan
Write-Host "  Headings not followed by a table are left alone." -ForegroundColor Cyan
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""

if ($ListOnly) { Write-Host "  -ListOnly: reporting matches, writing nothing." -ForegroundColor Yellow }
Write-Host ""

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# Boundary patterns. The character after the element name must be whitespace or
# the closing bracket, so <w:pPr>, <w:tblPr>, and <w:tblGrid> are never matched.
$elemPattern = '(?s)<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>|</w:p>|<w:tbl(?:\s[^>]*)?>|</w:tbl>'

function Get-DocumentChapterNumber {
    # The chapter number belongs to the DOCUMENT: HMList's ilvl 0 w:start IS the
    # number the headings render, which is where Set-NISTChapterNumbering and
    # Test-NISTDocument both read it. Returns $null when the document has none.
    #
    # This matters for the CACHED result of the STYLEREF field. -Chapter defaults
    # to 1, so a -UseStyleRef caption in chapter 5 was written with "1" as its
    # cached text and read "Figure 1-3" until someone opened it in Word and let
    # the field recompute. That is not cosmetic: it is exactly the stale number
    # an author then "corrects" by typing over it, and Set-TableCrossReferences
    # matches a prose mention against the number the caption currently SHOWS, so
    # a stale cache silently cost it every match in two chapters.
    param([string]$StylesXml, [string]$NumberingXml)
    if (-not $NumberingXml) { return $null }
    $abs = $null
    $sm = [regex]::Match($StylesXml, '(?s)<w:style\b[^>]*\bw:styleId="HMList"[^>]*>.*?</w:style>')
    if ($sm.Success) {
        $nid = [regex]::Match($sm.Value, '<w:numId\b[^>]*\bw:val="(\d+)"')
        if ($nid.Success) {
            $nm = [regex]::Match($NumberingXml, '(?s)<w:num\b[^>]*\bw:numId="' + $nid.Groups[1].Value + '"[^>]*>.*?</w:num>')
            if ($nm.Success) {
                $am = [regex]::Match($nm.Value, '<w:abstractNumId\b[^>]*\bw:val="(\d+)"')
                if ($am.Success) { $abs = $am.Groups[1].Value }
            }
        }
    }
    if (-not $abs) {
        # Fall back to the styleLink marker, which is the other end of the chain.
        $slp = [regex]::Match($NumberingXml, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"[^>]*>(?:(?!</w:abstractNum>).)*?<w:styleLink\b[^>]*\bw:val="HMList"')
        if ($slp.Success) { $abs = $slp.Groups[1].Value }
    }
    if (-not $abs) { return $null }
    $ab = [regex]::Match($NumberingXml, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="' + $abs + '"[^>]*>.*?</w:abstractNum>')
    if (-not $ab.Success) { return $null }
    $lv = [regex]::Match($ab.Value, '(?s)<w:lvl\b[^>]*\bw:ilvl="0"[^>]*>.*?</w:lvl>')
    if (-not $lv.Success) { return $null }
    $st = [regex]::Match($lv.Value, '<w:start\b[^>]*\bw:val="(\d+)"')
    if (-not $st.Success) { return $null }
    return $st.Groups[1].Value
}

function Get-CaptionRuns {
    param([int]$Number)

    $sb = New-Object System.Text.StringBuilder

    # "Table" + non-breaking space, so the label never wraps off its number.
    [void]$sb.Append("<w:r><w:t xml:space=`"preserve`">$Label$NBSP</w:t></w:r>")

    if ($UseStyleRef) {
        [void]$sb.Append('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
        # Byte-for-byte what the production chapters carry: no xml:space, no
        # trailing space. \s returns the %1 of the heading's number, so a
        # Heading 1 numbered "Chapter 4" yields "4", not "Chapter 4".
        [void]$sb.Append('<w:r><w:instrText>STYLEREF 1 \s</w:instrText></w:r>')
        [void]$sb.Append('<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
        $cached = if ($script:EffectiveChapter) { $script:EffectiveChapter } else { $Chapter }
        [void]$sb.Append("<w:r><w:rPr><w:noProof/></w:rPr><w:t>$cached</w:t></w:r>")
        [void]$sb.Append('<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    } else {
        [void]$sb.Append("<w:r><w:t>$Chapter</w:t></w:r>")
    }

    [void]$sb.Append("<w:r><w:t>$ENDASH</w:t></w:r>")

    # The table number itself, as a field so it renumbers on F9.
    [void]$sb.Append('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
    [void]$sb.Append("<w:r><w:instrText xml:space=`"preserve`"> SEQ $SeqName \* ARABIC \s 1 </w:instrText></w:r>")
    [void]$sb.Append('<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
    [void]$sb.Append("<w:r><w:rPr><w:noProof/></w:rPr><w:t>$Number</w:t></w:r>")
    [void]$sb.Append('<w:r><w:fldChar w:fldCharType="end"/></w:r>')

    # Period and the two spaces the production chapters use before the title.
    [void]$sb.Append('<w:r><w:t xml:space="preserve">.  </w:t></w:r>')

    return $sb.ToString()
}

function Remove-SourceCaptionUnits {
    # Drop a leading "<Label> <chapter><dash><number>." prefix, taking whole runs
    # and whole fields with it.
    #
    # Remove-SourceCaption clears the TEXT of the prefix. That is not enough when
    # the chapter number is already a STYLEREF field: emptying its cached result
    # leaves the field itself in place, and the next F9 renders the number again,
    # in front of the new one. A field has to be removed outright, so the walk is
    # over units - a plain run, or a whole begin-to-end field - rather than runs.
    param([string]$Para, [string]$Label)

    # A number can be carried three ways: plain runs, a begin-to-end field, or a
    # one-element <w:fldSimple>. All three have to come off, or the leftover
    # field renders its own number again at the next F9.
    $runs = [regex]::Matches($Para, '(?s)<w:fldSimple(?:\s[^>]*)?>.*?</w:fldSimple>|<w:r(?:\s[^>]*)?>.*?</w:r>')
    if ($runs.Count -eq 0) { return $Para }

    # Group the runs into units, collapsing each field into one.
    $units = New-Object System.Collections.ArrayList
    $depth = 0; $fieldStart = -1; $fieldEnd = -1; $fieldText = ''
    foreach ($r in $runs) {
        $v = $r.Value
        # w:t is NOT the only element that puts characters on the page, and this
        # function was reading only w:t. A noBreakHyphen run therefore measured
        # zero-length and was swept up by the "$left -ge $len" branch below for
        # free, which is the accident that made the walk appear to work - and the
        # same accident consumed any OTHER text-free run, including a tracked
        # DELETION. That is how Ch02 lost a w:delText digit out of an author's
        # revision. Count the hyphen as the character it is, keep delText and
        # instrText out (a deleted or instruction character is not on the page),
        # and the lengths then describe what a reader actually sees.
        #
        # The tab form comes FIRST: '<w:tab/>'.StartsWith('<w:t') is true, so an
        # unordered alternation lets the w:t branch try it first.
        $txt = -join ([regex]::Matches($v,
                '(?s)<w:tab\s*/>|<w:noBreakHyphen\s*/>|<w:softHyphen\s*/>|<w:t(?:\s[^>]*)?>(.*?)</w:t>') |
                ForEach-Object {
                    if     ($_.Value -like '<w:tab*')           { "`t" }
                    elseif ($_.Value -like '<w:noBreakHyphen*') { [char]0x2011 }
                    elseif ($_.Value -like '<w:softHyphen*')    { [char]0x00AD }
                    else                                        { $_.Groups[1].Value }
                })
        if ($v -like '<w:fldSimple*' -and $depth -eq 0) {
            [void]$units.Add(@{ Start=$r.Index; End=$r.Index + $r.Length; Text=$txt; Field=$true })
            continue
        }
        if ($v -match 'w:fldCharType="begin"') {
            if ($depth -eq 0) { $fieldStart = $r.Index; $fieldText = '' }
            $depth++
            $fieldEnd = $r.Index + $r.Length
            continue
        }
        if ($depth -gt 0) {
            $fieldEnd = $r.Index + $r.Length
            if ($v -match 'w:fldCharType="end"') {
                $depth--
                if ($depth -eq 0) {
                    [void]$units.Add(@{ Start=$fieldStart; End=$fieldEnd; Text=$fieldText; Field=$true })
                }
            } elseif ($v -notmatch '<w:instrText') {
                $fieldText += $txt      # the cached result between separate and end
            }
            continue
        }
        [void]$units.Add(@{ Start=$r.Index; End=$r.Index + $r.Length; Text=$txt; Field=$false })
    }
    if ($depth -ne 0) { return $Para }      # unbalanced field, leave it alone

    $joined = -join ($units | ForEach-Object { $_.Text })
    # The chapter part is OPTIONAL and its separator may be a dot. Both forms
    # are in the NWIRP 2-1 set: Ch05 and Ch06 carry a typed "Figure 5.1." with a
    # dot, and Ch04 carries "Figure 1." with no chapter at all because the chapter
    # was never numbered. The old pattern required digit-dash-digit, so it matched
    # neither and the stale number survived in front of the new one.
    # Two traps in this one pattern, both found by running it over a real set.
    #
    # No whitespace after the chapter separator. Allowing it let the chapter
    # group swallow "5.  " out of "Figure 5.  2020 Census", leaving \d+ to match
    # the YEAR as the figure number and strip it: the caption came out as
    # "Figure 5-1.  Census". Every real chapter number is written tight against
    # its separator ("5-1", "5.1", "B.1"), and the whitespace belongs to the
    # period that ENDS the number, so the group has no business consuming it.
    #
    # The number may be a PLACEHOLDER, not a digit: a caption with no number yet
    # is written "Figure X-X." and one sits in Ch02. The lookahead keeps that
    # narrow - without it [Xx] also matches the first letter of a title like
    # "Xavier's method" and strips "Figure X", leaving "avier's method".
    $rx = '^\s*' + [regex]::Escape($Label) +
          '[\s\u00A0]*(?:[A-Za-z0-9]+[\s\u00A0]*[.\-\u2010\u2011\u2013\u2014])?' +
          '(?:\d+|[Xx](?![A-Za-z]))\s*[.):]?(?:\s|\u00A0)*'
    $lead = [regex]::Match($joined, $rx)
    if (-not $lead.Success) { return $Para }

    $left = $lead.Length
    $edits = New-Object System.Collections.ArrayList
    $partial = $false
    $consumedTo = -1
    $i = -1
    foreach ($u in $units) {
        $i++
        if ($left -le 0) { break }
        $consumedTo = $i
        $len = $u.Text.Length
        # A unit that contributes NO text is never part of the number, whatever
        # is in it. "$left -ge $len" is trivially true when $len is 0, so such a
        # unit was deleted as part of the prefix while $left did not even
        # advance. Two things were lost that way and neither moved a count that
        # anyone was checking: a DRAWING, which is how two figures left Ch05,
        # and a tracked DELETION, which is how Ch02 lost a w:delText digit out
        # of an author's revision history. Listing the element types that may be
        # stepped over was the wrong shape for the rule - the rule is about
        # contributing no text. Step over all of them.
        if ($len -eq 0) { continue }
        if ($left -ge $len) {
            # The unit goes, but a bookmark inside it stays: it is the anchor a
            # cross-reference points at, and dropping it breaks every REF field
            # aimed here. Re-emitted in place, so the new number lands inside it.
            $keepMarks = -join ([regex]::Matches($Para.Substring($u.Start, $u.End - $u.Start),
                                                 '<w:bookmark(?:Start|End)\s[^>]*/>') | ForEach-Object { $_.Value })
            [void]$edits.Add(@{ Start=$u.Start; End=$u.End; Keep=$keepMarks })
            $left -= $len
        } else {
            if ($u.Field) { break }   # never cut a field in half
            # The rebuild below re-emits what is left as a plain w:t, which would
            # turn a noBreakHyphen or a tab into a literal character. Only split
            # a unit whose text is all w:t; otherwise stop and leave the rest.
            if ($Para.Substring($u.Start, $u.End - $u.Start) -match '<w:tab\s*/>|<w:noBreakHyphen\s*/>|<w:softHyphen\s*/>') { break }
            $keepText = $u.Text.Substring($left)
            $open = [regex]::Match($Para.Substring($u.Start), '^<w:r(?:\s[^>]*)?>').Value
            $rPrM = [regex]::Match($Para.Substring($u.Start, $u.End - $u.Start), '(?s)^<w:r(?:\s[^>]*)?>\s*(<w:rPr>.*?</w:rPr>)')
            $rPr = if ($rPrM.Success) { $rPrM.Groups[1].Value } else { '' }
            [void]$edits.Add(@{ Start=$u.Start; End=$u.End; Keep=($open + $rPr + '<w:t xml:space="preserve">' + $keepText + '</w:t></w:r>') })
            $left = 0
            $partial = $true
        }
    }

    # The number region can CONTINUE past the counted prefix when an author
    # typed part of a number in front of the field that already produced it.
    # Ch03 reads "Figure " + Sean's typed "3-2" + STYLEREF(3) + noBreakHyphen +
    # SEQ(3) + ".  ", which joins to "Figure 3-23-3.  ": the greedy digit run
    # matched "3-23" and stopped INSIDE the field pair, leaving "-3.  " sitting
    # in front of the new number as "Figure 3-2.  -3.  Visual Aid".
    #
    # Only extend when the counted prefix did NOT already end in a terminator.
    # That is the discriminator which keeps "Figure 5.  2020 Census" safe: its
    # match ends at ".  ", so the year is the title and is never touched. Here
    # the match ended on a digit with a separator still ahead of it, which only
    # happens mid-number.
    if (-not $partial -and $left -le 0 -and $lead.Value -notmatch '[.):][\s\u00A0]*$') {
        for ($j = $consumedTo + 1; $j -lt $units.Count; $j++) {
            $u = $units[$j]
            $t = $u.Text
            if ($t.Length -eq 0) { continue }
            $keepMarks = -join ([regex]::Matches($Para.Substring($u.Start, $u.End - $u.Start),
                                                 '<w:bookmark(?:Start|End)\s[^>]*/>') | ForEach-Object { $_.Value })
            if ($t -match '^[\s\u00A0]*[.\-\u2010\u2011\u2013\u2014][\s\u00A0]*$') {
                [void]$edits.Add(@{ Start=$u.Start; End=$u.End; Keep=$keepMarks })   # a separator
                continue
            }
            if ($t -match '^[\s\u00A0]*\d+[\s\u00A0]*$') {
                [void]$edits.Add(@{ Start=$u.Start; End=$u.End; Keep=$keepMarks })   # more of the number
                continue
            }
            # U+FFFC is the OBJECT REPLACEMENT glyph an instruction-less field
            # renders as, and Ch04 has one sitting between its number and the
            # terminator: "Figure 4-3<FFFC>.  FCC cell site ...". It is not a
            # character the author typed and it is not a number, so it blocked
            # the terminator from ever being reached and left ".  " stranded in
            # front of the title. A field with no instruction is an error
            # (Repair-WordFields removes it); consuming it here is the same
            # verdict reached earlier in the pipeline.
            if ($t -match '^[\s\u00A0\uFFFC\u200B]+$') {
                [void]$edits.Add(@{ Start=$u.Start; End=$u.End; Keep=$keepMarks })
                continue
            }
            # A terminator ends it. Take the terminator and its spaces only, and
            # leave the title alone.
            $term = [regex]::Match($t, '^[\s\u00A0]*[.):][\s\u00A0]*')
            if ($term.Success -and -not $u.Field) {
                if ($term.Length -ge $t.Length) {
                    [void]$edits.Add(@{ Start=$u.Start; End=$u.End; Keep=$keepMarks })
                } else {
                    $open = [regex]::Match($Para.Substring($u.Start), '^<w:r(?:\s[^>]*)?>').Value
                    $rPrM = [regex]::Match($Para.Substring($u.Start, $u.End - $u.Start), '(?s)^<w:r(?:\s[^>]*)?>\s*(<w:rPr>.*?</w:rPr>)')
                    $rPr  = if ($rPrM.Success) { $rPrM.Groups[1].Value } else { '' }
                    [void]$edits.Add(@{ Start=$u.Start; End=$u.End
                                        Keep=($keepMarks + $open + $rPr + '<w:t xml:space="preserve">' + $t.Substring($term.Length) + '</w:t></w:r>') })
                }
            }
            break
        }
    }

    $sb = New-Object System.Text.StringBuilder $Para
    for ($i = $edits.Count - 1; $i -ge 0; $i--) {
        $e = $edits[$i]
        [void]$sb.Remove($e.Start, $e.End - $e.Start)
        if ($e.Keep) { [void]$sb.Insert($e.Start, $e.Keep) }
    }
    return $sb.ToString()
}

function Remove-SourceCaption {
    # Drop a leading "<Label> <chapter><dash><number>." prefix and its trailing
    # whitespace, whatever the numbers are. Same run-by-run consumption as
    # Remove-SourceNumber: the prefix is measured against the joined text, then
    # taken off run by run, back to front so the offsets stay valid.
    param([string]$Para, [string]$Label)

    $ts = [regex]::Matches($Para, '(?s)(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)')
    if ($ts.Count -eq 0) { return $Para }

    $joined = -join ($ts | ForEach-Object { $_.Groups[2].Value })
    # The chapter part is OPTIONAL and its separator may be a dot. Both forms
    # are in the NWIRP 2-1 set: Ch05 and Ch06 carry a typed "Figure 5.1." with a
    # dot, and Ch04 carries "Figure 1." with no chapter at all because the chapter
    # was never numbered. The old pattern required digit-dash-digit, so it matched
    # neither and the stale number survived in front of the new one.
    # Two traps in this one pattern, both found by running it over a real set.
    #
    # No whitespace after the chapter separator. Allowing it let the chapter
    # group swallow "5.  " out of "Figure 5.  2020 Census", leaving \d+ to match
    # the YEAR as the figure number and strip it: the caption came out as
    # "Figure 5-1.  Census". Every real chapter number is written tight against
    # its separator ("5-1", "5.1", "B.1"), and the whitespace belongs to the
    # period that ENDS the number, so the group has no business consuming it.
    #
    # The number may be a PLACEHOLDER, not a digit: a caption with no number yet
    # is written "Figure X-X." and one sits in Ch02. The lookahead keeps that
    # narrow - without it [Xx] also matches the first letter of a title like
    # "Xavier's method" and strips "Figure X", leaving "avier's method".
    $rx = '^\s*' + [regex]::Escape($Label) +
          '[\s\u00A0]*(?:[A-Za-z0-9]+[\s\u00A0]*[.\-\u2010\u2011\u2013\u2014])?' +
          '(?:\d+|[Xx](?![A-Za-z]))\s*[.):]?(?:\s|\u00A0)*'
    $lead = [regex]::Match($joined, $rx)
    if (-not $lead.Success) { return $Para }

    $newTexts = New-Object 'System.Collections.Generic.List[string]'
    $left = $lead.Length
    foreach ($t in $ts) {
        $txt = $t.Groups[2].Value
        if ($left -le 0) { [void]$newTexts.Add($txt); continue }
        if ($left -ge $txt.Length) { $left -= $txt.Length; [void]$newTexts.Add('') }
        else { [void]$newTexts.Add($txt.Substring($left)); $left = 0 }
    }

    $sb = New-Object System.Text.StringBuilder $Para
    for ($i = $ts.Count - 1; $i -ge 0; $i--) {
        $t = $ts[$i]
        [void]$sb.Remove($t.Index, $t.Length)
        [void]$sb.Insert($t.Index, ($t.Groups[1].Value + $newTexts[$i] + $t.Groups[3].Value))
    }
    return $sb.ToString()
}

function Remove-SourceNumber {
    # Drop a leading "<n>", "<n>.", or "<n>)" plus its whitespace from the
    # title, but only when <n> is the number this caption is being given.
    #
    # The number is rarely in one run. A converter splits "1 Sample
    # Characteristics" into three: "1", " ", and the title. So the text of every
    # w:t is joined, the prefix is measured against the whole, and that many
    # characters are then consumed run by run.
    param([string]$Para, [int]$Number)

    $ts = [regex]::Matches($Para, '(?s)(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)')
    if ($ts.Count -eq 0) { return $Para }

    $joined = -join ($ts | ForEach-Object { $_.Groups[2].Value })
    $lead = [regex]::Match($joined, '^\s*(\d+)\s*[.)]?(?:\s|\u00A0)+')
    if (-not $lead.Success) { return $Para }
    if ([int]$lead.Groups[1].Value -ne $Number) { return $Para }

    # Characters are counted front to back, so each run's new text is worked out
    # first; the paragraph is then rebuilt back to front so offsets stay valid.
    $newTexts = New-Object 'System.Collections.Generic.List[string]'
    $left = $lead.Length
    foreach ($t in $ts) {
        $txt = $t.Groups[2].Value
        if ($left -le 0) { [void]$newTexts.Add($txt); continue }
        if ($left -ge $txt.Length) { $left -= $txt.Length; [void]$newTexts.Add('') }
        else { [void]$newTexts.Add($txt.Substring($left)); $left = 0 }
    }

    $sb = New-Object System.Text.StringBuilder $Para
    for ($i = $ts.Count - 1; $i -ge 0; $i--) {
        $t = $ts[$i]
        $replacement = $t.Groups[1].Value + $newTexts[$i] + $t.Groups[3].Value
        [void]$sb.Remove($t.Index, $t.Length)
        [void]$sb.Insert($t.Index, $replacement)
    }
    return $sb.ToString()
}

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $entries   = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $srcStream = [System.IO.File]::OpenRead($target)
    try {
        $srcZip = New-Object System.IO.Compression.ZipArchive($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $srcZip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
            }
        } finally { $srcZip.Dispose() }
    } finally { $srcStream.Dispose() }

    if (-not $entries.ContainsKey('word/document.xml')) { throw "No word/document.xml in $target" }
    if (-not $entries.ContainsKey('word/styles.xml'))   { throw "No word/styles.xml in $target" }

    # A reference to an undefined style renders as Normal, with no warning.
    $stylesXml = $utf8NoBom.GetString($entries['word/styles.xml'])
    if (-not [regex]::IsMatch($stylesXml, '<w:style\b[^>]*\bw:styleId="' + [regex]::Escape($Style) + '"')) {
        throw "Style '$Style' is not defined in $target. Apply the template's styles first (Set-NISTStyleArchitecture.ps1)."
    }

    $xml = $utf8NoBom.GetString($entries['word/document.xml'])

    # The cached STYLEREF result, read from this document rather than defaulted.
    $script:EffectiveChapter = $null
    if ($UseStyleRef) {
        $numXmlForChapter = if ($entries.ContainsKey('word/numbering.xml')) { $utf8NoBom.GetString($entries['word/numbering.xml']) } else { '' }
        $script:EffectiveChapter = Get-DocumentChapterNumber -StylesXml $stylesXml -NumberingXml $numXmlForChapter
        if ($PSBoundParameters.ContainsKey('Chapter')) {
            $script:EffectiveChapter = $Chapter
            Write-Host "  Cached chapter number $Chapter (given)." -ForegroundColor DarkGray
        } elseif ($script:EffectiveChapter) {
            Write-Host "  Cached chapter number $($script:EffectiveChapter) read from the document's own HMList." -ForegroundColor DarkGray
        } else {
            Write-Warning "  This document carries no chapter number; the STYLEREF cache will read '$Chapter' until Word recomputes it."
        }
    }

    # Work inside <w:body> only, so a sectPr or docPart elsewhere cannot confuse
    # the top-level scan.
    $bodyMatch = [regex]::Match($xml, '(?s)<w:body(?:\s[^>]*)?>.*</w:body>')
    if (-not $bodyMatch.Success) { throw "No <w:body> in $target" }
    $bodyStart = $bodyMatch.Index
    $body      = $bodyMatch.Value

    # Walk the body collecting top-level paragraphs and tables in document
    # order. Depth tracking keeps a table's own rows and any nested table out of
    # the list, so adjacency is measured between siblings.
    $elements = New-Object System.Collections.ArrayList
    $depth    = 0
    $openIdx  = -1
    $openKind = ''
    foreach ($m in [regex]::Matches($body, $elemPattern)) {
        $v = $m.Value
        if ($v -match '^<w:p(?:\s[^>]*)?/>$') {
            if ($depth -eq 0) {
                [void]$elements.Add(@{ Kind='p'; Start=$m.Index; End=$m.Index+$m.Length; Empty=$true; Gfx=$false })
            }
            continue
        }
        if ($v -eq '</w:p>' -or $v -eq '</w:tbl>') {
            $depth--
            if ($depth -eq 0 -and $openIdx -ge 0) {
                $end = $m.Index + $m.Length
                # A paragraph holding a drawing or a VML picture is the figure
                # itself, not a caption: -AfterGraphic anchors on it.
                $isGfx = $false
                if ($openKind -eq 'p') {
                    $isGfx = [regex]::IsMatch($body.Substring($openIdx, $end - $openIdx), '<w:(?:drawing|pict)\b')
                }
                [void]$elements.Add(@{ Kind=$openKind; Start=$openIdx; End=$end; Empty=$false; Gfx=$isGfx })
                $openIdx = -1
            }
            continue
        }
        if ($depth -eq 0) {
            $openIdx  = $m.Index
            $openKind = if ($v -like '<w:tbl*') { 'tbl' } else { 'p' }
        }
        $depth++
    }
    if ($depth -ne 0) { throw "Unbalanced paragraph or table tags in $target (depth $depth). File may be corrupt." }

    # Identify which top-level paragraphs are captions: on $FromStyle, with a
    # table as the next sibling once empty paragraphs are skipped.
    $captions = New-Object System.Collections.ArrayList
    $appendixSkipped = 0
    for ($i = 0; $i -lt $elements.Count; $i++) {
        $el = $elements[$i]
        if ($el.Kind -ne 'p' -or $el.Empty) { continue }
        $para = $body.Substring($el.Start, $el.End - $el.Start)
        if (-not [regex]::IsMatch($para, $fromStyleRx)) { continue }

        # An appendix caption ("Table E-4.") numbers in its own sequence, so it
        # is left exactly as it is: no SEQ field, no renumbering, no strip.
        $capText = -join ([regex]::Matches($para, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
        if ([regex]::IsMatch($capText, '^\s*' + [regex]::Escape($Label) + '[\s\u00A0]*[A-Za-z]\s*[.\-\u2010\u2011\u2013\u2014]\s*\d+')) {
            $appendixSkipped++
            continue
        }

        if ($InPlace) {
            # Already on the caption style, so it IS a caption. Only require the
            # text to begin with the label, so a Table caption is not pulled into
            # the Figure sequence when both share a run.
            if ([regex]::IsMatch($capText, '^\s*' + [regex]::Escape($Label) + '\b')) {
                [void]$captions.Add($i)
            }
            continue
        }

        if ($AfterGraphic) {
            # The caption is the first text paragraph after the graphic, so walk
            # BACKWARD over text-free paragraphs to the nearest sibling and
            # require it to be the graphic. A paragraph that holds the graphic
            # is never its own caption.
            if ($el.Gfx) { continue }
            $j = $i - 1
            while ($j -ge 0 -and $elements[$j].Kind -eq 'p' -and -not $elements[$j].Gfx) {
                $prevPara = $body.Substring($elements[$j].Start, $elements[$j].End - $elements[$j].Start)
                $hasText  = [regex]::Match($prevPara, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')
                if ($hasText.Success -and $hasText.Groups[1].Value.Trim().Length -gt 0) { break }
                $j--
            }
            if ($j -ge 0 -and $elements[$j].Kind -eq 'p' -and $elements[$j].Gfx) { [void]$captions.Add($i) }
        } else {
            $j = $i + 1
            while ($j -lt $elements.Count -and $elements[$j].Kind -eq 'p') {
                $nextPara = $body.Substring($elements[$j].Start, $elements[$j].End - $elements[$j].Start)
                $hasText  = [regex]::Match($nextPara, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')
                if ($hasText.Success -and $hasText.Groups[1].Value.Trim().Length -gt 0) { break }
                $j++
            }
            if ($j -lt $elements.Count -and $elements[$j].Kind -eq 'tbl') { [void]$captions.Add($i) }
        }
    }

    if ($appendixSkipped -gt 0) {
        Write-Host "  Appendix captions left as text: $appendixSkipped" -ForegroundColor DarkGray
    }
    $tableCount = ($elements | Where-Object { $_.Kind -eq 'tbl' }).Count
    Write-Host "  Top-level tables: $tableCount" -ForegroundColor Gray
    $adjTxt = if ($InPlace) { 'renumbered in place' }
              elseif ($AfterGraphic) { 'directly below a graphic' }
              else { 'directly above a table' }
    Write-Host "  $fromStyleTxt paragraphs ${adjTxt}: $($captions.Count)" -ForegroundColor Gray

    # Print the titles. A section heading that happens to sit above a table
    # matches the same rule as a caption; the text is the only thing that tells
    # them apart, so it goes on screen either way.
    $seq = 0
    foreach ($idx in $captions) {
        $seq++
        $para = $body.Substring($elements[$idx].Start, $elements[$idx].End - $elements[$idx].Start)
        # Preview the title the WRITE path will produce, which means applying the
        # same strip. Printing the raw text instead showed the new prefix in front
        # of the whole old caption - "Figure #-1.  Figure 2-1  High-capacity ..." -
        # so a correct run looked like it was about to double every number, and a
        # genuinely doubled result would have looked exactly as normal. -ListOnly
        # is the instrument for deciding whether to apply; it has to show the
        # outcome, not the input.
        if ($StripSourceCaption) { $para = Remove-SourceCaptionUnits -Para $para -Label $Label }
        $title = -join ([regex]::Matches($para, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
        if ($title.Length -gt 60) { $title = $title.Substring(0, 57) + '...' }
        Write-Host ("    $Label$NBSP$(if($UseStyleRef){'#'}else{$Chapter})$ENDASH{0}.  {1}" -f $seq, $title) -ForegroundColor DarkGray
    }

    if ($captions.Count -eq 0 -or $ListOnly) {
        if ($captions.Count -eq 0) { Write-Host "  Nothing to convert. File unchanged." -ForegroundColor DarkGray }
        else { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow }
        continue
    }

    # Rewrite back to front so recorded offsets stay valid. Numbering is assigned
    # front to back first.
    $numbers = @{}
    for ($n = 0; $n -lt $captions.Count; $n++) { $numbers[$captions[$n]] = $n + 1 }
    $stripped = 0

    $sb = New-Object System.Text.StringBuilder $body
    for ($n = $captions.Count - 1; $n -ge 0; $n--) {
        $idx  = $captions[$n]
        $el   = $elements[$idx]
        $para = $body.Substring($el.Start, $el.End - $el.Start)

        if ($StripSourceCaption) {
            $before = $para
            $para = Remove-SourceCaptionUnits -Para $para -Label $Label
            if ($para -ne $before) { $stripped++ }
        }

        if ($StripLeadingNumber) {
            $before = $para
            $para = Remove-SourceNumber -Para $para -Number $numbers[$idx]
            if ($para -ne $before) { $stripped++ }
        }

        # Point the paragraph at the caption style.
        $newPara = [regex]::Replace(
            $para,
            $fromStyleRx,
            "<w:pStyle w:val=`"$Style`"/>",
            1)

        # Insert the numbering runs after <w:pPr>, ahead of the title runs, so
        # the existing text and its formatting are untouched.
        $pPrMatch = [regex]::Match($newPara, '(?s)<w:pPr(?:\s[^>]*)?>.*?</w:pPr>')
        if (-not $pPrMatch.Success) { throw "Caption paragraph lost its pPr in $target" }
        $insertAt = $pPrMatch.Index + $pPrMatch.Length

        # Step over any bookmark that opens right after w:pPr. A cross-reference
        # to this caption (Set-TableCrossReferences.ps1) brackets the label and
        # number, so the new runs have to land INSIDE that bookmark or every
        # REF field pointing here comes back empty on the next F9.
        # Only a CROSS-REFERENCE anchor counts, and those are named _Ref*. An
        # _Hlk* bookmark is Word's inert "last edit location" marker - nothing
        # REFs one - and Ch05's caption opens with _Hlk229236279 whose end sits
        # at the far end of the paragraph. Treating it as an anchor moved the
        # period and its two spaces down there with it: "Table 5-1Potential
        # explanatory variables ... transportation model..  ". The number lost
        # its terminator and the caption gained a double period at the end.
        # Step over anything that must stay AHEAD of the number: a bookmark, and
        # - the case that matters here - the FIGURE itself.
        #
        # A caption paragraph may hold its own image, with the drawing run first
        # and the caption text after it, which is how several of Ch05's figures
        # are built. Inserting the number straight after w:pPr put it BEFORE that
        # drawing, so the number rendered above the picture and the title below
        # it: "Figure 5-4." on one line, the image, then "Maps comparing
        # predicted and observed disruptions ...". That is the split caption the
        # authors reported, re-created by the very pass meant to fix captions.
        $bmIds = New-Object System.Collections.ArrayList
        while ($true) {
            $rest = $newPara.Substring($insertAt)
            $bm = [regex]::Match($rest, '^<w:bookmarkStart\s[^>]*/>')
            if ($bm.Success) {
                $nameM = [regex]::Match($bm.Value, 'w:name="([^"]*)"')
                $idM   = [regex]::Match($bm.Value, 'w:id="(\d+)"')
                if ($idM.Success -and $nameM.Success -and $nameM.Groups[1].Value -like '_Ref*') {
                    [void]$bmIds.Add($idM.Groups[1].Value)
                }
                $insertAt += $bm.Length
                continue
            }
            $be = [regex]::Match($rest, '^<w:bookmarkEnd\s[^>]*/>')
            if ($be.Success) { $insertAt += $be.Length; continue }
            # A leading run that carries the picture and no text of its own.
            $rn = [regex]::Match($rest, '(?s)^<w:r(?:\s[^>]*)?>.*?</w:r>')
            if ($rn.Success -and $rn.Value -match '<w:drawing\b|<w:pict\b|<w:object\b') {
                $insertAt += $rn.Length
                continue
            }
            break
        }

        # The number goes inside the bookmark, the period and its two spaces
        # after it. A cross-reference reads "Table 4-1", never "Table 4-1.  ".
        $capRuns = Get-CaptionRuns -Number $numbers[$idx]
        $tailRun = '<w:r><w:t xml:space="preserve">.  </w:t></w:r>'
        $numRuns = $capRuns
        $hasTail = $false
        if ($capRuns.EndsWith($tailRun)) {
            $numRuns = $capRuns.Substring(0, $capRuns.Length - $tailRun.Length)
            $hasTail = $true
        }

        $newPara = $newPara.Insert($insertAt, $numRuns)

        if ($hasTail) {
            $tailAt = $insertAt + $numRuns.Length
            # Any of the bookmarks opened here may be the cross-reference anchor,
            # and some (a heading anchor, say) close further down the document
            # instead of in this paragraph. Take the last one that does close
            # here, so the period ends up outside every one of them.
            foreach ($bmId in $bmIds) {
                $closeRx = '<w:bookmarkEnd\s[^>]*w:id="' + [regex]::Escape($bmId) + '"[^>]*/>'
                $close = [regex]::Match($newPara, $closeRx)
                if ($close.Success -and $close.Index -ge $tailAt) {
                    $candidate = $close.Index + $close.Length
                    # The period may move past MARKUP only, never past text. If
                    # anything between here and the bookmark's end puts a
                    # character on the page, that bookmark spans the title too
                    # and is not the number's anchor.
                    $between = $newPara.Substring($tailAt, $candidate - $tailAt)
                    if ($between -match '<w:t[ >]|<w:noBreakHyphen|<w:tab\b|<w:delText[ >]') { continue }
                    if ($candidate -gt $tailAt) { $tailAt = $candidate }
                }
            }
            $newPara = $newPara.Insert($tailAt, $tailRun)
        }

        [void]$sb.Remove($el.Start, $el.End - $el.Start)
        [void]$sb.Insert($el.Start, $newPara)
    }

    $newBody = $sb.ToString()
    $newXml  = $xml.Remove($bodyStart, $body.Length).Insert($bodyStart, $newBody)

    if ($StripLeadingNumber -or $StripSourceCaption) { Write-Host "  Source numbers dropped from the title: $stripped" -ForegroundColor DarkGray }
    Write-Host "  Captions written: $($captions.Count) ($Label$NBSP$(if($UseStyleRef){'{STYLEREF}'}else{$Chapter})$ENDASH`1 through $($captions.Count))" -ForegroundColor Gray

    # --- guards -----------------------------------------------------------
    # This script had none, and it rewrites caption paragraphs by consuming
    # whole RUNS to cover the old number. A run holding a drawing contributes no
    # text, so "$left -ge $len" is true for it with $len = 0 and the picture is
    # swallowed as part of the prefix. That is how two figures disappeared from
    # NWIRP 2-1 Ch05 - both in the "caption paragraph that holds its own image"
    # shape, where the drawing run sits first and the caption text follows - and
    # nothing caught it: the package still held the media part and the
    # relationship, every XML part parsed, and the text was intact.
    #
    # The caption TEXT is expected to change here, so it cannot be compared.
    # Everything else must be untouched.
    try {
        $probe = New-Object System.Xml.XmlDocument
        $probe.XmlResolver = $null
        $probeXml = $newXml; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml)
    } catch {
        throw "Caption rewrite produced malformed document.xml for ${target}: $($_.Exception.Message). Nothing written."
    }
    foreach ($tag in @('<a:blip\b', '<w:drawing\b', '<w:pict\b', '<w:p\b', '<w:tbl>',
                       '<w:ins\b', '<w:del\b', '<w:delText\b', '<w:delInstrText\b',
                       '<w:commentReference\b', '<w:footnoteReference\b')) {
        $n = ([regex]::Matches($newXml, $tag)).Count
        $o = ([regex]::Matches($xml, $tag)).Count
        if ($n -ne $o) {
            throw ("Caption rewrite changed a protected element count in ${target}: $tag $o -> $n. " +
                   "Nothing written. A drawing or tracked change sharing a caption paragraph was consumed " +
                   "as part of the old number.")
        }
    }
    $strayClose = ([regex]::Matches($newXml, '/>\s*/>')).Count - ([regex]::Matches($xml, '/>\s*/>')).Count
    if ($strayClose -gt 0) { throw "Caption rewrite left $strayClose stray '/>' in ${target}. Nothing written." }

    if (-not $Force -and (Test-Path -LiteralPath "$target.captions.bak")) {
        Write-Warning "  Backup $target.captions.bak already exists; leaving it as-is."
    } else {
        if (-not $NoBackup) { Copy-Item -LiteralPath $target -Destination "$target.captions.bak" -Force }
    }

    $entries['word/document.xml'] = $utf8NoBom.GetBytes($newXml)

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "captions-$([guid]::NewGuid()).zip")
    $outStream = [System.IO.File]::Create($tempZip)
    try {
        $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $entries.Keys) {
                $ne = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try { $os.Write($entries[$name], 0, $entries[$name].Length) } finally { $os.Dispose() }
            }
        } finally { $outZip.Dispose() }
    } finally { $outStream.Dispose() }

    Move-Item -LiteralPath $tempZip -Destination $target -Force
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Select all and press F9 in Word to update the SEQ fields." -ForegroundColor Cyan
