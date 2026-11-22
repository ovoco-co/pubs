<#
.SYNOPSIS
    Restores the NIST chapter and section auto-numbering chain: the HMList
    multilevel list, its pStyle links to Heading1-4, and the numPr on those
    heading styles. Optionally removes a chapter number that was typed into
    the Heading1 text.

.DESCRIPTION
    NIST chapter numbering runs through a named multilevel list style, HMList,
    not through numbering applied to each heading. The full chain is:

        Heading1  -> numPr numId N  -> abstractNum A
                                       (A carries only <w:numStyleLink w:val="HMList"/>)
        HMList    -> numPr numId M  -> abstractNum B
                                       (B carries <w:styleLink w:val="HMList"/>
                                        and the nine w:lvl definitions)
        Heading2/3/4 -> numPr ilvl 1/2/3, numId N

    and inside B the first four levels name the style they number:

        ilvl 0  pStyle Heading1  decimal  "Chapter %1"
        ilvl 1  pStyle Heading2  decimal  "%1.%2."
        ilvl 2  pStyle Heading3  decimal  "%1.%2.%3."
        ilvl 3  pStyle Heading4  decimal  "%1.%2.%3.%4."

    Those pStyle elements are what makes the list appear in Word's Multilevel
    List gallery as a list tied to the heading styles, and what makes
    Heading2 render "5.1." rather than nothing.

    This script is the authority on those definitions - they are constants
    here, not read from a template file.

    Break any link in the chain and Word fails silently. NWIRP chapter 5 had
    lost all of it: ilvl 0 had been overwritten with ListBullet3's Wingdings
    glyph (see the note in OoXml/Numbering/Set-BulletGlyph.ps1), every pStyle
    link was gone, no numStyleLink abstractNum existed, and Heading1-4 carried
    no numPr at all. The chapter title was a typed "Chapter 5" instead.

.PARAMETER ChapterNumber
    The number ilvl 0 starts at, so a standalone chapter file renders its own
    number. Chapter 5's file must start at 5 or its title reads "Chapter 1".

    When omitted, it is read from the number already in the document - HMList's
    ilvl 0 w:start - which is what makes a re-run idempotent. It is NOT taken
    from the filename: the number ends up inside the document and has to
    survive a rename. With no number in the document and none given, the file
    is skipped with a warning rather than numbered from a guess.

.PARAMETER StripTypedNumber
    Removes a hand-typed chapter number from the Heading1 text, and the tab
    after it, so the generated number does not duplicate it. Only ever strips
    text that matches EXACTLY what the numbering will generate for this file
    ("Chapter 5" when -ChapterNumber is 5), on the same principle as
    Set-NISTTableCaptions -StripLeadingNumber. Anything else is reported and
    left alone.

.PARAMETER ListOnly
    Report what is wrong and write nothing.

.EXAMPLE
    .\Set-NISTChapterNumbering.ps1 -File "C:\...\Chapter-05.docx" -ListOnly

.EXAMPLE
    .\Set-NISTChapterNumbering.ps1 -File "C:\...\Chapter-05.docx" -StripTypedNumber -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [ValidateRange(1, 99)]
    [int]$ChapterNumber = 0,

    [switch]$StripTypedNumber,

    [switch]$ListOnly,

    [switch]$Force,

    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Canonical definitions
# ---------------------------------------------------------------------------

# ilvl 0 is rebuilt whole, because it is the level that gets destroyed. The
# rPr formats the generated number text ("Chapter 5"), not the heading text.
function Get-CanonicalLvl0 {
    param([int]$Start)
    return '<w:lvl w:ilvl="0">' +
           '<w:start w:val="' + $Start + '"/>' +
           '<w:numFmt w:val="decimal"/>' +
           '<w:pStyle w:val="Heading1"/>' +
           '<w:lvlText w:val="Chapter %1"/>' +
           '<w:lvlJc w:val="left"/>' +
           '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>' +
           '<w:rPr><w:rFonts w:ascii="Source Serif Pro" w:hAnsi="Source Serif Pro" w:hint="default"/>' +
           '<w:b/><w:i/><w:color w:val="595959" w:themeColor="text1" w:themeTint="A6"/><w:sz w:val="36"/></w:rPr>' +
           '</w:lvl>'
}

# ilvl 1-3 are only checked for their pStyle link. Their lvlText and indents
# are left exactly as the document has them: in the NWIRP chapters those
# already match, and rewriting a level wholesale is what caused the damage
# this script exists to repair.
$LEVEL_STYLES = @{ 1 = 'Heading2'; 2 = 'Heading3'; 3 = 'Heading4' }

# ilvl 1-8, for CREATING the definition in a document that has none. These are
# the certified chapter set's own levels. ilvl 6 is the appendix level: numFmt
# none with lvlText "Appendix %1 | ", which is why an appendix heading reads
# "Appendix B" and not a number. ilvl 7 and 8 are its subordinates.
#
# CT_Lvl is a SEQUENCE: start, numFmt, lvlRestart, pStyle, isLgl, suff,
# lvlText, lvlPicBulletId, legacy, lvlJc, pPr, rPr. pStyle therefore goes after
# numFmt/lvlRestart and before suff, which is also where Add-LvlPStyle puts it.
#
# w:nsid and w:tmpl are deliberately omitted. They identify a list across
# documents, and reusing the certified set's values in an author file risks
# Word merging this definition into one of the author's own lists.
function Get-CanonicalHMListLevels {
    return @(
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:suff w:val="tab"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1026"/></w:tabs><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>',
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading3"/><w:suff w:val="tab"/><w:lvlText w:val="%1.%2.%3."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="936" w:hanging="936"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>',
        '<w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading4"/><w:suff w:val="tab"/><w:lvlText w:val="%1.%2.%3.%4."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="864"/></w:tabs><w:ind w:left="1224" w:hanging="1224"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>',
        '<w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="tab"/><w:lvlText w:val="%1.%2.%3.%4.%5"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1368" w:hanging="1368"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>',
        '<w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="tab"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1656" w:hanging="1656"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>',
        '<w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="none"/><w:lvlRestart w:val="1"/><w:suff w:val="tab"/><w:lvlText w:val="Appendix %1 | "/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1296" w:hanging="1296"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>',
        '<w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:suff w:val="tab"/><w:lvlText w:val="%7%8.%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="648" w:hanging="648"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>',
        '<w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="tab"/><w:lvlText w:val="%8.%9."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>'
    )
}


# ---------------------------------------------------------------------------
# Targets
# ---------------------------------------------------------------------------

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

Write-Host "Set NIST Chapter Numbering" -ForegroundColor Cyan
Write-Host "  Restores HMList ilvl 0, the pStyle links on ilvl 0-3," -ForegroundColor Cyan
Write-Host "  and numPr on Heading1-4." -ForegroundColor Cyan
if ($StripTypedNumber) {
    Write-Host "  -StripTypedNumber: a typed chapter number matching the generated one will be removed." -ForegroundColor Cyan
}
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""
if ($ListOnly) { Write-Host "  -ListOnly: reporting only, writing nothing." -ForegroundColor Yellow }

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

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

function Get-Part {
    param($Pkg, [string]$Name)
    if (-not $Pkg.Data.ContainsKey($Name)) { return $null }
    $t = [System.Text.Encoding]::UTF8.GetString($Pkg.Data[$Name])
    # No String.StartsWith(char) on .NET Framework: the char is coerced and the
    # culture-sensitive compare treats U+FEFF as ignorable, so that test is
    # always true. Check the character.
    if ($t.Length -gt 0 -and $t[0] -eq [char]0xFEFF) { $t = $t.Substring(1) }
    return $t
}

function Get-StyleBlock {
    param([string]$Xml, [string]$Id)
    return [regex]::Match($Xml, '(?s)<w:style\b[^>]*\bw:styleId="' + [regex]::Escape($Id) + '"[^>]*>.*?</w:style>')
}

function Resolve-StyleAbstractNum {
    # style -> numId -> abstractNumId, every lookup attribute-order-safe.
    param([string]$StylesXml, [string]$NumberingXml, [string]$Id)
    $sm = Get-StyleBlock -Xml $StylesXml -Id $Id
    if (-not $sm.Success) { return $null }
    $nid = [regex]::Match($sm.Value, '(?s)<w:numPr>.*?<w:numId\b[^>]*\bw:val="(\d+)"')
    if (-not $nid.Success -or $nid.Groups[1].Value -eq '0') { return $null }
    $nm = [regex]::Match($NumberingXml, '(?s)<w:num\b[^>]*\bw:numId="' + $nid.Groups[1].Value + '"[^>]*>.*?</w:num>')
    if (-not $nm.Success) { return $null }
    $am = [regex]::Match($nm.Value, '<w:abstractNumId\b[^>]*\bw:val="(\d+)"')
    if (-not $am.Success) { return $null }
    return [pscustomobject]@{ NumId = $nid.Groups[1].Value; AbstractNumId = $am.Groups[1].Value }
}

function Set-NumPr {
    # Writes a numPr into a style block, replacing any existing one. w:pPr
    # child order puts numPr after pStyle/keepNext/keepLines and before
    # spacing/ind/jc/outlineLvl, so the insertion points are tried in that
    # order rather than blindly at the front of pPr.
    param([string]$Block, [int]$Ilvl, [string]$NumId)
    $numPr = '<w:numPr><w:ilvl w:val="' + $Ilvl + '"/><w:numId w:val="' + $NumId + '"/></w:numPr>'
    if ($Block -match '(?s)<w:numPr>.*?</w:numPr>') {
        return ([regex]'(?s)<w:numPr>.*?</w:numPr>').Replace($Block, $numPr, 1)
    }
    foreach ($anchor in @('<w:spacing\b', '<w:ind\b', '<w:jc\b', '<w:outlineLvl\b')) {
        if ($Block -match $anchor) {
            return ([regex]('(' + $anchor + ')')).Replace($Block, ($numPr + '$1'), 1)
        }
    }
    if ($Block -match '</w:pPr>') { return ([regex]'</w:pPr>').Replace($Block, ($numPr + '</w:pPr>'), 1) }
    if ($Block -match '<w:pPr\s*/>') { return ([regex]'<w:pPr\s*/>').Replace($Block, ('<w:pPr>' + $numPr + '</w:pPr>'), 1) }
    $bo = [regex]::Match($Block, '<w:basedOn\b[^>]*/>')
    if ($bo.Success) {
        return $Block.Insert($bo.Index + $bo.Length, '<w:pPr>' + $numPr + '</w:pPr>')
    }
    return ([regex]'</w:style>').Replace($Block, ('<w:pPr>' + $numPr + '</w:pPr></w:style>'), 1)
}

function Add-LvlPStyle {
    # CT_Lvl is a sequence: start, numFmt, lvlRestart, pStyle, isLgl, suff,
    # lvlText, lvlPicBulletId, legacy, lvlJc, pPr, rPr. pStyle therefore goes
    # after lvlRestart/numFmt/start and before everything from isLgl on.
    param([string]$Lvl, [string]$StyleId)
    if ($Lvl -match '<w:pStyle\b') {
        return ([regex]'<w:pStyle\b[^>]*/>').Replace($Lvl, ('<w:pStyle w:val="' + $StyleId + '"/>'), 1)
    }
    $el = '<w:pStyle w:val="' + $StyleId + '"/>'
    foreach ($anchor in @('<w:lvlRestart\b[^>]*/>', '<w:numFmt\b[^>]*/>', '<w:start\b[^>]*/>')) {
        $m = [regex]::Match($Lvl, $anchor)
        if ($m.Success) { return $Lvl.Insert($m.Index + $m.Length, $el) }
    }
    foreach ($anchor in @('<w:isLgl\b', '<w:suff\b', '<w:lvlText\b', '<w:lvlJc\b', '<w:pPr\b', '<w:rPr\b')) {
        $m = [regex]::Match($Lvl, $anchor)
        if ($m.Success) { return $Lvl.Insert($m.Index, $el) }
    }
    return ([regex]'</w:lvl>').Replace($Lvl, ($el + '</w:lvl>'), 1)
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$filesChanged = 0

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    $stylesXml = Get-Part -Pkg $pkg -Name 'word/styles.xml'
    $numXml    = Get-Part -Pkg $pkg -Name 'word/numbering.xml'
    $docXml    = Get-Part -Pkg $pkg -Name 'word/document.xml'
    if (-not $stylesXml) { Write-Warning "  No word/styles.xml. Skipping."; continue }
    if (-not $numXml)    { Write-Warning "  No word/numbering.xml. Skipping."; continue }

    # --- Chapter number -------------------------------------------------
    # NOT from the filename. The number ends up inside the document, so it has
    # to survive a rename - and a filename is metadata anyone can change, which
    # is exactly what happened the first time this shipped. Two sources, in
    # order:
    #
    #   1. -ChapterNumber, which the operator states.
    #   2. The number already in the document: HMList's ilvl 0 w:start. That IS
    #      the chapter number, it is self-describing, and reading it is what
    #      makes a re-run idempotent instead of resetting every chapter to 1.
    #
    # With neither, this REFUSES. Guessing produced a set where several
    # chapters silently numbered themselves 1.
    $chapter = $ChapterNumber
    if ($chapter -eq 0) {
        $existing = 0
        $hmProbe = Resolve-StyleAbstractNum -StylesXml $stylesXml -NumberingXml $numXml -Id 'HMList'
        $probeAbs = if ($hmProbe) { $hmProbe.AbstractNumId } else { $null }
        if (-not $probeAbs) {
            $slp = [regex]::Match($numXml, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"[^>]*>(?:(?!</w:abstractNum>).)*?<w:styleLink\b[^>]*\bw:val="HMList"')
            if ($slp.Success) { $probeAbs = $slp.Groups[1].Value }
        }
        if ($probeAbs) {
            $ab = [regex]::Match($numXml, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="' + $probeAbs + '"[^>]*>.*?</w:abstractNum>')
            if ($ab.Success) {
                $lv = [regex]::Match($ab.Value, '(?s)<w:lvl\b[^>]*\bw:ilvl="0"[^>]*>.*?</w:lvl>')
                if ($lv.Success) {
                    $stm = [regex]::Match($lv.Value, '<w:start\b[^>]*\bw:val="(\d+)"')
                    if ($stm.Success) { $existing = [int]$stm.Groups[1].Value }
                }
            }
        }
        if ($existing -gt 0) {
            $chapter = $existing
            Write-Host "  Chapter number $chapter read from the document's own HMList." -ForegroundColor DarkGray
        } else {
            Write-Warning "  This document carries no chapter number yet. Give -ChapterNumber; it is not guessed from the filename."
            continue
        }
    } else {
        Write-Host "  Chapter number $chapter (given)." -ForegroundColor DarkGray
    }

    $changes = New-Object System.Collections.ArrayList

    # --- Locate the HMList abstractNum ----------------------------------
    # Prefer the style's own chain; fall back to the styleLink marker, which is
    # what identifies a list definition owned by a named list style.
    $absId = $null
    $hm = Resolve-StyleAbstractNum -StylesXml $stylesXml -NumberingXml $numXml -Id 'HMList'
    if ($hm) { $absId = $hm.AbstractNumId }
    if (-not $absId) {
        $slm = [regex]::Match($numXml, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"[^>]*>(?:(?!</w:abstractNum>).)*?<w:styleLink\b[^>]*\bw:val="HMList"')
        if ($slm.Success) { $absId = $slm.Groups[1].Value }
    }
    if (-not $absId) {
        # An author chapter has no HMList: its headings run off whatever
        # multilevel list the author happened to start, so each file numbers
        # differently and none of them restarts at its own chapter. Build the
        # definition rather than skip, because skipping is what left seven of
        # eight chapters unnumbered.
        #
        # Three pieces, in this order: the base abstractNum carrying
        # styleLink HMList and the nine levels; a num pointing at it; and the
        # numbering STYLE whose numPr names that num. Without the style the
        # list has no name, so it never appears in Word's Multilevel List
        # gallery and no heading can be attached to it by name.
        $absIdsNow = @([regex]::Matches($numXml, '<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"') | ForEach-Object { [int]$_.Groups[1].Value })
        $numIdsNow = @([regex]::Matches($numXml, '<w:num\b[^>]*\bw:numId="(\d+)"') | ForEach-Object { [int]$_.Groups[1].Value })
        $newAbsId = 1 + ((@($absIdsNow) + 0 | Measure-Object -Maximum).Maximum)
        $newNumId = 1 + ((@($numIdsNow) + 0 | Measure-Object -Maximum).Maximum)

        $block = '<w:abstractNum w:abstractNumId="' + $newAbsId + '" w15:restartNumberingAfterBreak="0">' +
                 '<w:multiLevelType w:val="multilevel"/>' +
                 '<w:styleLink w:val="HMList"/>' +
                 (Get-CanonicalLvl0 -Start $chapter) +
                 ((Get-CanonicalHMListLevels) -join '') +
                 '</w:abstractNum>'
        $numEl = '<w:num w:numId="' + $newNumId + '"><w:abstractNumId w:val="' + $newAbsId + '"/></w:num>'

        # CT_Numbering is a sequence: every abstractNum precedes every num. Go
        # in ahead of the FIRST num, or at the end when there is none.
        $firstNum = [regex]::Match($numXml, '<w:num\b')
        if ($firstNum.Success) {
            $numXml = $numXml.Insert($firstNum.Index, $block)
            $firstNum2 = [regex]::Match($numXml, '<w:num\b[^>]*\bw:numId=')
            $numXml = $numXml.Insert($firstNum2.Index, $numEl)
        } else {
            $numXml = [regex]::Replace($numXml, '</w:numbering>', $block + $numEl + '</w:numbering>', 1)
        }

        if ($stylesXml -notmatch '<w:style\b[^>]*\bw:styleId="HMList"') {
            $hmStyle = '<w:style w:type="numbering" w:customStyle="1" w:styleId="HMList">' +
                       '<w:name w:val="HM List"/><w:uiPriority w:val="54"/><w:qFormat/>' +
                       '<w:pPr><w:numPr><w:numId w:val="' + $newNumId + '"/></w:numPr></w:pPr>' +
                       '</w:style>'
            $stylesXml = [regex]::Replace($stylesXml, '</w:styles>', $hmStyle + '</w:styles>', 1)
            [void]$changes.Add("created numbering style HMList -> numId $newNumId")
        } else {
            # the style is there but its numPr did not resolve: repoint it
            $hmM = [regex]::Match($stylesXml, '(?s)<w:style\b[^>]*\bw:styleId="HMList".*?</w:style>')
            $fixed = [regex]::Replace($hmM.Value, '<w:numId\b[^>]*\bw:val="\d+"\s*/>', '<w:numId w:val="' + $newNumId + '"/>', 1)
            $stylesXml = $stylesXml.Remove($hmM.Index, $hmM.Length).Insert($hmM.Index, $fixed)
            [void]$changes.Add("repointed numbering style HMList -> numId $newNumId")
        }
        $absId = [string]$newAbsId
        [void]$changes.Add("created HMList list definition: abstractNum $newAbsId, num $newNumId, nine levels")
    }
    Write-Host "  HMList list definition: abstractNum $absId" -ForegroundColor DarkGray

    $abM = [regex]::Match($numXml, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="' + $absId + '"[^>]*>.*?</w:abstractNum>')
    if (-not $abM.Success) { Write-Warning "  abstractNum $absId not found. Skipping."; continue }
    $abBlock = $abM.Value
    $newAb   = $abBlock

    # --- ilvl 0 ----------------------------------------------------------
    $lvl0M = [regex]::Match($newAb, '(?s)<w:lvl\b[^>]*\bw:ilvl="0"[^>]*>.*?</w:lvl>')
    $canonical0 = Get-CanonicalLvl0 -Start $chapter
    if (-not $lvl0M.Success) {
        $newAb = ([regex]'(<w:lvl\b)').Replace($newAb, ($canonical0 + '$1'), 1)
        [void]$changes.Add("ilvl 0 created (start $chapter)")
    } else {
        $cur = $lvl0M.Value
        $curFmt = [regex]::Match($cur, '<w:numFmt\b[^>]*\bw:val="([^"]+)"')
        $curTxt = [regex]::Match($cur, '<w:lvlText\b[^>]*\bw:val="([^"]*)"')
        $ok = ($curFmt.Success -and $curFmt.Groups[1].Value -eq 'decimal') -and
              ($curTxt.Success -and $curTxt.Groups[1].Value -eq 'Chapter %1') -and
              ($cur -match '<w:pStyle\b[^>]*\bw:val="Heading1"')
        $curStart = [regex]::Match($cur, '<w:start\b[^>]*\bw:val="(\d+)"')
        $startOk = ($curStart.Success -and [int]$curStart.Groups[1].Value -eq $chapter)
        if (-not $ok) {
            $was = if ($curTxt.Success) { $curTxt.Groups[1].Value } else { '(no lvlText)' }
            $newAb = $newAb.Remove($lvl0M.Index, $lvl0M.Length).Insert($lvl0M.Index, $canonical0)
            [void]$changes.Add("ilvl 0 rebuilt (was lvlText '$was') -> decimal 'Chapter %1', pStyle Heading1, start $chapter")
        } elseif (-not $startOk) {
            $newAb = $newAb.Remove($lvl0M.Index, $lvl0M.Length).Insert($lvl0M.Index, $canonical0)
            [void]$changes.Add("ilvl 0 start -> $chapter")
        }
    }

    # --- ilvl 1-3 pStyle links -------------------------------------------
    foreach ($ilvl in @(1, 2, 3)) {
        $styleId = $LEVEL_STYLES[$ilvl]
        $lm = [regex]::Match($newAb, '(?s)<w:lvl\b[^>]*\bw:ilvl="' + $ilvl + '"[^>]*>.*?</w:lvl>')
        if (-not $lm.Success) { [void]$changes.Add("ilvl $ilvl absent - no pStyle link added"); continue }
        $existing = [regex]::Match($lm.Value, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
        if ($existing.Success -and $existing.Groups[1].Value -eq $styleId) { continue }
        $fixed = Add-LvlPStyle -Lvl $lm.Value -StyleId $styleId
        $newAb = $newAb.Remove($lm.Index, $lm.Length).Insert($lm.Index, $fixed)
        if ($existing.Success) {
            [void]$changes.Add("ilvl $ilvl pStyle $($existing.Groups[1].Value) -> $styleId")
        } else {
            [void]$changes.Add("ilvl $ilvl pStyle -> $styleId")
        }
    }

    if ($newAb -ne $abBlock) {
        $numXml = $numXml.Remove($abM.Index, $abM.Length).Insert($abM.Index, $newAb)
    }

    # --- The numStyleLink abstractNum and its num ------------------------
    # Heading1-4 must point at an abstractNum that only forwards to HMList, not
    # at the HMList definition itself: that indirection is what lets the list
    # be edited in one place.
    $linkAbs = $null
    $lam = [regex]::Match($numXml, '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"[^>]*>(?:(?!</w:abstractNum>).)*?<w:numStyleLink\b[^>]*\bw:val="HMList"')
    if ($lam.Success) {
        $linkAbs = $lam.Groups[1].Value
    } else {
        $absIds = @([regex]::Matches($numXml, '<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"') |
                    ForEach-Object { [int]$_.Groups[1].Value })
        $linkAbs = [string](1 + (($absIds + 0) | Measure-Object -Maximum).Maximum)
        $block = '<w:abstractNum w:abstractNumId="' + $linkAbs + '">' +
                 '<w:multiLevelType w:val="multilevel"/>' +
                 '<w:numStyleLink w:val="HMList"/></w:abstractNum>'
        # Insert before the FIRST num: abstractNum must precede every num.
        if ($numXml -match '<w:num\b') {
            $numXml = ([regex]'(<w:num\b)').Replace($numXml, ($block + '$1'), 1)
        } else {
            $numXml = ([regex]'</w:numbering>').Replace($numXml, ($block + '</w:numbering>'), 1)
        }
        [void]$changes.Add("created abstractNum $linkAbs (numStyleLink HMList)")
    }

    $headNumId = $null
    $hnm = [regex]::Match($numXml, '(?s)<w:num\b[^>]*\bw:numId="(\d+)"[^>]*>(?:(?!</w:num>).)*?<w:abstractNumId\b[^>]*\bw:val="' + $linkAbs + '"')
    if ($hnm.Success) {
        $headNumId = $hnm.Groups[1].Value
    } else {
        $numIds = @([regex]::Matches($numXml, '<w:num\b[^>]*\bw:numId="(\d+)"') |
                    ForEach-Object { [int]$_.Groups[1].Value })
        $headNumId = [string](1 + (($numIds + 0) | Measure-Object -Maximum).Maximum)
        $numXml = ([regex]'</w:numbering>').Replace(
            $numXml,
            ('<w:num w:numId="' + $headNumId + '"><w:abstractNumId w:val="' + $linkAbs + '"/></w:num></w:numbering>'), 1)
        [void]$changes.Add("created num $headNumId -> abstractNum $linkAbs")
    }

    # --- Heading1-4 numPr -------------------------------------------------
    foreach ($ilvl in @(0, 1, 2, 3)) {
        $styleId = if ($ilvl -eq 0) { 'Heading1' } else { $LEVEL_STYLES[$ilvl] }
        $sm = Get-StyleBlock -Xml $stylesXml -Id $styleId
        if (-not $sm.Success) { [void]$changes.Add("$styleId absent - not wired"); continue }
        $curNum = [regex]::Match($sm.Value, '(?s)<w:numPr>.*?<w:numId\b[^>]*\bw:val="(\d+)"')
        $curIlvl = [regex]::Match($sm.Value, '(?s)<w:numPr>.*?<w:ilvl\b[^>]*\bw:val="(\d+)"')
        $already = $curNum.Success -and $curNum.Groups[1].Value -eq $headNumId -and
                   (($ilvl -eq 0 -and (-not $curIlvl.Success -or $curIlvl.Groups[1].Value -eq '0')) -or
                    ($curIlvl.Success -and $curIlvl.Groups[1].Value -eq [string]$ilvl))
        if ($already) { continue }
        $fixed = Set-NumPr -Block $sm.Value -Ilvl $ilvl -NumId $headNumId
        $stylesXml = $stylesXml.Remove($sm.Index, $sm.Length).Insert($sm.Index, $fixed)
        $had = if ($curNum.Success) { "numId $($curNum.Groups[1].Value)" } else { 'no numPr' }
        [void]$changes.Add("$styleId ($had) -> ilvl $ilvl, numId $headNumId")
    }

    # --- Typed chapter number in the Heading1 text ------------------------
    $typedFound = $null
    if ($docXml) {
        $expected = "Chapter $chapter"
        foreach ($pm in [regex]::Matches($docXml, '(?s)<w:p\b[^>]*>.*?</w:p>')) {
            if ($pm.Value -notmatch '<w:pStyle\b[^>]*\bw:val="Heading1"') { continue }
            # A run holding exactly the generated number, followed by a run
            # that begins with the tab the numbering suffix supplies.
            $rm = [regex]::Match($pm.Value,
                '<w:r(?:\s[^>]*)?>(?:<w:rPr\b(?:[^>]*/>|[^>]*>.*?</w:rPr>))?<w:t(?:\s[^>]*)?>' +
                [regex]::Escape($expected) + '</w:t></w:r>')
            if (-not $rm.Success) { continue }
            $typedFound = [pscustomobject]@{ ParaIndex = $pm.Index; ParaLength = $pm.Length; Para = $pm.Value; RunIndex = $rm.Index; RunLength = $rm.Length }
            break
        }
    }
    if ($typedFound) {
        if ($StripTypedNumber) {
            $newPara = $typedFound.Para.Remove($typedFound.RunIndex, $typedFound.RunLength)
            # The tab that separated the typed number from the title is now
            # supplied by the numbering suffix, so drop the first leading tab.
            # .NET has no \K, so the run prefix is captured and written back.
            $newPara = ([regex]'(<w:r(?:\s[^>]*)?>(?:<w:rPr\b(?:[^>]*/>|[^>]*>.*?</w:rPr>))?)<w:tab\s*/>').Replace($newPara, '$1', 1)
            $docXml = $docXml.Remove($typedFound.ParaIndex, $typedFound.ParaLength).Insert($typedFound.ParaIndex, $newPara)
            [void]$changes.Add("removed typed 'Chapter $chapter' and its tab from the Heading1 text")
        } else {
            Write-Host "  Heading1 has a typed 'Chapter $chapter'. The generated number will DUPLICATE it." -ForegroundColor DarkYellow
            Write-Host "  Re-run with -StripTypedNumber to remove it." -ForegroundColor DarkYellow
        }
    }

    # --- Report and save ---------------------------------------------------
    if ($changes.Count -eq 0) {
        Write-Host "  Chapter numbering already correct. Nothing to do." -ForegroundColor Green
        continue
    }
    foreach ($c in $changes) { Write-Host "    $c" -ForegroundColor DarkGray }

    foreach ($pair in @(@{ N='word/styles.xml'; V=$stylesXml }, @{ N='word/numbering.xml'; V=$numXml }, @{ N='word/document.xml'; V=$docXml })) {
        if (-not $pair.V) { continue }
        try {
            $probe = New-Object System.Xml.XmlDocument
            $probe.XmlResolver = $null
            $probeXml = $pair.V; if ($probeXml.Length -gt 0 -and $probeXml[0] -eq [char]0xFEFF) { $probeXml = $probeXml.Substring(1) }; $probe.LoadXml($probeXml)
        } catch {
            throw "Rewrite produced malformed $($pair.N) for ${target}: $($_.Exception.Message)"
        }
    }

    # numbering.xml order must still hold.
    $seq = @([regex]::Matches($numXml, '<w:(numPicBullet|abstractNum|num)\b') | ForEach-Object { $_.Groups[1].Value })
    $rank = @{ 'numPicBullet' = 0; 'abstractNum' = 1; 'num' = 2 }
    $prev = -1
    foreach ($t in $seq) {
        if ($rank[$t] -lt $prev) { throw "Rewrite left numbering.xml out of schema order for ${target}." }
        $prev = [Math]::Max($prev, $rank[$t])
    }

    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        if (-not $Force -and (Test-Path -LiteralPath "$target.chapternum.bak")) {
            Write-Warning "  Backup $target.chapternum.bak already exists; leaving it as-is."
        } else {
            Copy-Item -LiteralPath $target -Destination "$target.chapternum.bak" -Force
        }
    }

    $pkg.Data['word/styles.xml']    = $utf8NoBom.GetBytes($stylesXml)
    $pkg.Data['word/numbering.xml'] = $utf8NoBom.GetBytes($numXml)
    if ($docXml) { $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($docXml) }

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "chapternum-$([guid]::NewGuid()).zip")
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
