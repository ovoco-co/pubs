<#
.SYNOPSIS
    Applies the agreed NIST NCSTAR / NWIRP style architecture to a Word
    file's word/styles.xml: rebases styles onto the xx* family, deletes
    junk styles and table-style bloat, fixes basedOn references, and
    applies the color backfill from NWIRP to all targets.

.DESCRIPTION
    Backs up the original to <name>.styles.bak, then makes the following
    changes inside word/styles.xml:

    REBASE (change basedOn to break Normal inheritance):
      Heading1-9, Series                               -> xxheader
      Front_matter_head1-4                             -> Heading1-4 (match, no numbering)
      Caption                                          -> xxheader (Semibold 10pt, no bold)
      Caption_Table, Caption_Figure                    -> Caption
      BodyText                                         -> xxbase
      List Bullet, List Bullet 2-5                     -> xxbase
      List Number, List Number 2-5                     -> xxbase
      Table and Figure Footnotes                       -> xxtable
      Verso_head                                       -> xxheader
      Verso_text                                       -> xxbase
      Sidebar                                          -> xxtable (sans callout)
      Sidebar Bullet                                   -> Sidebar (parallel)
      Term-Abbrev                                      -> Frontmatterhead4
      Appendix_Head4                                   -> AppendixHead (parallel)
      Findings_, RecommendFindingData                  -> xxbase

    DELETE (junk styles, MS-Office paste artifacts, NWIRP-only list ladders):
      Crossmark, Frontmatterhead (unnumbered)
      msonormal, paragraph, BasicParagraph, Default, Style1
      eop, header-section-number, normaltextrun, superscript, math
      BodyText2, BodyText3, BodyTextIndent, BodyTextIndent2/3
      BodyTextFirstIndent, BodyTextFirstIndent2 (and Char siblings)
      TableParagraph
      BulletListLevel1/2/3, NumberedListLevel1/2/3

    DELETE (table-style bloat): every table-type style except
      TableNormal, TableGrid, HMCover, HMData

    COLORS BACKFILL (apply NWIRP's choices everywhere):
      header / Header Char  -> #073B4C
      caption               -> #00257A
      Verso_head            -> #000000
      Quote / Quote Char    -> #404040
      Appendix_Head4        -> #833C0B

    The script uses regex on the styles.xml text. It does not touch
    document.xml, so any pStyle references that pointed at a deleted
    style become orphaned (Word will fall back to Normal for those).
    Run this after applying Set-NISTTheme.ps1 since several decisions
    assume the new theme is in place.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. Applies to every Word file in the folder.

.PARAMETER Clean
    Clean-input mode. Applies the style architecture (canonical definitions,
    colors, indents, numbering, ligatures-off, recommended curation) but skips
    the legacy-fix passes that only exist to rescue messy inherited documents:
    the Heading 1/3 cs="Times New Roman" strip and the junk / paste-artifact
    style deletions. Use on the clean template or documents built from it.
    Set-NISTTemplateStyles.ps1 is a thin wrapper that calls this with -Clean.

.PARAMETER Force
    Skip prompts.

.EXAMPLE
    .\Set-NISTStyleArchitecture.ps1 -File "\\blue\E\path\to\template.dotx"

.EXAMPLE
    # Clean template workflow (no legacy fixes)
    .\Set-NISTStyleArchitecture.ps1 -File "HM-NCSTAR-Template.docx" -Clean

.EXAMPLE
    .\Set-NISTStyleArchitecture.ps1 -Folder "\\blue\E\00-clients\nist-HM\Templates" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$Clean,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Import the generic OoXml module for style and numbering manipulation
$ooXmlModule = Join-Path $PSScriptRoot 'OoXml\OoXml.psd1'
Import-Module $ooXmlModule -Force

# ---------------------------------------------------------------------------
# Operations
# ---------------------------------------------------------------------------

# Rebase: styleId -> new basedOn styleId
$rebase = [ordered]@{
    'Heading1'                 = 'xxheader'
    'Heading2'                 = 'xxheader'
    'Heading3'                 = 'xxheader'
    'Heading4'                 = 'xxheader'
    'Heading5'                 = 'xxheader'
    'Heading6'                 = 'xxheader'
    'Heading7'                 = 'xxheader'
    'Heading8'                 = 'xxheader'
    'Heading9'                 = 'xxheader'
    'BodyText'                 = 'xxbase'
    'Caption'                  = 'xxheader'
    # Caption_Table and Caption_Figure carry the caption look, so they hang off
    # Caption, not off xxheader. Based on xxheader they picked up the heading
    # engine's 12pt and bold and rendered as headings. Both styleId spellings
    # are listed because Word strips the underscore when it creates the id.
    'CaptionTable'             = 'Caption'
    'Caption_Table'            = 'Caption'
    'CaptionFigure'            = 'Caption'
    'Caption_Figure'           = 'Caption'
    'ListBullet'               = 'xxbase'
    'ListBullet2'              = 'xxbase'
    'ListBullet3'              = 'xxbase'
    'ListBullet4'              = 'xxbase'
    'ListBullet5'              = 'xxbase'
    'ListNumber'               = 'xxbase'
    'ListNumber2'              = 'xxbase'
    'ListNumber3'              = 'xxbase'
    'ListNumber4'              = 'xxbase'
    'ListNumber5'              = 'xxbase'
    'ListContinue'             = 'xxbase'
    'ListContinue2'            = 'xxbase'
    'ListContinue3'            = 'xxbase'
    'ListContinue4'            = 'xxbase'
    'ListContinue5'            = 'xxbase'
    'List'                     = 'xxbase'
    'List2'                    = 'xxbase'
    'List3'                    = 'xxbase'
    'List4'                    = 'xxbase'
    'List5'                    = 'xxbase'
    'ListParagraph'            = 'xxbase'
    'header'                   = 'xxheader'
    'footer'                   = 'xxheader'
    'TableandFigureFootnotes'  = 'xxtable'
    'Versohead'                = 'xxheader'
    'Versotext'                = 'xxbase'
    # Front matter heads 1-4 render exactly like Heading 1-4 (2-4 already were).
    'Frontmatterhead1'         = 'Heading1'
    'Frontmatterhead2'         = 'Heading2'
    'Frontmatterhead3'         = 'Heading3'
    'Frontmatterhead4'         = 'Heading4'
    'Term-Abbrev'              = 'Frontmatterhead4'
    # AppendixHead 1-4 render exactly like Heading 1-4: base each on its match.
    'AppendixHead'             = 'Heading1'
    'AppendixHead2'            = 'Heading2'
    'AppendixHead3'            = 'Heading3'
    'AppendixHead4'            = 'Heading4'
    'Series'                   = 'xxheader'
    'Findings_'                = 'xxbase'
    'RecommendFindingData'     = 'xxbase'
    # Sidebar callout family: root on the sans table-text engine (9pt Source
    # Sans Pro) so it stays a managed style; SidebarBullet keeps Sidebar as its
    # base for the frame + bullet numbering.
    'Sidebar'                  = 'xxtable'
    'SidebarBullet'            = 'Sidebar'
}

# Delete: styleIds to remove from styles.xml
$delete = @(
    # Junk / Normal-rooted styles
    'Crossmark', 'Frontmatterhead',
    # Word auto-generated "Style Heading 1_NoNumber + Before: 0 pt" linked-override artifact
    'StyleHeading1NoNumberBefore0pt',
    # MS-Office paste artifacts
    'msonormal', 'paragraph', 'BasicParagraph', 'Default', 'Style1',
    'eop', 'header-section-number', 'normaltextrun',
    # MathType styles preserved: MTDisplayEquation, MTDisplayEquationChar,
    # MTEquationSection, math, superscript
    # Body Text variant zoo (paragraph)
    'BodyText2', 'BodyText3',
    'BodyTextIndent', 'BodyTextIndent2', 'BodyTextIndent3',
    'BodyTextFirstIndent', 'BodyTextFirstIndent2',
    # Char variants of the above
    'BodyText2Char', 'BodyText3Char',
    'BodyTextIndentChar', 'BodyTextIndent2Char', 'BodyTextIndent3Char',
    'BodyTextFirstIndentChar', 'BodyTextFirstIndent2Char',
    # Table junk
    'TableParagraph',
    # NWIRP-only list ladders (replaced by built-in List Bullet / List Number)
    'BulletListLevel1', 'BulletListLevel2', 'BulletListLevel3',
    'NumberedListLevel1', 'NumberedListLevel2', 'NumberedListLevel3',
    # Other styles to remove (both bare and underscore-prefixed variants,
    # since Word strips underscores from styleIds but some files may keep them)
    'ChapterDocTitle', '_Chapter_Doc_Title',
    'Boldaccent', '_Boldaccent',
    'Checklistheader', 'Checklist_header',
    'Checklisttext', 'Checklist_text'
)

# Table-style allowlist: every style of type="table" NOT in this list gets deleted.
$tableKeep = @('TableNormal', 'TableGrid', 'HM-Cover', 'HM-Data', 'NIST-Data', 'SidebarTable')

# Colors to set on specific styles (styleId -> hex without #).
# The color is written as a hard-coded hex value (<w:color w:val="00257A"/>)
# with NO w:themeColor reference. This locks the color to the style so it
# survives Word's Import Styles / Organizer into a document with a different
# theme. (A theme reference would re-resolve against the destination's palette
# and reset to that document's slot color, e.g. the "Word 2003" default.)
# Trade-off: Design > Colors no longer recolors these styles automatically.
$setColor = [ordered]@{
    'header'         = '00257A'
    'HeaderChar'     = '00257A'
    'footer'         = '00257A'
    'FooterChar'     = '00257A'
    'caption'        = '00257A'
    'Heading1'       = '00257A'
    'Heading2'       = '00257A'
    'Heading3'       = '00257A'
    'Heading4'       = '00257A'
    'Heading5'       = '404040'   # neutral dark gray, distinct from the H1-H4 Resolution Blue
    'Heading6'       = '404040'   # match Heading5: the deep levels drop out of the blue
    'Heading7'       = '00257A'
    'Heading8'       = '00257A'
    'Heading9'       = '00257A'
    'Heading1Char'   = '00257A'
    'Heading2Char'   = '00257A'
    'Heading3Char'   = '00257A'
    'Heading4Char'   = '00257A'
    'Heading5Char'   = '404040'   # match Heading5
    'Heading6Char'   = '404040'   # match Heading6
    'Heading7Char'   = '00257A'
    'Heading8Char'   = '00257A'
    'Heading9Char'   = '00257A'
    'AppendixHead'   = '00257A'
    'AppendixHead2'  = '00257A'
    'AppendixHead3'  = '00257A'
    'Quote'          = '404040'
    'QuoteChar'      = '404040'
    'AppendixHead4'  = '00257A'
    # Front matter heads match Heading 1-4 exactly, including the Resolution
    # Blue color (per request). Versohead is the one that stays black.
    'Frontmatterhead1' = '00257A'
    'Frontmatterhead2' = '00257A'
    'Frontmatterhead3' = '00257A'
    'Frontmatterhead4' = '00257A'
}

# NIST palette reference (documentation only). The style colors above are
# written as fixed hex, NOT linked to these theme slots, so they never reset
# when styles are imported into a document with a different theme.
#   00257A  Resolution Blue   (Study/NIST primary)
#   D55E00  Tawny             (Schools)
#   009E73  Shamrock Green    (Hospitals)
#   990B38  Claret            (Disruptions)
#   56B4E9  Picton Blue
#   C15C94  Mulberry
#   355195  NCSTAR Secondary

# Set the default paragraph style inside a table style to TableText (instead of
# Normal). Inserts <w:pPr><w:pStyle w:val="TableText"/></w:pPr> only at the
# table-style root (just before the root <w:tblPr>). Idempotent: if the exact
# block is already present at the root, nothing is added. Never touches pPr
# blocks nested inside <w:tblStylePr> (conditional formatting), which previously
# caused corruption.
$tableUseTableText = {
    param($block)
    $newPPr = '<w:pPr><w:pStyle w:val="TableText"/></w:pPr>'
    # Inspect only the prefix before the first <w:tblPr> (the table-style root)
    $tblPrIdx = $block.IndexOf('<w:tblPr>')
    if ($tblPrIdx -lt 0) { return $block }
    $prefix = $block.Substring(0, $tblPrIdx)
    if ($prefix.Contains('<w:pStyle w:val="TableText"/>')) {
        return $block  # already inserted
    }
    return $prefix + $newPPr + $block.Substring($tblPrIdx)
}

# Table Text: ensure <w:next w:val="TableText"/> so Enter in a Table Text
# paragraph keeps the next paragraph in Table Text.
$tableTextEdit = {
    param($block)
    if ($block -match '<w:next\s+w:val="[^"]*"\s*/>') {
        return [regex]::Replace($block, '<w:next\s+w:val="[^"]*"\s*/>', '<w:next w:val="TableText"/>', 1)
    }
    if ($block -match '<w:basedOn\s+w:val="[^"]*"\s*/>') {
        return [regex]::Replace($block, '(<w:basedOn\s+w:val="[^"]*"\s*/>)', '$1<w:next w:val="TableText"/>', 1)
    }
    if ($block -match '<w:name\s+w:val="[^"]*"\s*/>') {
        return [regex]::Replace($block, '(<w:name\s+w:val="[^"]*"\s*/>)', '$1<w:next w:val="TableText"/>', 1)
    }
    return $block
}

# Canonical NIST-Data table style. Installed fresh into styles.xml on every run,
# overwriting any existing NIST-Data block. Authoritative source: this script.
# Includes firstRow (Resolution Blue header, white bold), lastRow (light yellow
# total, bold), band1Horz (light blue striped rows), firstCol (light blue tint).
# To use: in Word, place cursor in a table, open Table Design ribbon, pick
# NIST-Data from the gallery, tick Header Row + Total Row + Banded Rows.
$NIST_DATA_CANONICAL = '<w:style w:type="table" w:customStyle="1" w:styleId="NIST-Data"><w:name w:val="NIST-Data"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="52"/><w:qFormat/><w:rsid w:val="00B64025"/><w:pPr><w:pStyle w:val="TableText"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:jc w:val="center"/><w:tblBorders><w:top w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/><w:left w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/><w:bottom w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/><w:right w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/><w:insideH w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/><w:insideV w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/></w:tblBorders><w:tblCellMar><w:top w:w="14" w:type="dxa"/><w:left w:w="58" w:type="dxa"/><w:bottom w:w="14" w:type="dxa"/><w:right w:w="29" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:jc w:val="center"/></w:trPr><w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:bottom w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="00257A"/><w:vAlign w:val="center"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="lastRow"><w:rPr><w:b/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="firstCol"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr></w:tblStylePr></w:style>'

# Canonical Sidebar table style. A shaded callout box: whole-table light amber
# fill (FFF2CC) via the style-level tcPr, thin gray borders, sans table text
# (TableText), fixed 3" width (4325 twips) with fixed layout. Fixed hex fill so
# it survives Import Styles. Installed fresh on every run. Named "Sidebar Table"
# (styleId SidebarTable) so it does not collide with the "Sidebar" paragraph
# style. The right-margin float / text-wrapping (tblpPr) cannot live in a table
# style (the schema only allows it on the table instance), so Set-TableLook
# stamps it onto every table that uses this style.
$SIDEBAR_TABLE_CANONICAL = '<w:style w:type="table" w:customStyle="1" w:styleId="SidebarTable"><w:name w:val="Sidebar Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="53"/><w:qFormat/><w:pPr><w:pStyle w:val="TableText"/></w:pPr><w:tblPr><w:tblW w:w="4325" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:color="808080" w:sz="4" w:space="0"/><w:left w:val="single" w:color="808080" w:sz="4" w:space="0"/><w:bottom w:val="single" w:color="808080" w:sz="4" w:space="0"/><w:right w:val="single" w:color="808080" w:sz="4" w:space="0"/><w:insideH w:val="single" w:color="808080" w:sz="4" w:space="0"/><w:insideV w:val="single" w:color="808080" w:sz="4" w:space="0"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="58" w:type="dxa"/><w:left w:w="58" w:type="dxa"/><w:bottom w:w="58" w:type="dxa"/><w:right w:w="58" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:tcPr></w:style>'

# HM-Data: lock the lastRow yellow (drop themeFill so it cannot drift with theme),
# set the firstCol shading to F2F2F2, unify every rule/gridline to hard-coded
# A6A6A6, and set default cell margins:
# top=14 (0.01"), bottom=14 (0.01"), left=58 (0.04"), right=29 (0.02").
$hmDataEdit = {
    param($block)
    # Drop themeFill/themeFillTint attributes from any <w:shd> that has a fill="FFF2CC"
    # (the lastRow yellow). Keeps the explicit hex so it cannot be overridden by theme.
    $b = [regex]::Replace($block,
        '(<w:shd[^/]*?w:fill="FFF2CC"[^/]*?)\s*w:themeFill="[^"]*"',
        '$1')
    $b = [regex]::Replace($b,
        '(<w:shd[^/]*?w:fill="FFF2CC"[^/]*?)\s*w:themeFillTint="[^"]*"',
        '$1')
    # Replace the entire <w:tblCellMar> block (or insert it) with our canonical values.
    $cellMar = '<w:tblCellMar><w:top w:w="14" w:type="dxa"/><w:left w:w="58" w:type="dxa"/><w:bottom w:w="14" w:type="dxa"/><w:right w:w="29" w:type="dxa"/></w:tblCellMar>'
    if ($b -match '<w:tblCellMar>.*?</w:tblCellMar>') {
        $b = [regex]::Replace($b, '<w:tblCellMar>.*?</w:tblCellMar>', $cellMar, 1)
    } elseif ($b -match '</w:tblPr>') {
        $b = [regex]::Replace($b, '</w:tblPr>', "$cellMar</w:tblPr>", 1)
    }

    # Inject lastRow yellow if missing. Some chapter sources never had it.
    if ($b -notmatch '<w:tblStylePr\s+w:type="lastRow"') {
        $lastRow = '<w:tblStylePr w:type="lastRow"><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="single" w:color="808080" w:themeColor="background1" w:themeShade="80" w:sz="4" w:space="0"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:tcPr></w:tblStylePr>'
        $b = [regex]::Replace($b, '</w:style>', "$lastRow</w:style>", 1)
    }

    # Force white text on the firstRow header (dark blue background needs it).
    $b = [regex]::Replace($b, '(<w:tblStylePr\s+w:type="firstRow">.*?<w:rPr>)(?!.*?<w:color\s)',
        '$1<w:color w:val="FFFFFF"/>', 1)

    # First column shading: F2F2F2 (light gray), scoped to the firstCol
    # conditional format so band1Horz striping is untouched. Fixed hex, no
    # themeFill, so it survives Import Styles into a differently-themed doc.
    $firstColShd = '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
    if ($b -match '<w:tblStylePr\s+w:type="firstCol">') {
        $b = [regex]::Replace($b,
            '(?s)(<w:tblStylePr\s+w:type="firstCol">.*?</w:tblStylePr>)',
            {
                param($m)
                $blk = $m.Groups[1].Value
                if ($blk -match '<w:shd\b[^>]*/>') {
                    [regex]::Replace($blk, '<w:shd\b[^>]*/>', $firstColShd, 1)
                } elseif ($blk -match '</w:tcPr>') {
                    [regex]::Replace($blk, '</w:tcPr>', "$firstColShd</w:tcPr>", 1)
                } else {
                    [regex]::Replace($blk, '<w:tblPr\s*/>', "<w:tblPr/><w:tcPr>$firstColShd</w:tcPr>", 1)
                }
            }, 1)
    } else {
        # No firstCol conditional format present: add one.
        $firstCol = '<w:tblStylePr w:type="firstCol"><w:tblPr/><w:tcPr>' + $firstColShd + '</w:tcPr></w:tblStylePr>'
        $b = [regex]::Replace($b, '</w:style>', "$firstCol</w:style>", 1)
    }

    # All rules the same color: unify every table border (the faint whole-table
    # grid and the banding / header / column rules) to a single hard-coded
    # A6A6A6, and drop themeColor/themeShade so the rule color cannot drift on
    # Import Styles. Only border elements are touched - they carry
    # w:val="single" + w:color; the cell-margin <w:top>/<w:left> carry w:w and
    # are left alone.
    $b = [regex]::Replace($b,
        '<w:(top|left|bottom|right|insideH|insideV)\b[^>]*?w:color="(?:[0-9A-Fa-f]{6}|auto)"[^>]*?/>',
        {
            param($m)
            $el = $m.Value
            $el = [regex]::Replace($el, 'w:color="(?:[0-9A-Fa-f]{6}|auto)"', 'w:color="A6A6A6"')
            $el = [regex]::Replace($el, '\s*w:themeColor="[^"]*"', '')
            $el = [regex]::Replace($el, '\s*w:themeShade="[^"]*"', '')
            $el = [regex]::Replace($el, '\s*w:themeTint="[^"]*"', '')
            $el
        })

    return $b
}

# $ListSingleSpacingEdit is provided by the OoXml module.

# xxtable: 9pt (sz=18, szCs=18), 1pt before, 1pt after, line spacing 0.9 multiple.
# Children that don't override inherit these values.
$xxtableEdit = {
    param($block)
    $newSpacing = 'w:before="20" w:after="20" w:line="216" w:lineRule="auto"'

    # Spacing: replace any existing or insert at end of pPr
    if ($block -match '<w:spacing\b[^/]*?/>') {
        $b = [regex]::Replace($block, '<w:spacing\b[^/]*?/>', "<w:spacing $newSpacing/>", 1)
    } elseif ($block -match '</w:pPr>') {
        $b = [regex]::Replace($block, '</w:pPr>', "<w:spacing $newSpacing/></w:pPr>", 1)
    } else {
        $b = [regex]::Replace($block, '</w:style>', "<w:pPr><w:spacing $newSpacing/></w:pPr></w:style>", 1)
    }

    # Size: set sz=18 (9pt)
    if ($b -match '<w:sz\s+w:val="[^"]*"\s*/>') {
        $b = [regex]::Replace($b, '<w:sz\s+w:val="[^"]*"\s*/>', '<w:sz w:val="18"/>', 1)
    } elseif ($b -match '<w:rPr>') {
        $b = [regex]::Replace($b, '<w:rPr>', '<w:rPr><w:sz w:val="18"/>', 1)
    } else {
        $b = [regex]::Replace($b, '</w:style>', '<w:rPr><w:sz w:val="18"/></w:rPr></w:style>', 1)
    }
    # Size: set szCs=18 (complex script)
    if ($b -match '<w:szCs\s+w:val="[^"]*"\s*/>') {
        $b = [regex]::Replace($b, '<w:szCs\s+w:val="[^"]*"\s*/>', '<w:szCs w:val="18"/>', 1)
    } elseif ($b -match '<w:sz\s+w:val="18"\s*/>') {
        $b = [regex]::Replace($b, '(<w:sz\s+w:val="18"\s*/>)', '$1<w:szCs w:val="18"/>', 1)
    }
    return $b
}

# TableText: line spacing multiple 0.9 + space-after 2pt (40 twips).
$tableTextSpacingEdit = {
    param($block)
    $newSpacing = 'w:after="40" w:line="216" w:lineRule="auto"'
    if ($block -match '<w:spacing\b[^/]*?/>') {
        return [regex]::Replace($block,
            '<w:spacing\b([^/]*?)/>',
            {
                param($m)
                $attrs = $m.Groups[1].Value
                $attrs = [regex]::Replace($attrs, '\s*w:line="[^"]*"', '')
                $attrs = [regex]::Replace($attrs, '\s*w:lineRule="[^"]*"', '')
                $attrs = [regex]::Replace($attrs, '\s*w:after="[^"]*"', '')
                # Drop the "Auto" spacing flags - afterAutospacing="1" would
                # override the explicit w:after and show as Auto after, not 2pt.
                $attrs = [regex]::Replace($attrs, '\s*w:afterAutospacing="[^"]*"', '')
                $attrs = [regex]::Replace($attrs, '\s*w:beforeAutospacing="[^"]*"', '')
                "<w:spacing$attrs $newSpacing/>"
            }, 1)
    }
    if ($block -match '</w:pPr>') {
        return [regex]::Replace($block, '</w:pPr>', "<w:spacing $newSpacing/></w:pPr>", 1)
    }
    return [regex]::Replace($block, '</w:style>', "<w:pPr><w:spacing $newSpacing/></w:pPr></w:style>", 1)
}

# Header: Source Sans Pro Light, 9pt (sz=18), no bold. (Color #00257A is
# applied by the setColor pass.)
# Edit-RunRPrInStyle is provided by the OoXml module.
$headerFontEdit = {
    param($block)
    return Edit-RunRPrInStyle -Block $block -NewFont 'Source Sans Pro Light' -SizeHalfPts '18'
}

# Caption: Source Sans Pro Semibold, 10pt (sz=20), no bold, +0.2pt tracking.
# Caption_Table and
# Caption_Figure inherit all three. The weight comes from the Semibold face, so
# w:b is written explicitly OFF: w:b is a toggle property, and left absent it
# would inherit the bold from xxheader and stack synthesized weight on top of a
# face that is already semibold.
# (Color #00257A is applied by the setColor pass.)
#
# Tracking: +0.2pt. Source Sans Pro is a tight humanist sans and sets cramped at
# caption size, especially in the Semibold face. w:spacing inside rPr is the
# character-level tracking Word exposes as Font > Advanced > Spacing: Expanded,
# and is unrelated to the w:spacing inside pPr that sets before/after/line. The
# value is twentieths of a point, so 4 = 0.2pt. Caption_Table and Caption_Figure
# inherit it from Caption.
$captionFontEdit = {
    param($block)
    $b = Edit-RunRPrInStyle -Block $block -NewFont 'Source Sans Pro Semibold' -SizeHalfPts '20'
    $track = '<w:spacing w:val="4"/>'
    # In rPr the child order runs ... color, spacing, w, kern, position, sz ...
    # so tracking goes in before w:sz, never after it.
    $pPrCloseIdx = $b.IndexOf('</w:pPr>')
    if ($pPrCloseIdx -ge 0) {
        $cut = $pPrCloseIdx + '</w:pPr>'.Length
        $prefix = $b.Substring(0, $cut); $tail = $b.Substring($cut)
    } else {
        $prefix = ''; $tail = $b
    }
    $tail = [regex]::Replace($tail, '<w:spacing w:val="-?\d+"\s*/>', '')
    if ($tail -match '<w:sz\b[^/]*?/>') {
        $tail = [regex]::Replace($tail, '(<w:sz\b[^/]*?/>)', "$track`$1", 1)
    } elseif ($tail -match '<w:rPr>') {
        $tail = [regex]::Replace($tail, '</w:rPr>', "$track</w:rPr>", 1)
    }
    return $prefix + $tail
}

# Footer: Source Sans Pro Light, 9pt (sz=18), no bold.
$footerFontEdit = {
    param($block)
    return Edit-RunRPrInStyle -Block $block -NewFont 'Source Sans Pro Light' -SizeHalfPts '18'
}

# xxbase: 10pt (sz=20 half-points). Children that don't override their own
# size inherit this through the basedOn chain.
$xxbaseSizeEdit = {
    param($block)
    if ($block -match '<w:sz\s+w:val="[^"]*"\s*/>') {
        $b = [regex]::Replace($block, '<w:sz\s+w:val="[^"]*"\s*/>', '<w:sz w:val="20"/>', 1)
    } elseif ($block -match '<w:rPr>') {
        $b = [regex]::Replace($block, '<w:rPr>', '<w:rPr><w:sz w:val="20"/>', 1)
    } else {
        $b = [regex]::Replace($block, '</w:style>', '<w:rPr><w:sz w:val="20"/></w:rPr></w:style>', 1)
    }
    if ($b -match '<w:szCs\s+w:val="[^"]*"\s*/>') {
        $b = [regex]::Replace($b, '<w:szCs\s+w:val="[^"]*"\s*/>', '<w:szCs w:val="20"/>', 1)
    } elseif ($b -match '<w:sz\s+w:val="20"\s*/>') {
        $b = [regex]::Replace($b, '(<w:sz\s+w:val="20"\s*/>)', '$1<w:szCs w:val="20"/>', 1)
    }
    return $b
}

# xxbase: 0 before / 6pt after (0/120 twips), line spacing 1.1 multiple (264).
#
# Space-after only, with before at zero. Word takes the MAXIMUM of the first
# paragraph's after, the second's before, and the inter-line spacing - it never
# adds them - so setting both sides means the smaller value is simply invisible
# and the numbers in the style stop describing what lands on the page.
#
# Zero before is also what makes heading space-after work. The gap under a
# heading is max(heading after, body before), so a non-zero body before swallows
# any heading after smaller than it, which is what forced the heading after
# values up to 6-10pt and inverted them against their own before values. Every
# style set Microsoft ships does the same thing: after on the body, before on
# the headings, nothing set on both.
$xxbaseSpacingEdit = {
    param($block)
    $newSpacing = 'w:before="0" w:after="120" w:line="264" w:lineRule="auto"'
    if ($block -match '<w:spacing\b[^/]*?/>') {
        return [regex]::Replace($block, '<w:spacing\b[^/]*?/>', "<w:spacing $newSpacing/>", 1)
    }
    if ($block -match '</w:pPr>') {
        return [regex]::Replace($block, '</w:pPr>', "<w:spacing $newSpacing/></w:pPr>", 1)
    }
    return [regex]::Replace($block, '</w:style>', "<w:pPr><w:spacing $newSpacing/></w:pPr></w:style>", 1)
}

# $BodyTextSpacingEdit is provided by the OoXml module.
# $bodyTextEdit is a NIST alias for it.
$bodyTextEdit = $BodyTextSpacingEdit

# Block-level mutations: Caption_Figure_2 (Source Sans Pro 10pt, no bold,
# 4pt padding all around, light grey background,
# light grey paragraph background). Resolution Blue shading is too dark to
# read against, so the callout background is #F2F2F2 (Office "Background 1,
# Darker 5%") for a clearly visible but subtle band.
$captionFigure2Edit = {
    param($block)
    $b = Edit-RunRPrInStyle -Block $block -NewFont 'Source Sans Pro' -SizeHalfPts '20'
    # Border and Shading > Options: 4pt distance from text on all four sides, so
    # the grey band stands off the text instead of hugging it. The borders
    # themselves stay "none"; w:space is the padding, and Word needs the four
    # side elements present to carry it. pBdr precedes shd in pPr child order.
    $pad = '<w:pBdr>' +
           '<w:top w:val="none" w:sz="0" w:space="4" w:color="auto"/>' +
           '<w:left w:val="none" w:sz="0" w:space="4" w:color="auto"/>' +
           '<w:bottom w:val="none" w:sz="0" w:space="4" w:color="auto"/>' +
           '<w:right w:val="none" w:sz="0" w:space="4" w:color="auto"/>' +
           '</w:pBdr>'
    if ($b -match '(?s)<w:pBdr>.*?</w:pBdr>') {
        $b = [regex]::Replace($b, '(?s)<w:pBdr>.*?</w:pBdr>', $pad, 1)
    } elseif ($b -match '<w:pPr>') {
        $b = [regex]::Replace($b, '<w:pPr>', "<w:pPr>$pad", 1)
    }

    # Line spacing multiple 0.9 (216 = 0.9 x 240). It inherited 0.85 from the
    # header chain, which reads as cramped in the grey callout band.
    if ($b -match '<w:spacing\b[^/]*?/>') {
        $b = [regex]::Replace($b, '<w:spacing\b([^/]*?)/>',
            {
                param($m)
                $attrs = $m.Groups[1].Value
                $attrs = [regex]::Replace($attrs, '\s*w:line="[^"]*"', '')
                $attrs = [regex]::Replace($attrs, '\s*w:lineRule="[^"]*"', '')
                "<w:spacing$attrs w:line=`"216`" w:lineRule=`"auto`"/>"
            }, 1)
    } elseif ($b -match '(?s)<w:shd\b[^/]*?/>') {
        # spacing follows shd in pPr child order.
        $b = [regex]::Replace($b, '(<w:shd\b[^/]*?/>)', '$1<w:spacing w:line="216" w:lineRule="auto"/>', 1)
    } elseif ($b -match '</w:pPr>') {
        $b = [regex]::Replace($b, '</w:pPr>', '<w:spacing w:line="216" w:lineRule="auto"/></w:pPr>', 1)
    }
    # Paragraph shading: replace any existing shd inside pPr, or insert just
    # before </w:pPr>. Operate on the pPr region only so we never touch
    # run-level shading.
    $shd = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>'
    $closeIdx = $b.IndexOf('</w:pPr>')
    if ($closeIdx -ge 0) {
        $pPrRegion = $b.Substring(0, $closeIdx)
        $rest      = $b.Substring($closeIdx)
        if ($pPrRegion -match '<w:shd\s+[^/]*?/>') {
            $pPrRegion = [regex]::Replace($pPrRegion, '<w:shd\s+[^/]*?/>', $shd, 1)
        } else {
            $pPrRegion = $pPrRegion + $shd
        }
        $b = $pPrRegion + $rest
    } else {
        # No pPr in the block. Insert one after basedOn (or name, or opening tag).
        if ($b -match '<w:basedOn\s+w:val="[^"]*"\s*/>') {
            $b = [regex]::Replace($b, '(<w:basedOn\s+w:val="[^"]*"\s*/>)', "`$1<w:pPr>$shd</w:pPr>", 1)
        } elseif ($b -match '<w:name\s+w:val="[^"]*"\s*/>') {
            $b = [regex]::Replace($b, '(<w:name\s+w:val="[^"]*"\s*/>)', "`$1<w:pPr>$shd</w:pPr>", 1)
        } else {
            $b = [regex]::Replace($b, '</w:style>', "<w:pPr>$shd</w:pPr></w:style>", 1)
        }
    }
    return $b
}

# Header Odd and Header Even style definitions (created if absent)
# Inherit from Header (Source Sans Pro Light 9pt, no bold). HeaderEven left-aligned
# for chapter number on outer (left) edge of even pages. HeaderOdd right-aligned for
# chapter title on outer (right) edge of odd pages. Both carry the bottom border in
# Resolution Blue, so the header paragraphs only carry the STYLEREF field.
$headerOddBlock = @'
<w:style w:type="paragraph" w:customStyle="1" w:styleId="HeaderOdd">
  <w:name w:val="Header Odd"/>
  <w:basedOn w:val="Header"/>
  <w:qFormat/>
  <w:pPr>
    <w:pBdr>
      <w:bottom w:val="single" w:sz="6" w:space="4" w:color="00257A"/>
    </w:pBdr>
    <w:jc w:val="right"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Source Sans Pro Light" w:hAnsi="Source Sans Pro Light" w:cs="Source Sans Pro Light"/>
    <w:b w:val="0"/>
    <w:bCs w:val="0"/>
    <w:sz w:val="18"/>
    <w:szCs w:val="18"/>
    <w14:ligatures w14:val="none"/>
  </w:rPr>
</w:style>
'@

# Headerch: the chapter opener. Same face as Header but with NO rule
# underneath, for the banner table of color blocks that sits immediately before
# the first Heading 1. That cell uses Header today and picks up its horizontal
# rule, which does not belong inside the banner. Applied by hand, not by any
# script: this only guarantees the style exists and is defined correctly.
# The bottom border is written explicitly as "none" rather than left out, so it
# overrides what Header carries.
#
# Name and styleId are both "Headerch", with no underscore. Word derives the id
# from the NAME by stripping non-alphanumerics, so a name like "Header_ch" would
# yield the id "Headerch" and Word would silently rewrite any other casing on
# first save, repointing every reference and leaving the style looking like it
# had vanished. Keeping the two identical sidesteps that.
$headerChBlock = @'
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Headerch">
  <w:name w:val="Headerch"/>
  <w:basedOn w:val="Header"/>
  <w:qFormat/>
  <w:pPr>
    <w:pBdr>
      <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:pBdr>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Source Sans Pro Light" w:hAnsi="Source Sans Pro Light" w:cs="Source Sans Pro Light"/>
    <w:b w:val="0"/>
    <w:bCs w:val="0"/>
    <w:sz w:val="18"/>
    <w:szCs w:val="18"/>
    <w14:ligatures w14:val="none"/>
  </w:rPr>
</w:style>
'@

$headerEvenBlock = @'
<w:style w:type="paragraph" w:customStyle="1" w:styleId="HeaderEven">
  <w:name w:val="Header Even"/>
  <w:basedOn w:val="Header"/>
  <w:qFormat/>
  <w:pPr>
    <w:pBdr>
      <w:bottom w:val="single" w:sz="6" w:space="4" w:color="00257A"/>
    </w:pBdr>
    <w:jc w:val="left"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Source Sans Pro Light" w:hAnsi="Source Sans Pro Light" w:cs="Source Sans Pro Light"/>
    <w:b w:val="0"/>
    <w:bCs w:val="0"/>
    <w:sz w:val="18"/>
    <w:szCs w:val="18"/>
    <w14:ligatures w14:val="none"/>
  </w:rPr>
</w:style>
'@

# Footer Odd and Footer Even style definitions (created if absent)
# 9pt Source Sans Pro Light, tabs at center (4680 twips = 3.25") and right (9360 twips = 6.5").
# The two share the same definition; the difference between odd and even page output is in
# the footer*.xml content (tab placements for page-number left vs right), not in the style.
# Both carry the top border in Resolution Blue.
$footerOddBlock = @'
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FooterOdd">
  <w:name w:val="Footer Odd"/>
  <w:basedOn w:val="Footer"/>
  <w:qFormat/>
  <w:pPr>
    <w:pBdr>
      <w:top w:val="single" w:sz="6" w:space="4" w:color="00257A"/>
    </w:pBdr>
    <w:tabs>
      <w:tab w:val="clear" w:pos="4680"/>
      <w:tab w:val="clear" w:pos="9360"/>
      <w:tab w:val="right" w:pos="9360"/>
    </w:tabs>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Source Sans Pro Light" w:hAnsi="Source Sans Pro Light"/>
    <w:b w:val="0"/>
    <w:bCs w:val="0"/>
    <w:sz w:val="18"/>
    <w:szCs w:val="18"/>
    <w14:ligatures w14:val="none"/>
  </w:rPr>
</w:style>
'@

$footerEvenBlock = @'
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FooterEven">
  <w:name w:val="Footer Even"/>
  <w:basedOn w:val="Footer"/>
  <w:qFormat/>
  <w:pPr>
    <w:pBdr>
      <w:top w:val="single" w:sz="6" w:space="4" w:color="00257A"/>
    </w:pBdr>
    <w:tabs>
      <w:tab w:val="clear" w:pos="4680"/>
      <w:tab w:val="clear" w:pos="9360"/>
      <w:tab w:val="right" w:pos="9360"/>
    </w:tabs>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Source Sans Pro Light" w:hAnsi="Source Sans Pro Light"/>
    <w:b w:val="0"/>
    <w:bCs w:val="0"/>
    <w:sz w:val="18"/>
    <w:szCs w:val="18"/>
    <w14:ligatures w14:val="none"/>
  </w:rPr>
</w:style>
'@

# ---------------------------------------------------------------------------
# Build target list
# ---------------------------------------------------------------------------

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|dotx|docm|dotm)$' } |
        ForEach-Object { $_.FullName }
}

if ($targets.Count -eq 0) { throw "No supported Word files found." }

Write-Host "Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "  $_" }
if ($Clean) {
    Write-Host "  [Clean mode] applying architecture only; skipping legacy-fix passes" -ForegroundColor Cyan
    Write-Host "  (cs='Times New Roman' strip and junk-style deletions)." -ForegroundColor Cyan
}
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Helpers: provided by the OoXml module (Edit-StyleBlock, Remove-StyleBlock,
# Get-AllStyleIdsByType, Set-StyleBasedOn, Set-StyleColor, Edit-RunRPrInStyle,
# Set-StyleRecommended, Set-StyleHidden, Get-HeadingAbstractNumIds,
# Set-NumberingSuffix, Sync-ListIndents, Set-BulletGlyph,
# $ListSingleSpacingEdit, $BodyTextSpacingEdit, $SuppressNumbering)
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Process each target
# ---------------------------------------------------------------------------

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    # Read styles.xml AND numbering.xml (numbering edits land in a separate pass)
    $xml = $null; $numberingXml = $null
    $stream = [System.IO.File]::Open($target, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $entry = $zip.GetEntry('word/styles.xml')
            if (-not $entry) { Write-Warning "  No word/styles.xml. Skipping."; continue }
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }

            $numEntry = $zip.GetEntry('word/numbering.xml')
            if ($numEntry) {
                $r2 = New-Object System.IO.StreamReader($numEntry.Open())
                try { $numberingXml = $r2.ReadToEnd() } finally { $r2.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }

    # Normalize the 02257A typo to canonical Resolution Blue 00257A.
    # Originated in earlier toolkit revisions; scrub on every run.
    $beforeCount = ([regex]::Matches($xml, '02257A')).Count
    if ($beforeCount -gt 0) {
        $xml = $xml -replace '02257A', '00257A'
        Write-Host "  Normalized $beforeCount occurrence(s) of 02257A -> 00257A" -ForegroundColor DarkGray
    }

    # Apply rebase
    $rebaseHits = 0
    foreach ($id in $rebase.Keys) {
        $newBase = $rebase[$id]
        $result = Edit-StyleBlock -Xml $xml -Id $id -Mutator { param($block, $attrs, $body) Set-StyleBasedOn -Block $block -NewId $newBase }
        $xml = $result[0]
        if ($result[1]) { $rebaseHits++ }
    }
    Write-Host "  Rebased: $rebaseHits style(s)" -ForegroundColor DarkGray

    # AppendixHead 1-4 and Frontmatterhead 1-4 must render exactly like
    # Heading 1-4. The rebase above points each at its matching Heading; now
    # clear each one's own paragraph and run formatting (spacing, font, size
    # overrides) so nothing overrides the inherited look. Color is re-applied by
    # the setColor pass just below, and numbering suppression (numId=0) is added
    # afterward.
    $headMatchStrip = {
        param($block)
        $b = [regex]::Replace($block, '(?s)<w:pPr>.*?</w:pPr>', '')
        $b = [regex]::Replace($b, '<w:pPr\s*/>', '')
        $b = [regex]::Replace($b, '(?s)<w:rPr>.*?</w:rPr>', '')
        $b = [regex]::Replace($b, '<w:rPr\s*/>', '')
        return $b
    }
    foreach ($ahId in @('AppendixHead','AppendixHead2','AppendixHead3','AppendixHead4',
                        'Frontmatterhead1','Frontmatterhead2','Frontmatterhead3','Frontmatterhead4')) {
        $result = Edit-StyleBlock -Xml $xml -Id $ahId -Mutator $headMatchStrip
        $xml = $result[0]
        if ($result[1]) { Write-Host "  ${ahId}: cleared own formatting to inherit its Heading" -ForegroundColor DarkGray }
    }

    # Apply color sets
    $colorHits = 0
    foreach ($id in $setColor.Keys) {
        $hex = $setColor[$id]
        $result = Edit-StyleBlock -Xml $xml -Id $id -Mutator { param($block, $attrs, $body) Set-StyleColor -Block $block -HexNoHash $hex }
        $xml = $result[0]
        if ($result[1]) { $colorHits++ }
    }
    Write-Host "  Color set on: $colorHits style(s)" -ForegroundColor DarkGray

    # Heading 1 and Heading 3 carry a stray <w:rFonts ... w:cs="Times New Roman"/>
    # that the other headings (2, 4-9) do not. It only affects complex-script
    # (Arabic/Hebrew) runs, so it has no visible effect on Latin content, but it
    # makes the nine heading styles inconsistent. Drop the w:cs override so they
    # all resolve fonts through the theme the same way. (Legacy-fix: -Clean skips.)
    if (-not $Clean) {
        $csHits = 0
        foreach ($id in @('Heading1','Heading3')) {
            $csMut = {
                param($block)
                [regex]::Replace($block, '(<w:rFonts\b[^>]*?)\s+w:cs="Times New Roman"([^>]*/>)', '$1$2')
            }
            $res = Edit-StyleBlock -Xml $xml -Id $id -Mutator $csMut
            if ($res[0] -ne $xml) { $csHits++ }
            $xml = $res[0]
        }
        if ($csHits) { Write-Host "  Removed stray w:cs='Times New Roman' from $csHits heading(s)" -ForegroundColor DarkGray }
    }

    # Leaf styles (Versohead, Title/TitleChar, MTDisplayEquation/*Char) are defined
    # wholesale as canonical blocks below, not patched here - this replaces the old
    # per-property regex for Black->Bold, Merriweather, and Title-bold. See the
    # "$leafStyles" create-or-replace pass near the Header/Footer definitions.

    # Heading4: include in the heading numbering at level 4 (ilvl=3), using the
    # same numId Heading 1-3 use, so it renders "1.2.3.4". For an UNnumbered
    # level-4 heading use the separate "Heading 4 Sub-Head" style, which is left
    # alone. The heading list already defines ilvl=3 as "%1.%2.%3.%4.".
    $h3block = [regex]::Match($xml, '(?s)<w:style\b[^>]*w:styleId="Heading3"[^>]*>.*?</w:style>').Value
    $hnid = ([regex]::Match($h3block, '<w:numId\s+w:val="(\d+)"')).Groups[1].Value
    if (-not $hnid) {
        $h1block = [regex]::Match($xml, '(?s)<w:style\b[^>]*w:styleId="Heading1"[^>]*>.*?</w:style>').Value
        $hnid = ([regex]::Match($h1block, '<w:numId\s+w:val="(\d+)"')).Groups[1].Value
    }
    if ($hnid) {
        $h4np = '<w:numPr><w:ilvl w:val="3"/><w:numId w:val="' + $hnid + '"/></w:numPr>'
        $h4mut = {
            param($block)
            # numPr goes after keepNext/keepLines and before spacing/outlineLvl (schema order).
            if ($block -match '(?s)<w:numPr>.*?</w:numPr>') { return [regex]::Replace($block, '(?s)<w:numPr>.*?</w:numPr>', $h4np, 1) }
            if ($block -match '<w:spacing\b')    { return [regex]::Replace($block, '(<w:spacing\b)', "$h4np`$1", 1) }
            if ($block -match '<w:outlineLvl\b') { return [regex]::Replace($block, '(<w:outlineLvl\b)', "$h4np`$1", 1) }
            if ($block -match '</w:pPr>')        { return [regex]::Replace($block, '</w:pPr>', "$h4np</w:pPr>", 1) }
            if ($block -match '<w:pPr\s*/>')     { return [regex]::Replace($block, '<w:pPr\s*/>', "<w:pPr>$h4np</w:pPr>", 1) }
            if ($block -match '<w:basedOn\s+w:val="[^"]*"\s*/>') { return [regex]::Replace($block, '(<w:basedOn\s+w:val="[^"]*"\s*/>)', "`$1<w:pPr>$h4np</w:pPr>", 1) }
            return [regex]::Replace($block, '</w:style>', "<w:pPr>$h4np</w:pPr></w:style>", 1)
        }.GetNewClosure()
        $h4 = Edit-StyleBlock -Xml $xml -Id 'Heading4' -Mutator $h4mut
        $xml = $h4[0]
        if ($h4[1]) { Write-Host "  Heading4: numbered at level 4 (ilvl=3, numId=$hnid)" -ForegroundColor DarkGray }
    } else {
        Write-Warning "  Heading4: heading numId not found; numbering left as-is."
    }

    # AppendixHead and Frontmatterhead families: each is now basedOn its matching
    # Heading (1-4), and Heading 1-3 carry a chapter numId, so the inherited
    # auto-number would leak into appendix and front-matter titles (a Table of
    # Contents must not render as "Chapter 1"). Install the explicit no-numbering
    # marker <w:numPr><w:numId w:val="0"/></w:numPr> on every level so the
    # suppression holds regardless of which Heading it inherits from.
    # $SuppressNumbering is provided by the OoXml module.
    foreach ($ahId in @('AppendixHead','AppendixHead2','AppendixHead3','AppendixHead4',
                        'Appendix_Head','Appendix_Head2','Appendix_Head3','Appendix_Head4',
                        'Frontmatterhead1','Frontmatterhead2','Frontmatterhead3','Frontmatterhead4')) {
        $ah = Edit-StyleBlock -Xml $xml -Id $ahId -Mutator $SuppressNumbering
        $xml = $ah[0]
        if ($ah[1]) { Write-Host "  ${ahId}: numbering suppressed (numId=0)" -ForegroundColor DarkGray }
    }

    # NIST-Data: install canonical (overwrite any existing block, else insert before </w:styles>)
    if ($xml -match '<w:style[^>]*w:styleId="NIST-Data"[^>]*>.*?</w:style>') {
        $xml = [regex]::Replace($xml, '(?s)<w:style[^>]*w:styleId="NIST-Data"[^>]*>.*?</w:style>', { param($m) $NIST_DATA_CANONICAL }, 1)
        Write-Host "  NIST-Data: reinstalled" -ForegroundColor DarkGray
    } else {
        $xml = [regex]::Replace($xml, '</w:styles>', { param($m) $NIST_DATA_CANONICAL + '</w:styles>' }, 1)
        Write-Host "  NIST-Data: installed" -ForegroundColor DarkGray
    }

    # NIST-Data: ensure lsdException so Word shows it in the Table Styles gallery
    $lsdNist = '<w:lsdException w:name="NIST-Data" w:uiPriority="52" w:semiHidden="0" w:unhideWhenUsed="0" w:qFormat="1"/>'
    $latentMatch = [regex]::Match($xml, '<w:latentStyles\b.*?</w:latentStyles>', 'Singleline')
    if ($latentMatch.Success -and $latentMatch.Value -notmatch 'w:name="NIST-Data"') {
        $xml = [regex]::Replace($xml, '</w:latentStyles>', "$lsdNist</w:latentStyles>", 1)
        # Bump the count attribute on <w:latentStyles> by 1
        $xml = [regex]::Replace($xml, '(<w:latentStyles\b[^>]*?w:count=")(\d+)(")',
            { param($m) "$($m.Groups[1].Value)$([int]$m.Groups[2].Value + 1)$($m.Groups[3].Value)" }, 1)
        Write-Host "  NIST-Data: lsdException added (gallery visibility)" -ForegroundColor DarkGray
    }

    # Sidebar table: install canonical (overwrite any existing block, else insert before </w:styles>)
    if ($xml -match '<w:style[^>]*w:styleId="SidebarTable"[^>]*>.*?</w:style>') {
        $xml = [regex]::Replace($xml, '(?s)<w:style[^>]*w:styleId="SidebarTable"[^>]*>.*?</w:style>', { param($m) $SIDEBAR_TABLE_CANONICAL }, 1)
        Write-Host "  SidebarTable: reinstalled" -ForegroundColor DarkGray
    } else {
        $xml = [regex]::Replace($xml, '</w:styles>', { param($m) $SIDEBAR_TABLE_CANONICAL + '</w:styles>' }, 1)
        Write-Host "  SidebarTable: installed (Sidebar Table, amber FFF2CC fill)" -ForegroundColor DarkGray
    }

    # Sidebar table: ensure lsdException so Word shows it in the Table Styles gallery
    $lsdSidebar = '<w:lsdException w:name="Sidebar Table" w:uiPriority="53" w:semiHidden="0" w:unhideWhenUsed="0" w:qFormat="1"/>'
    $latentMatchSb = [regex]::Match($xml, '<w:latentStyles\b.*?</w:latentStyles>', 'Singleline')
    if ($latentMatchSb.Success -and $latentMatchSb.Value -notmatch 'w:name="Sidebar Table"') {
        $xml = [regex]::Replace($xml, '</w:latentStyles>', "$lsdSidebar</w:latentStyles>", 1)
        $xml = [regex]::Replace($xml, '(<w:latentStyles\b[^>]*?w:count=")(\d+)(")',
            { param($m) "$($m.Groups[1].Value)$([int]$m.Groups[2].Value + 1)$($m.Groups[3].Value)" }, 1)
        Write-Host "  SidebarTable: lsdException added (gallery visibility)" -ForegroundColor DarkGray
    }

    # HM-Data: lock yellow, set canonical cell margins
    $hm = Edit-StyleBlock -Xml $xml -Id 'HM-Data' -Mutator $hmDataEdit
    $xml = $hm[0]
    if ($hm[1]) { Write-Host "  HM-Data: locked lastRow yellow + set cell margins" -ForegroundColor DarkGray }

    # Default paragraph style inside HM-Data and HM-Cover = TableText
    foreach ($tid in @('HM-Data','HM-Cover')) {
        $r = Edit-StyleBlock -Xml $xml -Id $tid -Mutator $tableUseTableText
        $xml = $r[0]
        if ($r[1]) { Write-Host "  ${tid}: default paragraph style set to TableText" -ForegroundColor DarkGray }
    }

    # TableText: ensure next-paragraph is TableText (so Enter in a TableText paragraph stays in TableText)
    $tt = Edit-StyleBlock -Xml $xml -Id 'TableText' -Mutator $tableTextEdit
    $xml = $tt[0]
    if ($tt[1]) { Write-Host "  TableText: next style set to TableText" -ForegroundColor DarkGray }

    # xxbase: 10pt
    $xb = Edit-StyleBlock -Xml $xml -Id 'xxbase' -Mutator $xxbaseSizeEdit
    $xml = $xb[0]
    if ($xb[1]) { Write-Host "  xxbase: size set to 10pt (sz=20)" -ForegroundColor DarkGray }

    # xxbase: 4pt before / 4pt after spacing (BodyText and the list family inherit through here)
    $xbs = Edit-StyleBlock -Xml $xml -Id 'xxbase' -Mutator $xxbaseSpacingEdit
    $xml = $xbs[0]
    if ($xbs[1]) { Write-Host "  xxbase: spacing set to 4pt before / 4pt after (line 1.1)" -ForegroundColor DarkGray }

    # header: Source Sans Pro Light, 9pt, no bold
    $hd = Edit-StyleBlock -Xml $xml -Id 'header' -Mutator $headerFontEdit
    $xml = $hd[0]
    if ($hd[1]) { Write-Host "  header: Source Sans Pro Light 9pt, no bold" -ForegroundColor DarkGray }

    # footer: Source Sans Pro Light, 9pt, no bold
    $ft = Edit-StyleBlock -Xml $xml -Id 'footer' -Mutator $footerFontEdit
    $xml = $ft[0]
    if ($ft[1]) { Write-Host "  footer: Source Sans Pro Light 9pt, no bold" -ForegroundColor DarkGray }

    # xxtable: 9pt + 1pt before/after + line spacing 0.9 multiple
    $xt = Edit-StyleBlock -Xml $xml -Id 'xxtable' -Mutator $xxtableEdit
    $xml = $xt[0]
    if ($xt[1]) { Write-Host "  xxtable: 9pt, 1pt before/after, line 0.9" -ForegroundColor DarkGray }

    # TableText: line spacing multiple 0.9 + space-after 2pt
    $ts = Edit-StyleBlock -Xml $xml -Id 'TableText' -Mutator $tableTextSpacingEdit
    $xml = $ts[0]
    if ($ts[1]) { Write-Host "  TableText: line spacing multiple 0.9 + space-after 2pt" -ForegroundColor DarkGray }

    # _graphic: remove keep-with-next and keep-lines-together, and drop the
    # basedOn (was BodyText) so the style stands on its own. Its keep chain was
    # carried by its own elements, so stripping them turns both off.
    $graphicEdit = {
        param($block)
        $b = [regex]::Replace($block, '<w:keepNext\b[^>]*/>', '')
        $b = [regex]::Replace($b, '<w:keepLines\b[^>]*/>', '')
        $b = [regex]::Replace($b, '<w:basedOn\b[^>]*/>', '')
        return $b
    }
    $gk = Edit-StyleBlock -Xml $xml -Id 'graphic' -Mutator $graphicEdit
    $xml = $gk[0]
    if ($gk[1]) { Write-Host "  _graphic: keep-next/keep-lines removed, basedOn dropped" -ForegroundColor DarkGray }

    # Table of Figures: match TOC 1's paragraph spacing so the two lists read the
    # same. Both inherit line 0.85 from xxheader; TOC 1 tightens before/after to
    # 100 while Table of Figures otherwise inherits xxheader's 200. Copy TOC 1's
    # spacing element verbatim (resolved at run time) so they stay in agreement.
    $toc1sp = ([regex]::Match(
        [regex]::Match($xml, '(?s)<w:style\b[^>]*w:styleId="TOC1"[^>]*>.*?</w:style>').Value,
        '<w:spacing\b[^/]*/>')).Value
    if ($toc1sp) {
        $tofMut = {
            param($block)
            if ($block -match '<w:spacing\b[^/]*/>') { return [regex]::Replace($block, '<w:spacing\b[^/]*/>', $toc1sp, 1) }
            if ($block -match '<w:ind\b')            { return [regex]::Replace($block, '(<w:ind\b)', "$toc1sp`$1", 1) }
            if ($block -match '</w:pPr>')            { return [regex]::Replace($block, '</w:pPr>', "$toc1sp</w:pPr>", 1) }
            if ($block -match '<w:pPr\s*/>')         { return [regex]::Replace($block, '<w:pPr\s*/>', "<w:pPr>$toc1sp</w:pPr>", 1) }
            if ($block -match '<w:basedOn\s+w:val="[^"]*"\s*/>') { return [regex]::Replace($block, '(<w:basedOn\s+w:val="[^"]*"\s*/>)', "`$1<w:pPr>$toc1sp</w:pPr>", 1) }
            return [regex]::Replace($block, '</w:style>', "<w:pPr>$toc1sp</w:pPr></w:style>", 1)
        }.GetNewClosure()
        $tof = Edit-StyleBlock -Xml $xml -Id 'TableofFigures' -Mutator $tofMut
        $xml = $tof[0]
        if ($tof[1]) { Write-Host "  Table of Figures: spacing matched to TOC 1" -ForegroundColor DarkGray }
    }

    # Sidebar / Sidebar Bullet: install the callout paragraph styles if a chapter
    # is missing them (not every source chapter defines them). Both frameless -
    # the callout floats via the Sidebar Table style, so the paragraph text is
    # plain: Sidebar = sans table text (xxtable), Sidebar Bullet = Sidebar + a
    # bullet borrowed from List Bullet's numId.
    if ($xml -notmatch '<w:style[^>]*w:styleId="Sidebar"[^>]*>') {
        $sidebarCanon = '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Sidebar"><w:name w:val="Sidebar"/><w:basedOn w:val="xxtable"/><w:uiPriority w:val="93"/><w:qFormat/><w:pPr><w:spacing w:before="60" w:line="204" w:lineRule="auto"/></w:pPr></w:style>'
        $xml = [regex]::Replace($xml, '</w:styles>', "$sidebarCanon</w:styles>", 1)
        Write-Host "  Sidebar: created (was missing)" -ForegroundColor DarkGray
    }
    if ($xml -notmatch '<w:style[^>]*w:styleId="SidebarBullet"[^>]*>') {
        $lbBlock = [regex]::Match($xml, '(?s)<w:style\b[^>]*w:styleId="ListBullet"[^>]*>.*?</w:style>').Value
        $sbNumId = ([regex]::Match($lbBlock, '<w:numId\s+w:val="(\d+)"')).Groups[1].Value
        $bulletNp = if ($sbNumId) { '<w:numPr><w:numId w:val="' + $sbNumId + '"/></w:numPr>' } else { '' }
        $sbCanon = '<w:style w:type="paragraph" w:customStyle="1" w:styleId="SidebarBullet"><w:name w:val="Sidebar Bullet"/><w:basedOn w:val="Sidebar"/><w:uiPriority w:val="94"/><w:qFormat/><w:pPr>' + $bulletNp + '<w:ind w:left="240" w:hanging="180"/></w:pPr></w:style>'
        $xml = [regex]::Replace($xml, '</w:styles>', "$sbCanon</w:styles>", 1)
        Write-Host "  SidebarBullet: created (was missing)" -ForegroundColor DarkGray
    }

    # Turn OpenType ligatures OFF everywhere. Word's local "Save As PDF" can
    # mis-encode ligated glyphs (auto-numbering / curly quotes in Source Sans
    # Pro) as Cyrillic. Many styles carry <w14:ligatures w14:val="standardContextual"/>
    # (ON), which overrides any default, so flip every existing one to "none";
    # then set the docDefaults default to none for styles that don't set it.
    $ligOn = ([regex]::Matches($xml, '<w14:ligatures\b[^>]*w14:val="(?!none")[^"]*"[^>]*/>')).Count
    # 1. Flip every existing ligatures element (any val) to none.
    $xml = [regex]::Replace($xml, '(<w14:ligatures\b[^>]*w14:val=")[^"]*("[^>]*/>)', '${1}none$2')
    $lig = '<w14:ligatures w14:val="none"/>'
    # 2. Add an explicit ligatures=none to every rPr that does not already carry
    #    one, so the setting is present on every style's run properties rather
    #    than relying on docDefaults inheritance. w14:ligatures is the last child
    #    of rPr, so appending before </w:rPr> keeps schema order valid.
    $ligAdded = ([regex]::Matches($xml, '(?s)<w:rPr\b[^>]*?/>|<w:rPr\b[^>]*?>.*?</w:rPr>') |
                 Where-Object { $_.Value -notmatch '<w14:ligatures' }).Count
    $rprEval = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $w = $m.Value
        if ($w -match '<w14:ligatures') { return $w }
        if ($w -match '/>$') { return ($w -replace '/>$', ">$lig</w:rPr>") }
        return ($w -replace '</w:rPr>$', "$lig</w:rPr>")
    }.GetNewClosure()
    $xml = [System.Text.RegularExpressions.Regex]::Replace($xml, '(?s)<w:rPr\b[^>]*?/>|<w:rPr\b[^>]*?>.*?</w:rPr>', $rprEval)
    # 3. Guarantee docDefaults itself carries the setting (covers styles/runs
    #    that define no rPr at all and inherit from here).
    $rprDefault = [regex]::Match($xml, '(?s)<w:rPrDefault>.*?</w:rPrDefault>').Value
    if (-not $rprDefault -or $rprDefault -notmatch '<w14:ligatures') {
        if ($xml -match '<w:rPrDefault>\s*<w:rPr\b[^>]*>') {
            $xml = [regex]::Replace($xml, '(<w:rPrDefault>\s*<w:rPr\b[^>]*>)((?:(?!</w:rPr>).)*?)(</w:rPr>)', "`$1`$2$lig`$3")
        } elseif ($xml -match '<w:rPrDefault>\s*<w:rPr\s*/>') {
            $xml = [regex]::Replace($xml, '<w:rPrDefault>\s*<w:rPr\s*/>', "<w:rPrDefault><w:rPr>$lig</w:rPr>")
        } elseif ($xml -match '<w:rPrDefault\s*/>') {
            $xml = [regex]::Replace($xml, '<w:rPrDefault\s*/>', "<w:rPrDefault><w:rPr>$lig</w:rPr></w:rPrDefault>")
        } elseif ($xml -match '<w:docDefaults>') {
            $xml = [regex]::Replace($xml, '(<w:docDefaults>)', "`$1<w:rPrDefault><w:rPr>$lig</w:rPr></w:rPrDefault>")
        }
    }
    Write-Host "  Ligatures: none on every rPr ($ligOn flipped from on, $ligAdded added) + docDefaults" -ForegroundColor DarkGray

    # ---- Heading spacing ----
    # Space-before carries the gap, space-after stays small. Word takes the
    # MAXIMUM of the first paragraph's after, the second's before and the
    # inter-line spacing, so with body text at 0 before / 6pt after, a heading's
    # before must clear 6pt to open any space at all, and its after lands
    # unopposed because nothing below sets before. That is the pattern every
    # Microsoft style set ships, and it is why the old values (Heading 2 at 4pt
    # before / 6pt after) read as inverted: the heading sat closer to the text
    # above it than the text it introduced.
    #
    # before descends with level; after is small and near-flat.
    $headingSpacing = [ordered]@{
        'Heading1' = @{ Before = 360; After = 80 }   # 18pt / 4pt
        'Heading2' = @{ Before = 280; After = 60 }   # 14pt / 3pt
        'Heading3' = @{ Before = 240; After = 60 }   # 12pt / 3pt
        'Heading4' = @{ Before = 200; After = 40 }   # 10pt / 2pt
        'Heading5' = @{ Before = 180; After = 40 }   #  9pt / 2pt
        'Heading6' = @{ Before = 160; After = 40 }   #  8pt / 2pt
        # Front_matter_head4 doubles as the contributor sub-head (Participants,
        # Chapter 1, ...) in Acknowledgements and Author Chapter Credits, so it
        # keeps a slightly looser after than Heading 4.
        'Frontmatterhead4' = @{ Before = 200; After = 80 }   # 10pt / 4pt
    }
    foreach ($id in $headingSpacing.Keys) {
        $bef = $headingSpacing[$id].Before
        $aft = $headingSpacing[$id].After
        $newSp = "<w:spacing w:before=`"$bef`" w:after=`"$aft`"/>"
        $r = Edit-StyleBlock -Xml $xml -Id $id -Mutator {
            param($block)
            if ($block -match '<w:spacing\b[^/]*?/>') {
                return [regex]::Replace($block, '<w:spacing\b([^/]*?)/>', {
                    param($m)
                    $a = $m.Groups[1].Value
                    $a = [regex]::Replace($a, '\s*w:before="[^"]*"', '')
                    $a = [regex]::Replace($a, '\s*w:after="[^"]*"', '')
                    $a = [regex]::Replace($a, '\s*w:beforeAutospacing="[^"]*"', '')
                    $a = [regex]::Replace($a, '\s*w:afterAutospacing="[^"]*"', '')
                    "<w:spacing w:before=`"$bef`" w:after=`"$aft`"$a/>"
                }, 1)
            } elseif ($block -match '</w:pPr>') {
                return [regex]::Replace($block, '</w:pPr>', "$newSp</w:pPr>", 1)
            }
            return [regex]::Replace($block, '</w:style>', "<w:pPr>$newSp</w:pPr></w:style>", 1)
        }.GetNewClosure()
        $xml = $r[0]
        if ($r[1]) { Write-Host ("  {0}: before={1}pt after={2}pt" -f $id, ($bef/20), ($aft/20)) -ForegroundColor DarkGray }
    }

    # The Front_matter_head and Appendix_Head families are clones of Heading 1-4
    # and each carried its own spacing override, so they ignored the table above
    # and drifted. Stripping the override lets them inherit from their Heading
    # parent, which is the whole point of the basedOn chain.
    $cloneFamilies = @('Frontmatterhead1','Frontmatterhead2','Frontmatterhead3',
                       'AppendixHead','AppendixHead2','AppendixHead3','AppendixHead4',
                       'Heading4Sub-Head')
    $cloneHits = 0
    foreach ($id in $cloneFamilies) {
        $r = Edit-StyleBlock -Xml $xml -Id $id -Mutator {
            param($block)
            return [regex]::Replace($block, '<w:spacing\b[^/]*?/>', '')
        }
        $xml = $r[0]
        if ($r[1]) { $cloneHits++ }
    }
    Write-Host "  Front_matter_head / Appendix_Head / Heading4 Sub-Head: spacing override stripped on $cloneHits style(s), now inherit" -ForegroundColor DarkGray

    # Heading 4 Sub-Head is a Heading 4 without a number. numPr inherits through
    # basedOn, so leaving it out adopts Heading 4's numbering instead of removing
    # it - numId=0 is the documented "no numbering" sentinel. Eight of the ten
    # chapters were missing this and were numbering their sub-heads.
    $sub = Edit-StyleBlock -Xml $xml -Id 'Heading4Sub-Head' -Mutator $SuppressNumbering
    $xml = $sub[0]
    if ($sub[1]) { Write-Host "  Heading4 Sub-Head: numbering suppressed (numId=0)" -ForegroundColor DarkGray }

    # Heading 6: semibold italic, and grey via the color pass. It was bold blue
    # 12pt, identical to Heading 4.
    $h6 = Edit-StyleBlock -Xml $xml -Id 'Heading6' -Mutator {
        param($block)
        $b = Edit-RunRPrInStyle -Block $block -NewFont 'Source Sans Pro SemiBold' -SizeHalfPts '24'
        $pPrCloseIdx = $b.IndexOf('</w:pPr>')
        if ($pPrCloseIdx -ge 0) {
            $cut = $pPrCloseIdx + '</w:pPr>'.Length
            $prefix = $b.Substring(0, $cut); $tail = $b.Substring($cut)
        } else { $prefix = ''; $tail = $b }
        if ($tail -notmatch '<w:i\b') {
            $tail = [regex]::Replace($tail, '<w:rPr>', '<w:rPr><w:i/><w:iCs/>', 1)
        }
        return Sort-RunPropertyOrder -Block ($prefix + $tail)
    }
    $xml = $h6[0]
    if ($h6[1]) { Write-Host "  Heading6: Source Sans Pro SemiBold italic 12pt (grey via color pass)" -ForegroundColor DarkGray }

    # ---- Pagination: off by default, on where it earns its keep ----
    # Blanket keep-with-next and keep-lines fight a fixed page grid: they push
    # content to the next page and leave short pages behind. But turning them
    # off everywhere strands headings at the foot of a page and separates a
    # table caption from its table, which is worse. So: strip everything first,
    # then switch them back on only for the styles that need them.
    #
    # widowControl defaults to ON when absent, so leaving it stripped restores
    # Word's normal behaviour - a single stranded line is a real defect, not a
    # grid problem.
    $keepNextHits  = ([regex]::Matches($xml, '<w:keepNext\b[^>]*/>')).Count
    $keepLinesHits = ([regex]::Matches($xml, '<w:keepLines\b[^>]*/>')).Count
    $widowHits     = ([regex]::Matches($xml, '<w:widowControl\b[^>]*/>')).Count
    $xml = [regex]::Replace($xml, '<w:keepNext\b[^>]*/>', '')
    $xml = [regex]::Replace($xml, '<w:keepLines\b[^>]*/>', '')
    $xml = [regex]::Replace($xml, '<w:widowControl\b[^>]*/>', '')

    # A heading belongs to the text beneath it, and a table caption sits ABOVE
    # its table, so both need keep-with-next. _graphic carries the figure
    # itself: Word has no keep-with-PREVIOUS, so the figure keeps with the
    # caption that follows it. keepLines stops a two-line heading or caption
    # splitting across the break.
    $keepWithNext = @(
        'Heading1','Heading2','Heading3','Heading4','Heading5','Heading6',
        'Heading7','Heading8','Heading9','Heading4Sub-Head',
        'Frontmatterhead1','Frontmatterhead2','Frontmatterhead3','Frontmatterhead4',
        'AppendixHead','AppendixHead2','AppendixHead3','AppendixHead4',
        'Versohead','Series','Versohead',
        'CaptionTable','Caption_Table','graphic'
    ) | Select-Object -Unique
    # Captions below a figure only need to hold themselves together.
    $keepLinesOnly = @('CaptionFigure','Caption_Figure','CaptionFigure2','CaptionFigure2','caption','Caption')

    $kn = 0; $kl = 0
    foreach ($id in $keepWithNext) {
        $r = Edit-StyleBlock -Xml $xml -Id $id -Mutator {
            param($block)
            $add = '<w:keepNext/><w:keepLines/>'
            if ($block -match '<w:pPr>') {
                return [regex]::Replace($block, '<w:pPr>', "<w:pPr>$add", 1)
            } elseif ($block -match '<w:pPr\s*/>') {
                return [regex]::Replace($block, '<w:pPr\s*/>', "<w:pPr>$add</w:pPr>", 1)
            }
            return [regex]::Replace($block, '</w:style>', "<w:pPr>$add</w:pPr></w:style>", 1)
        }
        $xml = $r[0]
        if ($r[1]) { $kn++ }
    }
    foreach ($id in $keepLinesOnly) {
        $r = Edit-StyleBlock -Xml $xml -Id $id -Mutator {
            param($block)
            $add = '<w:keepLines/>'
            if ($block -match '<w:pPr>') {
                return [regex]::Replace($block, '<w:pPr>', "<w:pPr>$add", 1)
            } elseif ($block -match '<w:pPr\s*/>') {
                return [regex]::Replace($block, '<w:pPr\s*/>', "<w:pPr>$add</w:pPr>", 1)
            }
            return [regex]::Replace($block, '</w:style>', "<w:pPr>$add</w:pPr></w:style>", 1)
        }
        $xml = $r[0]
        if ($r[1]) { $kl++ }
    }
    Write-Host "  Pagination: cleared (keepNext $keepNextHits, keepLines $keepLinesHits, widow $widowHits); keep-with-next on $kn style(s), keep-lines on $kl more; widow/orphan left ON" -ForegroundColor DarkGray

    # List, List Bullet, List Continue, and List Number share ONE ladder so they
    # line up. The bullet/number marker steps 0.2" per level (0.20/0.40/0.60/
    # 0.80/1.00"); the text sits 0.15" (216 twips) to its right (hanging=216).
    # Plain List and List Continue have no marker (Hanging=0) but sit at the same
    # text column.
    #   Left (text) = marker + 216; Left column = 504/792/1080/1368/1656.
    $listIndents = [ordered]@{
        # Level 1: bullet 0.20" in, text 0.35", 0.15" gap.
        'ListBullet'  = @{ Left = 504;  Hanging = 216 }   # marker 0.20", text 0.35"
        'ListBullet2' = @{ Left = 792;  Hanging = 216 }   # marker 0.40", text 0.55"
        'ListBullet3' = @{ Left = 1080; Hanging = 216 }   # marker 0.60", text 0.75"
        'ListBullet4' = @{ Left = 1368; Hanging = 216 }   # marker 0.80", text 0.95"
        'ListBullet5' = @{ Left = 1656; Hanging = 216 }   # marker 1.00", text 1.15"
        'ListNumber'  = @{ Left = 504;  Hanging = 216 }
        'ListNumber2' = @{ Left = 792;  Hanging = 216 }
        'ListNumber3' = @{ Left = 1080; Hanging = 216 }
        'ListNumber4' = @{ Left = 1368; Hanging = 216 }
        'ListNumber5' = @{ Left = 1656; Hanging = 216 }
        # List Continue: continuation paragraphs (no marker), text aligned to
        # the same column as the matching bullet/number level. Hanging=0.
        'ListContinue'  = @{ Left = 504;  Hanging = 0 }
        'ListContinue2' = @{ Left = 792;  Hanging = 0 }
        'ListContinue3' = @{ Left = 1080; Hanging = 0 }
        'ListContinue4' = @{ Left = 1368; Hanging = 0 }
        'ListContinue5' = @{ Left = 1656; Hanging = 0 }
        # List: plain indented list paragraph (no marker, no hanging).
        'List'  = @{ Left = 504;  Hanging = 0 }
        'List2' = @{ Left = 792;  Hanging = 0 }
        'List3' = @{ Left = 1080; Hanging = 0 }
        'List4' = @{ Left = 1368; Hanging = 0 }
        'List5' = @{ Left = 1656; Hanging = 0 }
        # List Paragraph: applied by the ribbon's bullet/number buttons.
        # Same level-1 layout as ListBullet so single-level ribbon use matches.
        'ListParagraph' = @{ Left = 504; Hanging = 216 }
        # Table Bullet: bullets inside tables/sidebars. Bullet 0.1" in from the
        # cell edge, gap 0.13" to the text (text at 0.23"). Tight for a cell.
        'TableBullet'   = @{ Left = 331; Hanging = 187 }   # marker 0.1", text 0.23", gap 0.13"
    }

    # Per-style bullet glyph overrides. If a style appears here, the numbering
    # propagation pass replaces the <w:lvlText> and <w:rFonts> in the linked
    # abstractNum's ilvl=0. Word's classic List Bullet 2 uses Courier New "o",
    # which renders as a small open circle visually distinct from List Bullet's
    # filled Symbol-font dot.
    # Glyph progression (all pure-ASCII characters; the font does the visual work):
    #   L1  default filled disc       (Symbol font, untouched)
    #   L2  open circle               Courier New "o"
    #   L3  filled small square       Wingdings  "n"
    #   L4  filled diamond            Wingdings  "u"
    #   L5  open square               Wingdings  "q"
    $listBulletOverrides = @{
        'ListBullet2' = @{ LvlText = 'o'; Font = 'Courier New' }
        'ListBullet3' = @{ LvlText = 'n'; Font = 'Wingdings' }
        'ListBullet4' = @{ LvlText = 'u'; Font = 'Wingdings' }
        'ListBullet5' = @{ LvlText = 'q'; Font = 'Wingdings' }
    }
    foreach ($id in $listIndents.Keys) {
        $left = $listIndents[$id].Left
        $hang = $listIndents[$id].Hanging
        # Emit w:right only for entries that specify it (e.g. ListBullet right=0);
        # the others keep their inherited right indent as before.
        $rightAttr = if ($listIndents[$id].ContainsKey('Right')) { " w:right=`"$($listIndents[$id].Right)`"" } else { '' }
        $newInd = "<w:ind w:left=`"$left`" w:hanging=`"$hang`"$rightAttr/>"
        $r = Edit-StyleBlock -Xml $xml -Id $id -Mutator {
            param($block)
            if ($block -match '<w:ind\s+[^/]*?/>') {
                return [regex]::Replace($block, '<w:ind\s+[^/]*?/>', $newInd, 1)
            } elseif ($block -match '</w:pPr>') {
                return [regex]::Replace($block, '</w:pPr>', "$newInd</w:pPr>", 1)
            } elseif ($block -match '<w:pPr>') {
                return [regex]::Replace($block, '<w:pPr>', "<w:pPr>$newInd", 1)
            } else {
                return [regex]::Replace($block, '</w:style>', "<w:pPr>$newInd</w:pPr></w:style>", 1)
            }
        }.GetNewClosure()
        $xml = $r[0]
        if ($r[1]) { Write-Host "  ${id}: ind left=$left hanging=$hang" -ForegroundColor DarkGray }
    }

    # List Bullet / List Number family: single line spacing
    $listIds = @('ListBullet','ListBullet2','ListBullet3','ListBullet4','ListBullet5',
                 'ListNumber','ListNumber2','ListNumber3','ListNumber4','ListNumber5',
                 'ListContinue','ListContinue2','ListContinue3','ListContinue4','ListContinue5',
                 'List','List2','List3','List4','List5',
                 'ListParagraph')
    $listSingleHits = 0
    foreach ($id in $listIds) {
        $r = Edit-StyleBlock -Xml $xml -Id $id -Mutator $ListSingleSpacingEdit
        $xml = $r[0]
        if ($r[1]) { $listSingleHits++ }
    }
    Write-Host "  List family: single line spacing on $listSingleHits style(s)" -ForegroundColor DarkGray

    # BodyText: strip any spacing override so it inherits xxbase cleanly
    $bt = Edit-StyleBlock -Xml $xml -Id 'BodyText' -Mutator $bodyTextEdit
    $xml = $bt[0]
    if ($bt[1]) { Write-Host "  BodyText: spacing override stripped (inherits xxbase)" -ForegroundColor DarkGray }

    # Caption: Source Sans Pro Semibold 10pt, no bold
    foreach ($id in @('Caption')) {
        $cap = Edit-StyleBlock -Xml $xml -Id $id -Mutator $captionFontEdit
        $xml = $cap[0]
        if ($cap[1]) { Write-Host "  ${id}: Source Sans Pro Semibold 10pt, no bold, +0.2pt tracking" -ForegroundColor DarkGray }
    }

    # Caption_Figure_2: Source Sans Pro 10pt, no bold, 4pt padding, 0.9 line, light grey background
    $cf2 = Edit-StyleBlock -Xml $xml -Id 'Caption_Figure_2' -Mutator $captionFigure2Edit
    $xml = $cf2[0]
    if ($cf2[1]) { Write-Host "  Caption_Figure_2: Source Sans Pro 10pt, no bold, 4pt padding, 0.9 line, light grey background" -ForegroundColor DarkGray }

    # Caption_Figure_2 styleId variations (Word strips underscore)
    $cf2b = Edit-StyleBlock -Xml $xml -Id 'CaptionFigure2' -Mutator $captionFigure2Edit
    $xml = $cf2b[0]
    if ($cf2b[1]) { Write-Host "  CaptionFigure2: Source Sans Pro 10pt, no bold, 4pt padding, 0.9 line, light grey background" -ForegroundColor DarkGray }

    # Header Odd / Header Even / Footer Odd / Footer Even: create if absent, replace if present
    foreach ($pair in @(@{ Id='HeaderOdd';  Block=$headerOddBlock },
                        @{ Id='HeaderEven'; Block=$headerEvenBlock },
                        @{ Id='Headerch';  Block=$headerChBlock },
                        @{ Id='FooterOdd';  Block=$footerOddBlock },
                        @{ Id='FooterEven'; Block=$footerEvenBlock })) {
        $rmResult = Remove-StyleBlock -Xml $xml -Id $pair.Id
        $xml = $rmResult[0]
        # Insert before </w:styles>
        if ($xml -match '</w:styles>') {
            $xml = [regex]::Replace($xml, '</w:styles>', ($pair.Block + '</w:styles>'), 1)
            Write-Host "  $($pair.Id): created/replaced" -ForegroundColor DarkGray
        }
    }

    # Leaf styles as canonical definitions (create-or-replace), instead of
    # per-property regex surgery. Each linked pair shares ONE rPr string so the
    # paragraph style and its Char style are always identical:
    #   Versohead               - Source Sans Pro Bold, black (stays black).
    #   Title / TitleChar       - Source Sans Pro Bold, Resolution Blue, 18pt.
    #   MTDisplayEquation / *Char - Source Serif Pro 12pt (Merriweather retired).
    $titleRpr = '<w:rPr><w:rFonts w:ascii="Source Sans Pro" w:eastAsia="Times New Roman" w:hAnsi="Source Sans Pro" w:cs="Source Sans Pro"/><w:b/><w:bCs/><w:color w:val="00257A"/><w:spacing w:val="-12"/><w:sz w:val="36"/><w:szCs w:val="60"/><w14:ligatures w14:val="none"/></w:rPr>'
    $eqRpr    = '<w:rPr><w:rFonts w:ascii="Source Serif Pro" w:eastAsiaTheme="minorHAnsi" w:hAnsi="Source Serif Pro" w:cstheme="minorBidi"/><w:bCs/><w:iCs/><w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="24"/><w14:ligatures w14:val="none"/></w:rPr>'
    $leafStyles = [ordered]@{
        'Versohead' = '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Versohead"><w:name w:val="Verso_head"/><w:basedOn w:val="xxheader"/><w:uiPriority w:val="110"/><w:qFormat/><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:spacing w:before="500"/></w:pPr><w:rPr><w:rFonts w:ascii="Source Sans Pro" w:hAnsi="Source Sans Pro"/><w:b/><w:bCs/><w:color w:val="000000"/><w14:ligatures w14:val="none"/></w:rPr></w:style>'
        'Title' = '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="109"/><w:next w:val="BodyText"/><w:link w:val="TitleChar"/><w:qFormat/><w:pPr><w:framePr w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:xAlign="right" w:y="4321"/><w:suppressAutoHyphens/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:spacing w:after="160"/><w:suppressOverlap/><w:textAlignment w:val="center"/></w:pPr>' + $titleRpr + '</w:style>'
        'TitleChar' = '<w:style w:type="character" w:customStyle="1" w:styleId="TitleChar"><w:name w:val="Title Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:link w:val="Title"/><w:unhideWhenUsed/>' + $titleRpr + '</w:style>'
        'MTDisplayEquation' = '<w:style w:type="paragraph" w:customStyle="1" w:styleId="MTDisplayEquation"><w:name w:val="MTDisplayEquation"/><w:basedOn w:val="BodyText"/><w:next w:val="Normal"/><w:link w:val="MTDisplayEquationChar"/><w:uiPriority w:val="88"/><w:qFormat/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4320"/><w:tab w:val="right" w:pos="8640"/></w:tabs><w:spacing w:before="240" w:after="240"/><w:ind w:firstLine="360"/></w:pPr>' + $eqRpr + '</w:style>'
        'MTDisplayEquationChar' = '<w:style w:type="character" w:customStyle="1" w:styleId="MTDisplayEquationChar"><w:name w:val="MTDisplayEquation Char"/><w:basedOn w:val="BodyTextChar"/><w:link w:val="MTDisplayEquation"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>' + $eqRpr + '</w:style>'
    }
    foreach ($id in $leafStyles.Keys) {
        $rm = Remove-StyleBlock -Xml $xml -Id $id
        $xml = $rm[0]
        if ($xml -match '</w:styles>') {
            $xml = [regex]::Replace($xml, '</w:styles>', ($leafStyles[$id] + '</w:styles>'), 1)
            Write-Host "  ${id}: canonical definition" -ForegroundColor DarkGray
        }
    }

    # Apply explicit deletions of junk / paste-artifact styles. (Legacy-fix:
    # a clean template has none of these, so -Clean skips the pass.)
    if (-not $Clean) {
        $deleteHits = 0
        foreach ($id in $delete) {
            $result = Remove-StyleBlock -Xml $xml -Id $id
            $xml = $result[0]
            if ($result[1]) { $deleteHits++ }
        }
        Write-Host "  Deleted (explicit list): $deleteHits style(s)" -ForegroundColor DarkGray
    }

    # ---------------------------------------------------------------
    # Recommended-styles curation: which styles appear in Word's
    # "Recommended" pane, and assign uiPriority values in alphabetical
    # order so "Sort by: As Recommended" shows them A-Z.
    # ---------------------------------------------------------------

    # Built-in styleIds we want shown in Recommended (in addition to all
    # custom styles that aren't bases or character-style siblings)
    $builtinRecommend = @(
        'Heading1','Heading2','Heading3','Heading4','Heading5','Heading6',
        'Heading7','Heading8','Heading9',
        'Title','Subtitle','BodyText','Caption','Quote','IntenseQuote',
        'ListBullet','ListBullet2','ListBullet3','ListBullet4','ListBullet5',
        'ListNumber','ListNumber2','ListNumber3','ListNumber4','ListNumber5',
        'ListContinue','ListContinue2','ListContinue3','ListContinue4','ListContinue5',
        'List','List2','List3','List4','List5',
        'ListParagraph'
    )

    # Parse all styles in the file
    $styleParse = [regex]'(?s)<w:style\s+([^>]*?)>(.*?)</w:style>'
    $allStyles = @()
    foreach ($m in $styleParse.Matches($xml)) {
        $a = $m.Groups[1].Value
        $b = $m.Groups[2].Value
        $sid = ''; $stype = ''; $sname = ''
        if ($a -match 'w:styleId="([^"]+)"') { $sid = $Matches[1] }
        if ($a -match 'w:type="([^"]+)"') { $stype = $Matches[1] }
        $isCustom = $a.Contains('w:customStyle="1"')
        if ($b -match '<w:name w:val="([^"]+)"') { $sname = $Matches[1] }
        if ($sid) { $allStyles += [pscustomobject]@{ Id = $sid; Type = $stype; IsCustom = $isCustom; Name = $sname } }
    }

    # Classify into recommend vs hide
    $recList = @(); $hideList = @()
    foreach ($s in $allStyles) {
        if ($s.Type -eq 'numbering') {
            # Custom numbering styles (e.g. HMList "HM List", which carries the
            # Chapter/Section/Subsection auto-numbering pattern) are promoted to
            # Recommended so authors can find them in the Styles pane without
            # switching to "All Styles". Built-in numbering (NoList) is left alone.
            if ($s.IsCustom) { $recList += $s }
            continue
        }
        if ($s.Type -eq 'character') { $hideList += $s.Id; continue }
        if ($s.Type -eq 'table') {
            if ($s.Id -in @('HM-Cover','HM-Data','NIST-Data','SidebarTable')) { $recList += $s } else { $hideList += $s.Id }
            continue
        }
        # paragraph
        if (-not $s.IsCustom) {
            if ($builtinRecommend -contains $s.Id) { $recList += $s } else { $hideList += $s.Id }
            continue
        }
        # custom paragraph: hide xx* bases and *Char siblings; recommend the rest
        if ($s.Name -match '^xx' -or $s.Name -match 'Char$') { $hideList += $s.Id }
        else { $recList += $s }
    }

    # Set-StyleRecommended and Set-StyleHidden are provided by the OoXml module.

    # Sort the recommend list alphabetical by Name and assign sequential uiPriority
    $sortedRec = $recList | Sort-Object Name
    $priority = 10
    foreach ($s in $sortedRec) {
        $xml = Set-StyleRecommended -Xml $xml -Id $s.Id -Priority $priority
        $priority++
    }
    foreach ($id in $hideList) {
        $xml = Set-StyleHidden -Xml $xml -Id $id
    }
    Write-Host "  Recommended: $($sortedRec.Count) style(s) alphabetized; $($hideList.Count) hidden" -ForegroundColor DarkGray

    # Cut table-style bloat
    $allTableIds = Get-AllStyleIdsByType -Xml $xml -Type 'table'
    $tableDeleteHits = 0
    foreach ($id in $allTableIds) {
        if ($tableKeep -contains $id) { continue }
        $result = Remove-StyleBlock -Xml $xml -Id $id
        $xml = $result[0]
        if ($result[1]) { $tableDeleteHits++ }
    }
    Write-Host "  Deleted (table bloat): $tableDeleteHits style(s)" -ForegroundColor DarkGray

    # ---- Create the List Number family when the document has no definition ----
    # Edit-StyleBlock only edits what is already there. A document that never
    # carried List Number therefore came out of every pass above untouched, and
    # Word then invents the style from its own built-ins: Times New Roman, based
    # on Normal, ignoring the whole xx* architecture. List Bullet was fine in the
    # same files only because it happened to be defined.
    #
    # Creating the style alone is not enough - a numPr pointing at a numId that
    # does not exist renders unnumbered - so the abstractNum and num are built
    # alongside it, with ids allocated above whatever the document already uses.
    if ($numberingXml) {
        $numberFormats = [ordered]@{
            'ListNumber'  = @{ Fmt = 'decimal';     Text = '%1.' }
            'ListNumber2' = @{ Fmt = 'lowerLetter'; Text = '%1)' }
            'ListNumber3' = @{ Fmt = 'lowerRoman';  Text = '%1.' }
            'ListNumber4' = @{ Fmt = 'decimal';     Text = '%1.' }
            'ListNumber5' = @{ Fmt = 'decimal';     Text = '%1.' }
        }
        $missingNumber = @($numberFormats.Keys | Where-Object {
            -not [regex]::IsMatch($xml, '<w:style\b[^>]*\bw:styleId="' + $_ + '"')
        })
        if ($missingNumber.Count -gt 0) {
            # Next free ids. abstractNumId and numId are separate sequences.
            $absIds = [regex]::Matches($numberingXml, '<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"') |
                      ForEach-Object { [int]$_.Groups[1].Value }
            $numIds = [regex]::Matches($numberingXml, '<w:num\b[^>]*w:numId="(\d+)"') |
                      ForEach-Object { [int]$_.Groups[1].Value }
            $nextAbs = 1 + (($absIds | Measure-Object -Maximum).Maximum)
            $nextNum = 1 + (($numIds | Measure-Object -Maximum).Maximum)

            $newAbstracts = ''
            $newNums      = ''
            foreach ($id in $missingNumber) {
                $left = $listIndents[$id].Left
                $hang = $listIndents[$id].Hanging
                $fmt  = $numberFormats[$id].Fmt
                $text = $numberFormats[$id].Text

                # One single-level abstractNum per style, matching how the
                # existing List Bullet family is built: these are flat lists, not
                # levels of one nested list.
                $newAbstracts += '<w:abstractNum w:abstractNumId="' + $nextAbs + '">' +
                    '<w:multiLevelType w:val="singleLevel"/>' +
                    '<w:lvl w:ilvl="0">' +
                    '<w:start w:val="1"/>' +
                    '<w:numFmt w:val="' + $fmt + '"/>' +
                    '<w:lvlText w:val="' + $text + '"/>' +
                    '<w:lvlJc w:val="left"/>' +
                    '<w:pPr><w:ind w:left="' + $left + '" w:hanging="' + $hang + '"/></w:pPr>' +
                    '</w:lvl></w:abstractNum>'
                $newNums += '<w:num w:numId="' + $nextNum + '"><w:abstractNumId w:val="' + $nextAbs + '"/></w:num>'

                $styleBlock = '<w:style w:type="paragraph" w:styleId="' + $id + '">' +
                    '<w:name w:val="' + ($id -replace '^ListNumber', 'List Number ').TrimEnd() + '"/>' +
                    '<w:basedOn w:val="xxbase"/>' +
                    '<w:uiPriority w:val="66"/><w:qFormat/>' +
                    '<w:pPr><w:numPr><w:numId w:val="' + $nextNum + '"/></w:numPr>' +
                    '<w:contextualSpacing/>' +
                    '<w:ind w:left="' + $left + '" w:hanging="' + $hang + '"/></w:pPr>' +
                    '</w:style>'
                $xml = [regex]::Replace($xml, '</w:styles>', $styleBlock + '</w:styles>', 1)

                Write-Host "  ${id}: created (numId $nextNum, $fmt '$text', ind left=$left)" -ForegroundColor DarkGray
                $nextAbs++
                $nextNum++
            }
            # abstractNum elements must precede every num element in numbering.xml.
            if ($numberingXml -match '<w:num\b') {
                $numberingXml = [regex]::Replace($numberingXml, '(<w:num\b)', $newAbstracts + '$1', 1)
            } else {
                $numberingXml = [regex]::Replace($numberingXml, '</w:numbering>', $newAbstracts, 1)
            }
            $numberingXml = [regex]::Replace($numberingXml, '</w:numbering>', $newNums + '</w:numbering>', 1)
            Write-Host "  List Number family: $($missingNumber.Count) style(s) created with numbering" -ForegroundColor DarkGray
        }
    }

    # ---- rPr element order ----
    # The children of w:rPr are a schema sequence, not a set. Successive regex
    # insertions scramble it (a function writing rFonts then sz, both "just
    # after <w:rPr>", leaves sz in front of rFonts). Word reads any order, so
    # nothing looks wrong, but validators and other consumers may reject it.
    # Sorting the whole file once catches styles no mutator touched.
    $beforeSort = ([regex]::Matches($xml, '(?s)<w:rPr(?:\s[^>]*?)?>.*?</w:rPr>')).Count
    $xml = Sort-RunPropertyOrder -Block $xml
    Write-Host "  rPr element order: normalised across $beforeSort run-property block(s)" -ForegroundColor DarkGray

    # Numbering pass: uses OoXml module functions
    if ($numberingXml) {
        $headingAbstractNumIds = Get-HeadingAbstractNumIds -StylesXml $xml -NumberingXml $numberingXml
        $numberingXml = Set-NumberingSuffix -NumberingXml $numberingXml -HeadingAbstractNumIds $headingAbstractNumIds
        Write-Host "  numbering.xml: suff=space on heading numbering, suff=tab on list numbering" -ForegroundColor DarkGray

        $numberingXml = Sync-ListIndents -NumberingXml $numberingXml -StylesXml $xml -IndentMap $listIndents
        Write-Host "  numbering.xml: propagated list indents into abstractNums" -ForegroundColor DarkGray

        $numberingXml = Set-BulletGlyph -NumberingXml $numberingXml -StylesXml $xml -OverrideMap $listBulletOverrides
        Write-Host "  numbering.xml: bullet glyphs overridden" -ForegroundColor DarkGray
    }

    # Write back
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "styles-edit-$([guid]::NewGuid()).zip")
    $srcStream = [System.IO.File]::OpenRead($target)
    try {
        $srcZip = New-Object System.IO.Compression.ZipArchive($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $outStream = [System.IO.File]::Create($tempZip)
            try {
                $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
                try {
                    foreach ($e in $srcZip.Entries) {
                        $newEntry = $outZip.CreateEntry($e.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
                        if ($e.FullName -eq 'word/styles.xml') {
                            $w = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
                            try { $w.Write($xml) } finally { $w.Dispose() }
                        } elseif ($e.FullName -eq 'word/numbering.xml' -and $numberingXml) {
                            $w = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
                            try { $w.Write($numberingXml) } finally { $w.Dispose() }
                        } else {
                            $es = $e.Open()
                            try {
                                $ns = $newEntry.Open()
                                try { $es.CopyTo($ns) } finally { $ns.Dispose() }
                            } finally { $es.Dispose() }
                        }
                    }
                } finally { $outZip.Dispose() }
            } finally { $outStream.Dispose() }
        } finally { $srcZip.Dispose() }
    } finally { $srcStream.Dispose() }

    Move-Item -LiteralPath $tempZip -Destination $target -Force
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done. Open in Word and verify." -ForegroundColor Green
