<#
.SYNOPSIS
    Audits a Word file (or every Word file in a folder) for the presence
    of the NIST style ids that downstream toolkit scripts depend on.
    Catches the "renamed style breaks the script silently" trap.

.DESCRIPTION
    Thin wrapper around publishing/word/Audit-StyleSet.ps1 that supplies the
    48 required and 24 optional NIST NCSTAR/NWIRP style sets.

    Every script in this toolkit (Set-NISTStyleArchitecture, Build-NISTReport,
    Test-NISTFrontMatter, Set-NISTHeaderFooterContent, etc.) keys off
    specific style ids in word/styles.xml. If a template renames one of
    those styles, the downstream script silently misses it.

    Required and expected sets were built by reading the styles.xml of a
    conformant NCSTAR/NWIRP template (144 styles) and cross-referencing
    with the toolkit scripts.

    Exit codes:
        0 = clean (every required style present in every file)
        1 = drift (one or more required styles missing somewhere)
        2 = input error

.PARAMETER File
    Single Word file to audit. Either -File or -Folder is required.

.PARAMETER Folder
    Folder whose Word files to audit. Recurses one level only.

.PARAMETER Filter
    Optional filename filter for -Folder mode (default: *.docx).

.PARAMETER ShowOptional
    Also report missing optional styles (default: only required misses).

.PARAMETER Json
    Emit a JSON report on stdout instead of the colored table.

.EXAMPLE
    .\Audit-NISTStyleSet.ps1 -File ".\NCSTAR_4-3_Chapter-01.docx"

.EXAMPLE
    .\Audit-NISTStyleSet.ps1 -Folder "C:\Users\Admin\Documents\260512" -ShowOptional

.EXAMPLE
    .\Audit-NISTStyleSet.ps1 -Folder "C:\reports" -Json | ConvertFrom-Json
#>

[CmdletBinding(DefaultParameterSetName='Folder')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$Filter = '*.docx',

    [switch]$ShowOptional,

    [switch]$Json
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Canonical NIST style sets
# ---------------------------------------------------------------------------

$requiredStyles = [ordered]@{
    # Cover and title page
    'Title'              = @{ Category='Cover/Title'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1') }
    'HMCoverDate'        = @{ Category='Cover/Title'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1') }
    'Series'             = @{ Category='Cover/Title'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1') }
    'ReportTitle'        = @{ Category='Cover/Title'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1','Set-NISTHeaderFooterContent.ps1') }
    'ReportSubtitle'     = @{ Category='Cover/Title'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1') }
    'DraftStage'         = @{ Category='Cover/Title'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1') }
    'Authors'            = @{ Category='Cover/Title'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1') }

    # Disclaimer / front matter
    'Disclaimer'         = @{ Category='Front matter'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1','Build-NWIRPBuildingBlocks.ps1') }
    'frontmattertext'    = @{ Category='Front matter'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1') }
    'Frontmatterhead1'   = @{ Category='Front matter'; Type='paragraph'; UsedBy=@('Test-NISTFrontMatter.ps1','Set-NISTStyleArchitecture.ps1') }

    # Heading ladder
    'Heading1'           = @{ Category='Headings'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1','Set-NISTHeaderFooterContent.ps1','Build-NISTReport.ps1') }
    'Heading2'           = @{ Category='Headings'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'Heading3'           = @{ Category='Headings'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'Heading4'           = @{ Category='Headings'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'Heading1Char'       = @{ Category='Headings'; Type='character'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }

    # Appendix
    'AppendixHead'       = @{ Category='Appendix'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1','Build-NISTReport.ps1') }
    'AppendixHead2'      = @{ Category='Appendix'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'AppendixHead3'      = @{ Category='Appendix'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'AppendixHead4'      = @{ Category='Appendix'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }

    # Body
    'BodyText'           = @{ Category='Body'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1','Reapply-Styles.ps1','Build-NWIRPBuildingBlocks.ps1') }

    # Captions
    'Caption'            = @{ Category='Captions'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1','Build-NWIRPBuildingBlocks.ps1') }
    'CaptionFigure'      = @{ Category='Captions'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'CaptionFigure2'     = @{ Category='Captions'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'CaptionTable'       = @{ Category='Captions'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }

    # Lists
    'ListBullet'         = @{ Category='Lists'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'ListBullet2'        = @{ Category='Lists'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'ListBullet3'        = @{ Category='Lists'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'ListBullet4'        = @{ Category='Lists'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'ListBullet5'        = @{ Category='Lists'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'ListNumber'         = @{ Category='Lists'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'ListParagraph'      = @{ Category='Lists'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }

    # Headers and footers
    'Header'             = @{ Category='Header/Footer'; Type='paragraph'; UsedBy=@('Set-NISTHeaderFooterContent.ps1') }
    'Footer'             = @{ Category='Header/Footer'; Type='paragraph'; UsedBy=@('Set-NISTHeaderFooterContent.ps1') }
    'HeaderEven'         = @{ Category='Header/Footer'; Type='paragraph'; UsedBy=@('Set-NISTHeaderFooterContent.ps1') }
    'HeaderOdd'          = @{ Category='Header/Footer'; Type='paragraph'; UsedBy=@('Set-NISTHeaderFooterContent.ps1') }
    'FooterEven'         = @{ Category='Header/Footer'; Type='paragraph'; UsedBy=@('Set-NISTHeaderFooterContent.ps1') }
    'FooterOdd'          = @{ Category='Header/Footer'; Type='paragraph'; UsedBy=@('Set-NISTHeaderFooterContent.ps1') }

    # TOC / lists of stuff
    'TOC1'               = @{ Category='TOC/Lists-of'; Type='paragraph'; UsedBy=@('Build-NISTReport.ps1') }
    'TOC2'               = @{ Category='TOC/Lists-of'; Type='paragraph'; UsedBy=@('Build-NISTReport.ps1') }
    'TOC3'               = @{ Category='TOC/Lists-of'; Type='paragraph'; UsedBy=@('Build-NISTReport.ps1') }
    'TOC4'               = @{ Category='TOC/Lists-of'; Type='paragraph'; UsedBy=@('Build-NISTReport.ps1') }
    'TableofFigures'     = @{ Category='TOC/Lists-of'; Type='paragraph'; UsedBy=@('Build-NISTReport.ps1') }

    # Footnotes
    'FootnoteText'       = @{ Category='Footnotes'; Type='paragraph'; UsedBy=@('Build-NISTReport.ps1') }
    'FootnoteReference'  = @{ Category='Footnotes'; Type='character'; UsedBy=@('Build-NISTReport.ps1') }

    # Tables
    'HM-Cover'           = @{ Category='Tables'; Type='table'; UsedBy=@('Test-NISTFrontMatter.ps1') }
    'NIST-Data'          = @{ Category='Tables'; Type='table'; UsedBy=@('Set-NISTStyleArchitecture.ps1','Set-NISTTableLook.ps1') }
    'TableGrid'          = @{ Category='Tables'; Type='table'; UsedBy=@('Set-NISTTableLook.ps1') }
    'TableText'          = @{ Category='Tables'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'TableandFigureFootnotes' = @{ Category='Tables'; Type='paragraph'; UsedBy=@('Build-NWIRPBuildingBlocks.ps1') }

    # Numbering chain
    'HMList'             = @{ Category='Numbering'; Type='numbering'; UsedBy=@('Set-NISTStyleArchitecture.ps1','Build-NISTReport.ps1') }

    # Architecture base
    'xxbase'             = @{ Category='Architecture'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
}

$optionalStyles = [ordered]@{
    'Versohead'          = @{ Category='Front matter'; Type='paragraph' }
    'Versotext'          = @{ Category='Front matter'; Type='paragraph' }
    'Frontmatterhead2'   = @{ Category='Front matter'; Type='paragraph' }
    'Frontmatterhead3'   = @{ Category='Front matter'; Type='paragraph' }
    'Frontmatterhead4'   = @{ Category='Front matter'; Type='paragraph' }
    'AppendixText'       = @{ Category='Appendix'; Type='paragraph' }
    'Heading5'           = @{ Category='Headings'; Type='paragraph' }
    'Heading6'           = @{ Category='Headings'; Type='paragraph' }
    'ListContinue'       = @{ Category='Lists'; Type='paragraph' }
    'ListContinue2'      = @{ Category='Lists'; Type='paragraph' }
    'ListContinue3'      = @{ Category='Lists'; Type='paragraph' }
    'TOCHeading'         = @{ Category='TOC/Lists-of'; Type='paragraph' }
    'TOC5'               = @{ Category='TOC/Lists-of'; Type='paragraph' }
    'References'         = @{ Category='Body'; Type='paragraph' }
    'Bibliography'       = @{ Category='Body'; Type='paragraph' }
    'Quote'              = @{ Category='Body'; Type='paragraph' }
    'IntenseQuote'       = @{ Category='Body'; Type='paragraph' }
    'Hyperlink'          = @{ Category='Body'; Type='character' }
    'Term-Abbrev'        = @{ Category='Body'; Type='paragraph' }
    'HM-Data'            = @{ Category='Tables'; Type='table' }
    'TableBullet'        = @{ Category='Tables'; Type='paragraph' }
    'Appendix'           = @{ Category='Numbering'; Type='numbering' }
    'xxheader'           = @{ Category='Architecture'; Type='paragraph' }
    'xxgraphic'          = @{ Category='Architecture'; Type='paragraph' }
    'xxtable'            = @{ Category='Architecture'; Type='paragraph' }
}

# ---------------------------------------------------------------------------
# Delegate to generic audit engine
# ---------------------------------------------------------------------------

$genericScript = Join-Path (Split-Path $PSScriptRoot -Parent) 'word\Audit-StyleSet.ps1'

$args_ = @{
    RequiredStyles = $requiredStyles
    OptionalStyles = $optionalStyles
    ReportTitle    = 'NIST style set audit'
    Filter         = $Filter
    ShowOptional   = $ShowOptional
    Json           = $Json
}

if ($PSCmdlet.ParameterSetName -eq 'File') {
    $args_.File = $File
} else {
    $args_.Folder = $Folder
}

& $genericScript @args_
