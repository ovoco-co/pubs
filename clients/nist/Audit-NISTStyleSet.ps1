<#
.SYNOPSIS
    Audits a Word file (or every Word file in a folder) for the presence
    of the NIST style ids that downstream toolkit scripts depend on.
    Catches the "renamed style breaks the script silently" trap.

.DESCRIPTION
    Every script in this toolkit (Set-NISTStyleArchitecture, Build-NISTReport,
    Test-NISTFrontMatter, Set-NISTHeaderFooterContent, etc.) keys off
    specific style ids in word/styles.xml. If a template renames one of
    those styles, the downstream script silently misses it: numbering
    drifts, the validator under-reports, captions don't pick up the right
    font.

    This script parses styles.xml directly, lists every defined styleId,
    and compares against:
        - a REQUIRED set whose absence breaks at least one downstream script
        - an EXPECTED set that is usually present and good to flag if absent

    Each REQUIRED entry includes a "UsedBy" list naming which script(s)
    depend on it, so a missing-style report points straight at what will
    break next.

    Required and expected sets were built by reading the styles.xml of a
    conformant NCSTAR/NWIRP template (144 styles) and cross-referencing
    with the toolkit scripts. They are intentionally permissive: only
    style ids that are actively used in code or assembly are gated as
    REQUIRED. Visual / authoring styles (Quote, BlockText, etc.) are not
    enforced.

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
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Canonical sets
# ---------------------------------------------------------------------------
# Each entry: { Category, Type, UsedBy }
#   Category - groups the report output
#   Type     - paragraph | character | table | numbering
#   UsedBy   - list of script filenames that key off this style id
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
    'BodyText'           = @{ Category='Body'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1','Reapply-NISTStyles.ps1','Build-NWIRPBuildingBlocks.ps1') }

    # Captions
    'Caption'            = @{ Category='Captions'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1','Build-NWIRPBuildingBlocks.ps1') }
    'CaptionFigure'      = @{ Category='Captions'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'CaptionFigure2'     = @{ Category='Captions'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }
    'CaptionTable'       = @{ Category='Captions'; Type='paragraph'; UsedBy=@('Set-NISTStyleArchitecture.ps1') }

    # Lists (the indent + glyph ladder the architecture script rewrites)
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

    # Numbering chain (chapter/section auto-numbering)
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
# File list
# ---------------------------------------------------------------------------

$files = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) {
        Write-Error "File not found: $File"; exit 2
    }
    $files += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
        Write-Error "Folder not found: $Folder"; exit 2
    }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $files = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles($Filter) |
        Where-Object { $_.Extension -imatch '^\.(docx|dotx|docm|dotm)$' } |
        ForEach-Object { $_.FullName }
}

if ($files.Count -eq 0) { Write-Error "No files to audit."; exit 2 }

# ---------------------------------------------------------------------------
# Audit each file
# ---------------------------------------------------------------------------

function Get-StyleIds {
    param([string]$Path)
    $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, 'Read')
        try {
            $entry = $zip.GetEntry('word/styles.xml')
            if (-not $entry) { return @{} }
            $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            try { $stylesXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }

    # styleId -> type
    $present = @{}
    $matches = [regex]::Matches($stylesXml, '<w:style\s+w:type="(\w+)"[^>]*w:styleId="([^"]+)"')
    foreach ($m in $matches) {
        $present[$m.Groups[2].Value] = $m.Groups[1].Value
    }
    return $present
}

$reportFiles = New-Object System.Collections.ArrayList
$anyRequiredMiss = $false

foreach ($f in $files) {
    $present = Get-StyleIds -Path $f
    $missingRequired = @()
    $typeMismatch    = @()
    $missingOptional = @()

    foreach ($k in $requiredStyles.Keys) {
        if (-not $present.ContainsKey($k)) {
            $missingRequired += [pscustomobject]@{
                StyleId  = $k
                Category = $requiredStyles[$k].Category
                Type     = $requiredStyles[$k].Type
                UsedBy   = $requiredStyles[$k].UsedBy -join ', '
            }
        } elseif ($present[$k] -ne $requiredStyles[$k].Type) {
            $typeMismatch += [pscustomobject]@{
                StyleId  = $k
                Expected = $requiredStyles[$k].Type
                Actual   = $present[$k]
                Category = $requiredStyles[$k].Category
                UsedBy   = $requiredStyles[$k].UsedBy -join ', '
            }
        }
    }
    foreach ($k in $optionalStyles.Keys) {
        if (-not $present.ContainsKey($k)) {
            $missingOptional += [pscustomobject]@{
                StyleId  = $k
                Category = $optionalStyles[$k].Category
                Type     = $optionalStyles[$k].Type
            }
        }
    }

    if ($missingRequired.Count -gt 0 -or $typeMismatch.Count -gt 0) { $anyRequiredMiss = $true }

    [void]$reportFiles.Add([pscustomobject]@{
        File             = $f
        TotalStyles      = $present.Count
        MissingRequired  = $missingRequired
        TypeMismatch     = $typeMismatch
        MissingOptional  = $missingOptional
    })
}

# ---------------------------------------------------------------------------
# Render
# ---------------------------------------------------------------------------

if ($Json) {
    [pscustomobject]@{
        Clean = -not $anyRequiredMiss
        Files = $reportFiles
    } | ConvertTo-Json -Depth 6
} else {
    Write-Host ""
    Write-Host "NIST style set audit" -ForegroundColor Cyan
    Write-Host ""
    foreach ($r in $reportFiles) {
        $name = Split-Path -Leaf $r.File
        Write-Host "  $name" -ForegroundColor White
        Write-Host ("    Total styles defined: {0}" -f $r.TotalStyles) -ForegroundColor DarkGray
        if ($r.MissingRequired.Count -eq 0 -and $r.TypeMismatch.Count -eq 0) {
            Write-Host "    Required styles: all present" -ForegroundColor Green
        } else {
            if ($r.MissingRequired.Count -gt 0) {
                Write-Host ("    MISSING required ({0}):" -f $r.MissingRequired.Count) -ForegroundColor Red
                foreach ($m in $r.MissingRequired) {
                    Write-Host ("      - {0,-26} [{1,-12}] {2}" -f $m.StyleId, $m.Category, "needed by: $($m.UsedBy)") -ForegroundColor Red
                }
            }
            if ($r.TypeMismatch.Count -gt 0) {
                Write-Host ("    TYPE MISMATCH ({0}):" -f $r.TypeMismatch.Count) -ForegroundColor Yellow
                foreach ($m in $r.TypeMismatch) {
                    Write-Host ("      - {0,-26} expected {1}, found {2}" -f $m.StyleId, $m.Expected, $m.Actual) -ForegroundColor Yellow
                }
            }
        }
        if ($ShowOptional -and $r.MissingOptional.Count -gt 0) {
            Write-Host ("    Missing optional ({0}):" -f $r.MissingOptional.Count) -ForegroundColor DarkGray
            foreach ($m in $r.MissingOptional) {
                Write-Host ("      - {0,-26} [{1}]" -f $m.StyleId, $m.Category) -ForegroundColor DarkGray
            }
        }
        Write-Host ""
    }
    if ($anyRequiredMiss) {
        Write-Host "Drift: at least one file is missing a required style id." -ForegroundColor Red
    } else {
        Write-Host "Clean. Every audited file defines every required style id." -ForegroundColor Green
    }
    Write-Host ""
}

if ($anyRequiredMiss) { exit 1 } else { exit 0 }
