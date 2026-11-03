<#
.SYNOPSIS
    Adds the NIST NCSTAR / NWIRP building block set to an unpacked Word
    template's glossary. Creates the glossary scaffolding if missing.

.DESCRIPTION
    Run after Expand-WordPackage.ps1 has unpacked your template. This
    script:

      1. Creates word/glossary/ scaffolding (document.xml, styles.xml,
         settings.xml, webSettings.xml, fontTable.xml, _rels) if any of
         those parts are missing.
      2. Registers the glossary parts in [Content_Types].xml.
      3. Adds 22 NIST NCSTAR / NWIRP building blocks to the glossary,
         spanning cover/front matter, captions, comparison tables,
         model results, end matter, and stats reporting.

    The blocks reference styles that already exist in your template
    (Heading1NoNumber, BodyText, Caption, Disclaimer, References,
    TableandFigureFootnotes). If a style does not exist Word will fall
    back to Normal and the block will still insert; you can then edit
    the block in Word and re-save it.

    After running, repack with Compress-WordPackage.ps1, open the
    result in Word, refine any block in place, and re-save the Quick
    Part if you want to evolve it.

.PARAMETER SourceDir
    The unpacked folder produced by Expand-WordPackage.ps1.

.PARAMETER Force
    Skip prompts when overwriting an existing glossary/document.xml.

.EXAMPLE
    .\Expand-WordPackage.ps1 -File "HM-NWIRP-Template.dotx"
    .\Build-NWIRPBuildingBlocks.ps1 -SourceDir "HM-NWIRP-Template_unpacked"
    .\Compress-WordPackage.ps1 -SourceDir "HM-NWIRP-Template_unpacked"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDir,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SourceDir -PathType Container)) {
    throw "SourceDir not found: $SourceDir"
}

$resolvedDir = (Resolve-Path -LiteralPath $SourceDir).ProviderPath

# ---------------------------------------------------------------------------
# Ensure glossary scaffolding exists
# ---------------------------------------------------------------------------

$glossaryDir = Join-Path $resolvedDir 'word\glossary'
$glossaryRelsDir = Join-Path $glossaryDir '_rels'
if (-not (Test-Path -LiteralPath $glossaryDir)) { New-Item -ItemType Directory -Path $glossaryDir | Out-Null }
if (-not (Test-Path -LiteralPath $glossaryRelsDir)) { New-Item -ItemType Directory -Path $glossaryRelsDir | Out-Null }

$glossaryStylesPath   = Join-Path $glossaryDir 'styles.xml'
$glossarySettingsPath = Join-Path $glossaryDir 'settings.xml'
$glossaryWebSettings  = Join-Path $glossaryDir 'webSettings.xml'
$glossaryFontTable    = Join-Path $glossaryDir 'fontTable.xml'
$glossaryDocPath      = Join-Path $glossaryDir 'document.xml'
$glossaryRelsPath     = Join-Path $glossaryRelsDir 'document.xml.rels'

function Ensure-File {
    param([string]$Path, [string]$Content)
    if (-not (Test-Path -LiteralPath $Path)) {
        Set-Content -LiteralPath $Path -Value $Content -Encoding UTF8
        Write-Host "  Created $Path" -ForegroundColor DarkGreen
    }
}

# Minimal valid glossary support files. If your main template's styles.xml
# is more elaborate, you can copy it over to glossary/styles.xml after
# running this script; Word treats glossary/styles.xml as overrides for
# the glossary content.
$minimalStyles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/></w:style>
</w:styles>
'@

$minimalSettings = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>
'@

$minimalWebSettings = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:optimizeForBrowser/>
</w:webSettings>
'@

$minimalFontTable = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Arial">
    <w:panose1 w:val="020B0604020202020204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E0002AFF" w:usb1="C0007843" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
</w:fonts>
'@

$minimalRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>
'@

Ensure-File -Path $glossaryStylesPath   -Content $minimalStyles
Ensure-File -Path $glossarySettingsPath -Content $minimalSettings
Ensure-File -Path $glossaryWebSettings  -Content $minimalWebSettings
Ensure-File -Path $glossaryFontTable    -Content $minimalFontTable
Ensure-File -Path $glossaryRelsPath     -Content $minimalRels

# ---------------------------------------------------------------------------
# Add FrontMatterHeader and AppendixHeader styles to word/styles.xml if missing.
# Both inherit from Heading1NoNumber (so they pick up its base behavior) but
# carry distinct NIST palette colors so they are visibly distinguishable.
# Tweak the colors/sizes in Word later if you want different appearances.
# ---------------------------------------------------------------------------

$mainStylesPath = Join-Path $resolvedDir 'word\styles.xml'
if (Test-Path -LiteralPath $mainStylesPath) {
    $stylesContent = Get-Content -LiteralPath $mainStylesPath -Raw
    $stylesToInject = ''

    if ($stylesContent -notmatch 'w:styleId="FrontMatterHeader"') {
        $stylesToInject += @'
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="FrontMatterHeader">
    <w:name w:val="Front Matter Header"/>
    <w:basedOn w:val="Heading1NoNumber"/>
    <w:next w:val="BodyText"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:pageBreakBefore/>
      <w:spacing w:before="240" w:after="120"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
      <w:b/>
      <w:color w:val="00257A"/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
'@
    }

    if ($stylesContent -notmatch 'w:styleId="AppendixHeader"') {
        $stylesToInject += @'
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="AppendixHeader">
    <w:name w:val="Appendix Header"/>
    <w:basedOn w:val="Heading1NoNumber"/>
    <w:next w:val="BodyText"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:pageBreakBefore/>
      <w:spacing w:before="240" w:after="120"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
      <w:b/>
      <w:color w:val="355195"/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
'@
    }

    if ($stylesToInject) {
        if ($stylesContent -notmatch '</w:styles>') {
            throw "word/styles.xml is malformed (no </w:styles>)."
        }
        $stylesContent = $stylesContent -replace '</w:styles>', ($stylesToInject + '</w:styles>')
        Set-Content -LiteralPath $mainStylesPath -Value $stylesContent -Encoding UTF8
        Write-Host "  Added FrontMatterHeader and/or AppendixHeader to word/styles.xml" -ForegroundColor DarkGreen
    }
} else {
    Write-Warning "word/styles.xml not found. The script cannot add FrontMatterHeader/AppendixHeader; building blocks that reference these will fall back to Normal."
}

# ---------------------------------------------------------------------------
# Define the building blocks
# ---------------------------------------------------------------------------

# OOXML helpers (return XML strings)
function P {
    param([string]$Style, [string]$Text)
    if ($Style) {
        return "<w:p><w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>"
    } else {
        return "<w:p><w:r><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>"
    }
}

function PI {
    # Paragraph with italic run
    param([string]$Style, [string]$Text)
    if ($Style) {
        return "<w:p><w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>"
    } else {
        return "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>"
    }
}

function PB {
    # Paragraph with bold run
    param([string]$Style, [string]$Text)
    if ($Style) {
        return "<w:p><w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>"
    } else {
        return "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>"
    }
}

function PageBreak {
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
}

# Table cell with optional shading and bold text
function TC {
    param([string]$Text, [string]$ShadingHex, [bool]$Bold = $false, [string]$ColorHex)
    $shd = if ($ShadingHex) { "<w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$ShadingHex`"/>" } else { '' }
    $rPr = ''
    if ($Bold -or $ColorHex) {
        $bits = ''
        if ($Bold) { $bits += '<w:b/>' }
        if ($ColorHex) { $bits += "<w:color w:val=`"$ColorHex`"/>" }
        $rPr = "<w:rPr>$bits</w:rPr>"
    }
    return "<w:tc><w:tcPr><w:tcW w:w=`"3000`" w:type=`"dxa`"/>$shd</w:tcPr><w:p><w:r>$rPr<w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p></w:tc>"
}

function TR {
    param([string]$Cells)
    return "<w:tr>$Cells</w:tr>"
}

function Table {
    param([string]$Rows, [int]$Cols = 2)
    $grid = ''
    for ($i=0; $i -lt $Cols; $i++) { $grid += '<w:gridCol w:w="3000"/>' }
    return @"
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders>
    <w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>
  <w:tblGrid>$grid</w:tblGrid>
  $Rows
</w:tbl>
"@
}

# Each block: Name, Gallery (quickParts | coverPg | autoTxt), Category, Description, Body
$blocks = @()

# --- PHASE 1: Front matter -------------------------------------------------

$blocks += @{
    Name = 'NCSTAR Cover Page'
    Gallery = 'coverPg'
    Category = 'NIST NCSTAR'
    Description = 'NCSTAR-style cover page with publication number, title, authors, date, and disclaimer.'
    Body = (PB '' 'NIST NCSTAR PUBLICATION NUMBER') + (PageBreak) + (PB '' '[Document Title]') + (P '' '[Subtitle]') + (P '' '[Author 1]') + (P '' '[Author 2]') + (P '' '[Affiliation]') + (P '' '[Month Year]') + (P 'Disclaimer' 'Certain commercial entities, equipment, products, or materials are identified in this document to describe a procedure or concept adequately. Such identification is not intended to imply recommendation, endorsement, or implication that the entities, products, materials, or equipment are necessarily the best available for the purpose.')
}

$blocks += @{
    Name = 'NIST Standard Disclaimer'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Standard NIST disclaimer text using the Disclaimer style.'
    Body = (P 'Disclaimer' 'Certain commercial entities, equipment, products, or materials are identified in this document to describe a procedure or concept adequately. Such identification is not intended to imply recommendation, endorsement, or implication that the entities, products, materials, or equipment are necessarily the best available for the purpose. Nor does such identification imply a determination of the level of quality of these products by the National Institute of Standards and Technology.')
}

$blocks += @{
    Name = 'Abstract Block'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Front-matter Abstract heading + body + Keywords line.'
    Body = (P 'FrontMatterHeader' 'Abstract') + (P 'BodyText' '[Abstract paragraph. Two to three sentences summarizing the report.]') + (PB 'BodyText' 'Keywords: ') + (P 'BodyText' '[keyword 1]; [keyword 2]; [keyword 3]')
}

$blocks += @{
    Name = 'Acknowledgments Block'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Front-matter Acknowledgments heading + body with funder placeholder.'
    Body = (P 'FrontMatterHeader' 'Acknowledgments') + (P 'BodyText' 'The authors gratefully acknowledge [funder] for support of this work under [grant or program]. We thank [collaborating institutions] for their participation. The opinions, findings, conclusions, and recommendations expressed in this publication are those of the authors and do not necessarily reflect the views of [funder].')
}

# --- PHASE 2: Chapter structure --------------------------------------------

$blocks += @{
    Name = 'Chapter Opener'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Page break + numbered chapter title (Heading1) + tagline.'
    Body = (PageBreak) + (P 'Heading1' '[Chapter Title]') + (PI 'BodyText' '[One-sentence chapter tagline.]')
}

$blocks += @{
    Name = 'Section Divider'
    Gallery = 'autoTxt'
    Category = 'NIST NCSTAR'
    Description = 'Page break + Heading 1 placeholder for major section break.'
    Body = (PageBreak) + (P 'Heading1' '[Section Title]')
}

# --- PHASE 3: Recurring research artifacts ---------------------------------

$blocks += @{
    Name = 'Figure Caption Block'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Figure caption + alt-text reminder + source line.'
    Body = (P 'Caption' 'Figure [N-N]. [Figure title.]') + (PI 'TableandFigureFootnotes' 'Source: [data source citation]') + (P 'BodyText' '[Insert figure here. ALT TEXT: describe the figure for screen readers.]')
}

$blocks += @{
    Name = 'Table Caption Block'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Table caption above + source line below for tables.'
    Body = (P 'Caption' 'Table [N-N]. [Table title.]') + (P 'BodyText' '[Insert table here.]') + (PI 'TableandFigureFootnotes' 'Source: [data source citation]')
}

$blocks += @{
    Name = 'Methodology Box'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Bordered single-cell callout for methodology side-explanations.'
    Body = (Table -Cols 1 -Rows (TR (TC 'Methods Note' '#CCD3E4' $true)) ) + (P 'BodyText' '[Brief methods explanation that does not fit the main narrative.]')
}

$blocks += @{
    Name = 'Key Findings Sidebar'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Bordered callout box for highlighted findings.'
    Body = (Table -Cols 1 -Rows ((TR (TC 'Key Findings' '#00257A' $true '#FFFFFF')) + (TR (TC '[Finding 1]' '' $false)) + (TR (TC '[Finding 2]' '' $false)) + (TR (TC '[Finding 3]' '' $false)) ))
}

$blocks += @{
    Name = 'Interview Quote Block'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Indented italicized interview excerpt with attribution.'
    Body = (PI 'BodyText' '"[Interview quote text...]"') + (P 'BodyText' '[Attribution, e.g., Hospital administrator, San Juan]')
}

# --- PHASE 4: Schools/Hospitals comparison ---------------------------------

$tawnyHdr = (TC 'Schools' '#D55E00' $true '#FFFFFF')
$shamrockHdr = (TC 'Hospitals' '#009E73' $true '#FFFFFF')
$blocks += @{
    Name = 'Schools-Hospitals Comparison'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Two-column comparison table with NIST entity colors (Tawny + Shamrock Green).'
    Body = (Table -Cols 2 -Rows ((TR ($tawnyHdr + $shamrockHdr)) + (TR ((TC '[Schools row 1]' '') + (TC '[Hospitals row 1]' ''))) + (TR ((TC '[Schools row 2]' '') + (TC '[Hospitals row 2]' ''))) + (TR ((TC '[Schools row 3]' '') + (TC '[Hospitals row 3]' ''))) ))
}

$blocks += @{
    Name = 'Demographic Summary Table'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Three-row demographic summary (Schools, Hospitals, Total) with sample size and region columns.'
    Body = (P 'Caption' 'Table [N-N]. Sample demographics.') + (Table -Cols 4 -Rows ((TR ((TC 'Group' '#00257A' $true '#FFFFFF') + (TC 'N' '#00257A' $true '#FFFFFF') + (TC 'Region' '#00257A' $true '#FFFFFF') + (TC 'Notes' '#00257A' $true '#FFFFFF'))) + (TR ((TC 'Schools' '#D55E00' $true '#FFFFFF') + (TC '[N]' '') + (TC '[Regions]' '') + (TC '[Notes]' ''))) + (TR ((TC 'Hospitals' '#009E73' $true '#FFFFFF') + (TC '[N]' '') + (TC '[Regions]' '') + (TC '[Notes]' ''))) + (TR ((TC 'Total' '#677CAF' $true '#FFFFFF') + (TC '[N]' '') + (TC '[Regions]' '') + (TC '[Notes]' ''))) ))
}

$blocks += @{
    Name = 'Recovery Status Legend'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Inline legend showing the four-point recovery scale with colored swatches.'
    Body = (Table -Cols 4 -Rows (TR ((TC 'Not Started' '#990B38' $true '#FFFFFF') + (TC 'In Progress' '#E69F00' $true '#FFFFFF') + (TC 'Substantially Recovered' '#56B4E9' $true '#FFFFFF') + (TC 'Fully Recovered' '#009E73' $true '#FFFFFF'))) )
}

# --- PHASE 5: Recovery model results table ---------------------------------

$blocks += @{
    Name = 'Recovery Model Results Table'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Pre-formatted recovery model results table modeled on Ch05 result tables.'
    Body = (P 'Caption' 'Table [N-N]. [Recovery model] results.') + (Table -Cols 5 -Rows ((TR ((TC 'Variable' '#00257A' $true '#FFFFFF') + (TC 'M (SD)' '#00257A' $true '#FFFFFF') + (TC 'n' '#00257A' $true '#FFFFFF') + (TC 'p' '#00257A' $true '#FFFFFF') + (TC 'Effect' '#00257A' $true '#FFFFFF'))) + (TR ((TC '[Variable 1]' '') + (TC '' '') + (TC '' '') + (TC '' '') + (TC '' ''))) + (TR ((TC '[Variable 2]' '') + (TC '' '') + (TC '' '') + (TC '' '') + (TC '' ''))) + (TR ((TC '[Variable 3]' '') + (TC '' '') + (TC '' '') + (TC '' '') + (TC '' ''))) ))
}

# --- PHASE 6: End matter and stats -----------------------------------------

$blocks += @{
    Name = 'References Section Opener'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Page break + References heading + first entry placeholder.'
    Body = (PageBreak) + (P 'AppendixHeader' 'References') + (P 'References' '[Author, A. (Year). Title of work. Publisher or journal, vol(issue), pages. URL or DOI.]')
}

$blocks += @{
    Name = 'Bibliography Entry'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Single bibliography entry placeholder using References style.'
    Body = (P 'References' '[Author, A. (Year). Title of work. Publisher or journal, vol(issue), pages. URL or DOI.]')
}

$blocks += @{
    Name = 'Appendix Opener'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Page break + Appendix heading.'
    Body = (PageBreak) + (P 'AppendixHeader' '[Appendix A. Appendix Title]') + (P 'BodyText' '[Appendix introduction.]')
}

$blocks += @{
    Name = 'Acronyms Appendix Table'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Two-column acronym/definition table seeded with NIST and HM project acronyms.'
    Body = (P 'AppendixHeader' 'List of Acronyms and Abbreviations') + (Table -Cols 2 -Rows ((TR ((TC 'Acronym' '#00257A' $true '#FFFFFF') + (TC 'Definition' '#00257A' $true '#FFFFFF'))) + (TR ((TC 'NIST' '') + (TC 'National Institute of Standards and Technology' ''))) + (TR ((TC 'NCSTAR' '') + (TC 'National Construction Safety Team Act Report' ''))) + (TR ((TC 'NCST' '') + (TC 'National Construction Safety Team' ''))) + (TR ((TC 'NWIRP' '') + (TC '[Define NWIRP for your context]' ''))) + (TR ((TC 'FEMA' '') + (TC 'Federal Emergency Management Agency' ''))) + (TR ((TC 'FIPS' '') + (TC 'Federal Information Processing Standards' ''))) + (TR ((TC 'PR' '') + (TC 'Puerto Rico' ''))) ))
}

$blocks += @{
    Name = 'Stats Inline Result'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Inline statistical reporting block (M, SD, n).'
    Body = "<w:p><w:r><w:t xml:space=`"preserve`">(</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>M</w:t></w:r><w:r><w:t xml:space=`"preserve`"> = , </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>SD</w:t></w:r><w:r><w:t xml:space=`"preserve`"> = , </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>n</w:t></w:r><w:r><w:t xml:space=`"preserve`"> = )</w:t></w:r></w:p>"
}

$blocks += @{
    Name = 'Effect Size Citation'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Inline effect size + significance citation.'
    Body = "<w:p><w:r><w:t xml:space=`"preserve`">(</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>r</w:t></w:r><w:r><w:t xml:space=`"preserve`"> = , </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>p</w:t></w:r><w:r><w:t xml:space=`"preserve`"> &lt; .05)</w:t></w:r></w:p>"
}

$blocks += @{
    Name = 'Sample Size Statement'
    Gallery = 'quickParts'
    Category = 'NIST NCSTAR'
    Description = 'Standard sample-size and response-rate sentence.'
    Body = (P 'BodyText' 'Of [N] facilities surveyed, [n] responded ([X]%).')
}

# ---------------------------------------------------------------------------
# Read existing glossary/document.xml or create a new one, then append docParts
# ---------------------------------------------------------------------------

# Build the docParts XML
$docPartsXml = New-Object System.Text.StringBuilder

foreach ($b in $blocks) {
    $guid = ([guid]::NewGuid().ToString('B').ToUpper())  # {ABCDEF...}
    $name = [System.Web.HttpUtility]::HtmlEncode($b.Name)
    $cat  = [System.Web.HttpUtility]::HtmlEncode($b.Category)
    $desc = [System.Web.HttpUtility]::HtmlEncode($b.Description)
    $gallery = $b.Gallery

    $entry = @"
<w:docPart>
  <w:docPartPr>
    <w:name w:val="$name"/>
    <w:category>
      <w:name w:val="$cat"/>
      <w:gallery w:val="$gallery"/>
    </w:category>
    <w:types>
      <w:type w:val="normal"/>
    </w:types>
    <w:behaviors>
      <w:behavior w:val="content"/>
    </w:behaviors>
    <w:description w:val="$desc"/>
    <w:guid w:val="$guid"/>
  </w:docPartPr>
  <w:docPartBody>
    $($b.Body)
    <w:p/>
  </w:docPartBody>
</w:docPart>
"@
    [void]$docPartsXml.Append($entry)
}

Add-Type -AssemblyName System.Web | Out-Null

if (Test-Path -LiteralPath $glossaryDocPath) {
    if (-not $Force) {
        $existing = Get-Content -LiteralPath $glossaryDocPath -Raw
        $existingCount = ([regex]::Matches($existing, '<w:docPart>')).Count
        Write-Host "Existing glossary document.xml has $existingCount docParts." -ForegroundColor Yellow
        $reply = Read-Host "Append the $($blocks.Count) new building blocks to it? (y/N)"
        if ($reply -notmatch '^[yY]') { throw "Aborted." }
    }
    # Append by inserting before </w:docParts>
    $existing = Get-Content -LiteralPath $glossaryDocPath -Raw
    if ($existing -notmatch '</w:docParts>') {
        throw "Existing glossary/document.xml does not contain </w:docParts>. Cannot safely append."
    }
    $newContent = $existing -replace '</w:docParts>', ($docPartsXml.ToString() + '</w:docParts>')
    Set-Content -LiteralPath $glossaryDocPath -Value $newContent -Encoding UTF8
    Write-Host "Appended $($blocks.Count) building blocks to existing glossary." -ForegroundColor Green
} else {
    # Create a new glossary document.xml
    $newDoc = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:glossaryDocument
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docParts>
    $($docPartsXml.ToString())
  </w:docParts>
</w:glossaryDocument>
"@
    Set-Content -LiteralPath $glossaryDocPath -Value $newDoc -Encoding UTF8
    Write-Host "Created new glossary with $($blocks.Count) building blocks." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Update [Content_Types].xml to register glossary parts (if not registered)
# ---------------------------------------------------------------------------

$ctPath = Join-Path $resolvedDir '[Content_Types].xml'
if (-not (Test-Path -LiteralPath $ctPath)) { throw "[Content_Types].xml not found at $ctPath." }

$ct = Get-Content -LiteralPath $ctPath -Raw
$ctOriginal = $ct

$overrides = @(
    @{ Part = '/word/glossary/document.xml';     Type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml' }
    @{ Part = '/word/glossary/styles.xml';        Type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' }
    @{ Part = '/word/glossary/settings.xml';      Type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml' }
    @{ Part = '/word/glossary/webSettings.xml';   Type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml' }
    @{ Part = '/word/glossary/fontTable.xml';     Type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml' }
)

$injection = ''
foreach ($o in $overrides) {
    $needle = "PartName=`"$($o.Part)`""
    if ($ct -notlike "*$needle*") {
        $injection += "<Override PartName=`"$($o.Part)`" ContentType=`"$($o.Type)`"/>"
    }
}

if ($injection) {
    if ($ct -notmatch '</Types>') { throw "[Content_Types].xml is malformed (no </Types>)." }
    $ct = $ct -replace '</Types>', ($injection + '</Types>')
    Set-Content -LiteralPath $ctPath -Value $ct -Encoding UTF8
    Write-Host "Registered glossary content types in [Content_Types].xml" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Update word/_rels/document.xml.rels to add glossary relationship if missing
# ---------------------------------------------------------------------------

$docRelsPath = Join-Path $resolvedDir 'word\_rels\document.xml.rels'
if (Test-Path -LiteralPath $docRelsPath) {
    $rels = Get-Content -LiteralPath $docRelsPath -Raw
    if ($rels -notlike '*Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument"*') {
        # Pick a non-conflicting Id
        $existingIds = [regex]::Matches($rels, 'Id="(rId\d+)"') | ForEach-Object { [int]($_.Groups[1].Value -replace 'rId','') }
        $nextId = if ($existingIds) { ($existingIds | Measure-Object -Maximum).Maximum + 1 } else { 1 }
        $relEntry = "<Relationship Id=`"rId$nextId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument`" Target=`"glossary/document.xml`"/>"
        $rels = $rels -replace '</Relationships>', ($relEntry + '</Relationships>')
        Set-Content -LiteralPath $docRelsPath -Value $rels -Encoding UTF8
        Write-Host "Added glossary relationship to word/_rels/document.xml.rels" -ForegroundColor Green
    }
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "Building blocks added:" -ForegroundColor Cyan
$blocks | Group-Object Gallery | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-12} {1}" -f $_.Name, $_.Count)
}
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Compress-WordPackage.ps1 -SourceDir `"$SourceDir`""
Write-Host "  2. Open the resulting .dotx in Word."
Write-Host "  3. Insert > Quick Parts > Building Blocks Organizer to verify the blocks appear under the 'NIST NCSTAR' category."
Write-Host "  4. Insert any block to test it; modify the inserted content; right-click and 'Save Selection to Quick Part Gallery' to evolve it."
