<#
.SYNOPSIS
    Builds a single .docx file containing all 22 NIST NCSTAR / NWIRP snippets
    as labeled sections, ready to copy and paste from.

.DESCRIPTION
    Instead of installing snippets as Quick Parts in a template (which is
    finicky), this produces a regular Word document with every snippet
    included inline. Each snippet has a heading with its name and a brief
    description, the snippet content below it, and a page break before
    the next one.

    Open the resulting .docx in Word, navigate by the heading list
    (View > Navigation Pane), copy whatever block you need into your
    chapter document, paste, and tweak.

    Uses your master template's styles.xml, theme, and fonts so the
    snippets look the way they will when you paste into a real chapter.

.PARAMETER SourceTemplate
    Path to the master template (.dotx, .docx, etc.) whose styles and
    theme should be used. The template is copied (not modified).

.PARAMETER OutputFile
    Where to write the output .docx. Defaults to a sibling file named
    "<template-basename>-Snippets.docx".

.PARAMETER Force
    Overwrite OutputFile if it exists.

.EXAMPLE
    .\Build-NWIRPSnippetsDoc.ps1 -SourceTemplate "C:\templates\HM-NWIRP-Template.dotx"

.EXAMPLE
    .\Build-NWIRPSnippetsDoc.ps1 -SourceTemplate "C:\t\HM-NWIRP-Template.dotx" -OutputFile "C:\t\snippets-2026-04.docx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceTemplate,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SourceTemplate -PathType Leaf)) {
    throw "SourceTemplate not found: $SourceTemplate"
}

$resolvedSrc = (Resolve-Path -LiteralPath $SourceTemplate).ProviderPath
if ($resolvedSrc -notmatch '\.(docx|dotx|docm|dotm)$') {
    throw "SourceTemplate must be .docx, .dotx, .docm, or .dotm. Got: $resolvedSrc"
}

$srcDir   = [System.IO.Path]::GetDirectoryName($resolvedSrc)
$srcBase  = [System.IO.Path]::GetFileNameWithoutExtension($resolvedSrc)

if (-not $OutputFile) {
    $OutputFile = Join-Path $srcDir ($srcBase + '-Snippets.docx')
}

if ((Test-Path -LiteralPath $OutputFile) -and -not $Force) {
    $reply = Read-Host "Overwrite $OutputFile`? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

# ---------------------------------------------------------------------------
# Snippet content. Same content as Build-NWIRPBuildingBlocks but emitted as
# labeled sections in document.xml instead of glossary docPart entries.
# ---------------------------------------------------------------------------

function P {
    param([string]$Style, [string]$Text)
    if ($Style) { "<w:p><w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>" }
    else        { "<w:p><w:r><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>" }
}
function PI {
    param([string]$Style, [string]$Text)
    if ($Style) { "<w:p><w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>" }
    else        { "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>" }
}
function PB {
    param([string]$Style, [string]$Text)
    if ($Style) { "<w:p><w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>" }
    else        { "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p>" }
}
function PageBreak { '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' }

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
    "<w:tc><w:tcPr><w:tcW w:w=`"3000`" w:type=`"dxa`"/>$shd</w:tcPr><w:p><w:r>$rPr<w:t xml:space=`"preserve`">$Text</w:t></w:r></w:p></w:tc>"
}
function TR { param([string]$Cells) "<w:tr>$Cells</w:tr>" }
function Table {
    param([string]$Rows, [int]$Cols = 2)
    $grid = ''
    for ($i=0; $i -lt $Cols; $i++) { $grid += '<w:gridCol w:w="3000"/>' }
    @"
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
    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>
  <w:tblGrid>$grid</w:tblGrid>
  $Rows
</w:tbl>
<w:p/>
"@
}

$snippets = @()

$snippets += @{ Name='NCSTAR Cover Page'; Description='Cover page with publication number, title, authors, date, disclaimer.';
    Body = (PB '' 'NIST NCSTAR PUBLICATION NUMBER') + (PageBreak) + (PB '' '[Document Title]') + (P '' '[Subtitle]') + (P '' '[Author 1]') + (P '' '[Author 2]') + (P '' '[Affiliation]') + (P '' '[Month Year]') + (P 'Disclaimer' 'Certain commercial entities, equipment, products, or materials are identified in this document to describe a procedure or concept adequately. Such identification is not intended to imply recommendation, endorsement, or implication that the entities, products, materials, or equipment are necessarily the best available for the purpose.') }

$snippets += @{ Name='NIST Standard Disclaimer'; Description='Standard NIST disclaimer text using the Disclaimer style.';
    Body = (P 'Disclaimer' 'Certain commercial entities, equipment, products, or materials are identified in this document to describe a procedure or concept adequately. Such identification is not intended to imply recommendation, endorsement, or implication that the entities, products, materials, or equipment are necessarily the best available for the purpose. Nor does such identification imply a determination of the level of quality of these products by the National Institute of Standards and Technology.') }

$snippets += @{ Name='Abstract Block'; Description='Front-matter Abstract heading + body + Keywords.';
    Body = (P 'FrontMatterHeader' 'Abstract') + (P 'BodyText' '[Abstract paragraph. Two to three sentences summarizing the report.]') + (PB 'BodyText' 'Keywords: ') + (P 'BodyText' '[keyword 1]; [keyword 2]; [keyword 3]') }

$snippets += @{ Name='Acknowledgments Block'; Description='Front-matter Acknowledgments with funder placeholder.';
    Body = (P 'FrontMatterHeader' 'Acknowledgments') + (P 'BodyText' 'The authors gratefully acknowledge [funder] for support of this work under [grant or program]. We thank [collaborating institutions] for their participation. The opinions, findings, conclusions, and recommendations expressed in this publication are those of the authors and do not necessarily reflect the views of [funder].') }

$snippets += @{ Name='Chapter Opener'; Description='Page break + numbered chapter title (Heading 1) + tagline.';
    Body = (PageBreak) + (P 'Heading1' '[Chapter Title]') + (PI 'BodyText' '[One-sentence chapter tagline.]') }

$snippets += @{ Name='Section Divider'; Description='Page break + Heading 1 placeholder for major section break.';
    Body = (PageBreak) + (P 'Heading1' '[Section Title]') }

$snippets += @{ Name='Figure Caption Block'; Description='Figure caption + alt-text reminder + source.';
    Body = (P 'Caption' 'Figure [N-N]. [Figure title.]') + (PI 'TableandFigureFootnotes' 'Source: [data source citation]') + (P 'BodyText' '[Insert figure here. ALT TEXT: describe the figure for screen readers.]') }

$snippets += @{ Name='Table Caption Block'; Description='Table caption above + source line below.';
    Body = (P 'Caption' 'Table [N-N]. [Table title.]') + (P 'BodyText' '[Insert table here.]') + (PI 'TableandFigureFootnotes' 'Source: [data source citation]') }

$snippets += @{ Name='Methodology Box'; Description='Bordered single-cell callout for methodology side-explanations.';
    Body = (Table -Cols 1 -Rows (TR (TC 'Methods Note' '#CCD3E4' $true)) ) + (P 'BodyText' '[Brief methods explanation that does not fit the main narrative.]') }

$snippets += @{ Name='Key Findings Sidebar'; Description='Bordered callout box for highlighted findings.';
    Body = (Table -Cols 1 -Rows ((TR (TC 'Key Findings' '#00257A' $true '#FFFFFF')) + (TR (TC '[Finding 1]' '' $false)) + (TR (TC '[Finding 2]' '' $false)) + (TR (TC '[Finding 3]' '' $false)) )) }

$snippets += @{ Name='Interview Quote Block'; Description='Indented italicized interview excerpt with attribution.';
    Body = (PI 'BodyText' '"[Interview quote text...]"') + (P 'BodyText' '[Attribution, e.g., Hospital administrator, San Juan]') }

$snippets += @{ Name='Schools-Hospitals Comparison'; Description='Two-column comparison with NIST entity colors (Tawny + Shamrock Green).';
    Body = (Table -Cols 2 -Rows ((TR ((TC 'Schools' '#D55E00' $true '#FFFFFF') + (TC 'Hospitals' '#009E73' $true '#FFFFFF'))) + (TR ((TC '[Schools row 1]' '') + (TC '[Hospitals row 1]' ''))) + (TR ((TC '[Schools row 2]' '') + (TC '[Hospitals row 2]' ''))) + (TR ((TC '[Schools row 3]' '') + (TC '[Hospitals row 3]' ''))) )) }

$snippets += @{ Name='Demographic Summary Table'; Description='Three-row demographic summary (Schools, Hospitals, Total).';
    Body = (P 'Caption' 'Table [N-N]. Sample demographics.') + (Table -Cols 4 -Rows ((TR ((TC 'Group' '#00257A' $true '#FFFFFF') + (TC 'N' '#00257A' $true '#FFFFFF') + (TC 'Region' '#00257A' $true '#FFFFFF') + (TC 'Notes' '#00257A' $true '#FFFFFF'))) + (TR ((TC 'Schools' '#D55E00' $true '#FFFFFF') + (TC '[N]' '') + (TC '[Regions]' '') + (TC '[Notes]' ''))) + (TR ((TC 'Hospitals' '#009E73' $true '#FFFFFF') + (TC '[N]' '') + (TC '[Regions]' '') + (TC '[Notes]' ''))) + (TR ((TC 'Total' '#677CAF' $true '#FFFFFF') + (TC '[N]' '') + (TC '[Regions]' '') + (TC '[Notes]' ''))) )) }

$snippets += @{ Name='Recovery Status Legend'; Description='Inline legend showing the four-point recovery scale with colored swatches.';
    Body = (Table -Cols 4 -Rows (TR ((TC 'Not Started' '#990B38' $true '#FFFFFF') + (TC 'In Progress' '#E69F00' $true '#FFFFFF') + (TC 'Substantially Recovered' '#56B4E9' $true '#FFFFFF') + (TC 'Fully Recovered' '#009E73' $true '#FFFFFF'))) ) }

$snippets += @{ Name='Recovery Model Results Table'; Description='Pre-formatted recovery model results table.';
    Body = (P 'Caption' 'Table [N-N]. [Recovery model] results.') + (Table -Cols 5 -Rows ((TR ((TC 'Variable' '#00257A' $true '#FFFFFF') + (TC 'M (SD)' '#00257A' $true '#FFFFFF') + (TC 'n' '#00257A' $true '#FFFFFF') + (TC 'p' '#00257A' $true '#FFFFFF') + (TC 'Effect' '#00257A' $true '#FFFFFF'))) + (TR ((TC '[Variable 1]' '') + (TC '' '') + (TC '' '') + (TC '' '') + (TC '' ''))) + (TR ((TC '[Variable 2]' '') + (TC '' '') + (TC '' '') + (TC '' '') + (TC '' ''))) + (TR ((TC '[Variable 3]' '') + (TC '' '') + (TC '' '') + (TC '' '') + (TC '' ''))) )) }

$snippets += @{ Name='References Section Opener'; Description='Page break + References heading + first entry placeholder.';
    Body = (PageBreak) + (P 'AppendixHeader' 'References') + (P 'References' '[Author, A. (Year). Title of work. Publisher or journal, vol(issue), pages. URL or DOI.]') }

$snippets += @{ Name='Bibliography Entry'; Description='Single bibliography entry placeholder.';
    Body = (P 'References' '[Author, A. (Year). Title of work. Publisher or journal, vol(issue), pages. URL or DOI.]') }

$snippets += @{ Name='Appendix Opener'; Description='Page break + Appendix heading.';
    Body = (PageBreak) + (P 'AppendixHeader' '[Appendix A. Appendix Title]') + (P 'BodyText' '[Appendix introduction.]') }

$snippets += @{ Name='Acronyms Appendix Table'; Description='Two-column acronym/definition table seeded with NIST and HM project acronyms.';
    Body = (P 'AppendixHeader' 'List of Acronyms and Abbreviations') + (Table -Cols 2 -Rows ((TR ((TC 'Acronym' '#00257A' $true '#FFFFFF') + (TC 'Definition' '#00257A' $true '#FFFFFF'))) + (TR ((TC 'NIST' '') + (TC 'National Institute of Standards and Technology' ''))) + (TR ((TC 'NCSTAR' '') + (TC 'National Construction Safety Team Act Report' ''))) + (TR ((TC 'NCST' '') + (TC 'National Construction Safety Team' ''))) + (TR ((TC 'NWIRP' '') + (TC '[Define NWIRP for your context]' ''))) + (TR ((TC 'FEMA' '') + (TC 'Federal Emergency Management Agency' ''))) + (TR ((TC 'FIPS' '') + (TC 'Federal Information Processing Standards' ''))) + (TR ((TC 'PR' '') + (TC 'Puerto Rico' ''))) )) }

$snippets += @{ Name='Stats Inline Result'; Description='Inline statistical reporting block (M, SD, n).';
    Body = "<w:p><w:r><w:t xml:space=`"preserve`">(</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>M</w:t></w:r><w:r><w:t xml:space=`"preserve`"> = , </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>SD</w:t></w:r><w:r><w:t xml:space=`"preserve`"> = , </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>n</w:t></w:r><w:r><w:t xml:space=`"preserve`"> = )</w:t></w:r></w:p>" }

$snippets += @{ Name='Effect Size Citation'; Description='Inline effect size + significance citation.';
    Body = "<w:p><w:r><w:t xml:space=`"preserve`">(</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>r</w:t></w:r><w:r><w:t xml:space=`"preserve`"> = , </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>p</w:t></w:r><w:r><w:t xml:space=`"preserve`"> &lt; .05)</w:t></w:r></w:p>" }

$snippets += @{ Name='Sample Size Statement'; Description='Standard sample-size and response-rate sentence.';
    Body = (P 'BodyText' 'Of [N] facilities surveyed, [n] responded ([X]%).') }

# ---------------------------------------------------------------------------
# Build the document.xml body: each snippet as a labeled section
# ---------------------------------------------------------------------------

$bodyXml = New-Object System.Text.StringBuilder
[void]$bodyXml.Append((P 'Heading1' 'NIST NCSTAR Snippets Reference'))
[void]$bodyXml.Append((P 'BodyText' 'This document contains 22 reusable snippets for NIST NCSTAR / NWIRP work. Open the Navigation Pane (View > Navigation Pane), pick a snippet, copy what you need, paste into your chapter document.'))
[void]$bodyXml.Append((PageBreak))

$index = 1
foreach ($s in $snippets) {
    [void]$bodyXml.Append((P 'Heading1' "$index. $($s.Name)"))
    [void]$bodyXml.Append((PI 'BodyText' $s.Description))
    [void]$bodyXml.Append('<w:p/>')
    [void]$bodyXml.Append($s.Body)
    [void]$bodyXml.Append((PageBreak))
    $index++
}

# ---------------------------------------------------------------------------
# Build the new document.xml
# ---------------------------------------------------------------------------

$newDocumentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    $($bodyXml.ToString())
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

# ---------------------------------------------------------------------------
# Build the output zip: copy entries from source template, replace
# document.xml, change Content_Types, drop glossary if present
# ---------------------------------------------------------------------------

Write-Host "Building snippets doc from template: $resolvedSrc" -ForegroundColor Cyan

$tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "snippets-$([guid]::NewGuid()).zip")

$srcStream = [System.IO.File]::OpenRead($resolvedSrc)
try {
    $srcZip = New-Object System.IO.Compression.ZipArchive($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $outStream = [System.IO.File]::Create($tempZip)
        try {
            $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($entry in $srcZip.Entries) {
                    $name = $entry.FullName

                    # Skip glossary parts entirely (we are not using them in this output)
                    if ($name -like 'word/glossary/*') { continue }

                    if ($name -eq 'word/document.xml') {
                        # Replace with our new body
                        $newEntry = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                        $w = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
                        try { $w.Write($newDocumentXml) } finally { $w.Dispose() }
                        continue
                    }

                    if ($name -eq '[Content_Types].xml') {
                        # Read, edit, re-write: change template type to document type, drop glossary overrides
                        $r = New-Object System.IO.StreamReader($entry.Open())
                        try { $ct = $r.ReadToEnd() } finally { $r.Dispose() }

                        # Change the main document override from template to document content type
                        $ct = $ct -replace 'application/vnd\.openxmlformats-officedocument\.wordprocessingml\.template\.main\+xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
                        # Strip glossary overrides
                        $ct = $ct -replace '<Override\s+PartName="/word/glossary/[^"]+"\s+ContentType="[^"]+"\s*/>', ''

                        $newEntry = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                        $w = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
                        try { $w.Write($ct) } finally { $w.Dispose() }
                        continue
                    }

                    if ($name -eq 'word/_rels/document.xml.rels') {
                        # Strip any glossary relationship
                        $r = New-Object System.IO.StreamReader($entry.Open())
                        try { $rels = $r.ReadToEnd() } finally { $r.Dispose() }
                        $rels = $rels -replace '<Relationship\s+Id="[^"]+"\s+Type="[^"]*glossaryDocument[^"]*"\s+Target="[^"]+"\s*/>', ''

                        $newEntry = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                        $w = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
                        try { $w.Write($rels) } finally { $w.Dispose() }
                        continue
                    }

                    # Default: copy verbatim
                    $newEntry = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                    $entryStream = $entry.Open()
                    try {
                        $newStream = $newEntry.Open()
                        try { $entryStream.CopyTo($newStream) } finally { $newStream.Dispose() }
                    } finally { $entryStream.Dispose() }
                }
            } finally { $outZip.Dispose() }
        } finally { $outStream.Dispose() }
    } finally { $srcZip.Dispose() }
} finally { $srcStream.Dispose() }

# Validate
$validateOk = $true
try {
    $vfs = [System.IO.File]::OpenRead($tempZip)
    try {
        $vzip = New-Object System.IO.Compression.ZipArchive($vfs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($p in @('[Content_Types].xml','_rels/.rels','word/document.xml')) {
                if (-not $vzip.GetEntry($p)) { $validateOk = $false; break }
            }
        } finally { $vzip.Dispose() }
    } finally { $vfs.Dispose() }
} catch { $validateOk = $false }

if (-not $validateOk) {
    Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue
    throw "Validation failed: rebuilt zip is missing required parts."
}

if (Test-Path -LiteralPath $OutputFile) { Remove-Item -LiteralPath $OutputFile -Force }
Move-Item -LiteralPath $tempZip -Destination $OutputFile -Force

Write-Host ""
Write-Host "Wrote: $OutputFile" -ForegroundColor Green
Write-Host "Open in Word. View > Navigation Pane gives you the snippet index." -ForegroundColor DarkGray
