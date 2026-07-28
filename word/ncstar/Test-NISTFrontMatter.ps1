<#
.SYNOPSIS
    Audits a Chapter-00 front matter file against the canonical NIST
    front-matter sequence for either a Project Report or a Summary
    Report. Reports missing sections, extras, and out-of-order items.

.DESCRIPTION
    The assembly script (Build-NISTReport.ps1) does not impose front
    matter ordering. Whatever sequence Chapter-00 has is what lands in
    the combined deliverable. This script audits Chapter-00 against the
    canonical sequence so drift is caught before assembly.

    Canonical sequences (encoded in this script):

    Project Report:
        1.  Cover Page
        2.  Title Page
        3.  Disclaimers
        4.  Dedication
        5.  Abstract and Keywords
        6.  Table of Contents
        7.  List of Figures
        8.  List of Tables
        9.  Preface
        10. Acknowledgments
        11. Chapter Authors

    Summary Report:
        1.  Cover Page
        2.  Title Page
        3.  Disclaimers
        4.  National Construction Safety Team
        5.  Contributors to the Investigation
        6.  Dedication
        7.  Abstract and Keywords
        8.  Table of Contents
        9.  List of Figures
        10. List of Tables
        11. List of Acronyms and Abbreviations
        12. Preface
        13. Acknowledgments
        14. Executive Summary

    Cover Page is detected by the Title or HMCoverDate paragraph styles
    at the top of the document, or by the presence of the HM-Cover table
    style if the cover is purely table-based.

    Exit codes:
        0 = clean (every expected landmark present in the right order,
            no unexpected landmarks)
        1 = drift (missing, extra, or out-of-order landmarks)
        2 = input error (file not found, etc.)

.PARAMETER File
    Path to a single Chapter-00 docx. Either -File or -Folder is required.

.PARAMETER Folder
    Folder containing the report files. Use with -Series.

.PARAMETER Series
    Series prefix (e.g. NCSTAR_4-3, NWIRP_2-3). Used with -Folder to
    locate the Chapter-00 file.

.PARAMETER ReportType
    Project | Summary. Required.

.PARAMETER Json
    Emit the audit result as JSON on stdout instead of the colored table.

.EXAMPLE
    .\Test-NISTFrontMatter.ps1 -Folder "C:\Users\Admin\Documents\260512" `
        -Series NCSTAR_4-3 -ReportType Project

.EXAMPLE
    .\Test-NISTFrontMatter.ps1 -File ".\NWIRP_2-3_Chapter-00.docx" `
        -ReportType Summary -Json
#>

[CmdletBinding(DefaultParameterSetName='Folder')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Series,

    [Parameter(Mandatory=$true)]
    [ValidateSet('Project','Summary')]
    [string]$ReportType,

    [switch]$Json
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Resolve target file
# ---------------------------------------------------------------------------

if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) {
        Write-Error "File not found: $File"
        exit 2
    }
    $targetFile = (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
        Write-Error "Folder not found: $Folder"
        exit 2
    }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $seriesEsc = [regex]::Escape($Series)
    $candidate = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object { $_.Extension -ieq '.docx' -and $_.Name -imatch "^${seriesEsc}\W.*Chapter-0+(?:\W|$)" } |
        Select-Object -First 1
    if (-not $candidate) {
        Write-Error "No '${Series}*Chapter-00*.docx' found in $folderResolved"
        exit 2
    }
    $targetFile = $candidate.FullName
}

# ---------------------------------------------------------------------------
# Canonical sequences
# ---------------------------------------------------------------------------
# Each landmark has:
#   Key      - stable identifier
#   Label    - human-readable name
#   Patterns - case-insensitive regex patterns matched against paragraph text
#   Optional - if $true, do not report MISSING for this landmark
# ---------------------------------------------------------------------------

$landmarkDefs = [ordered]@{
    'CoverPage'         = @{ Label = 'Cover Page'                       ; Patterns = @() }   # detected by Title / HMCoverDate styles or HM-Cover table
    'TitlePage'         = @{ Label = 'Title Page'                       ; Patterns = @() }   # detected by Series / ReportTitle / ReportSubtitle / Authors styles
    'Disclaimers'       = @{ Label = 'Disclaimers'                      ; Patterns = @('commercial entities', 'International System of Units', '^Disclaimer') }
    'NCSTRoster'        = @{ Label = 'National Construction Safety Team'; Patterns = @('^National Construction Safety Team\s*$') }
    'Contributors'      = @{ Label = 'Contributors to the Investigation'; Patterns = @('^Contributors to the Investigation\s*$','^Contributors\s*$') }
    'Dedication'        = @{ Label = 'Dedication'                       ; Patterns = @('^Dedication\s*$') }
    'AbstractKeywords'  = @{ Label = 'Abstract and Keywords'            ; Patterns = @('^Abstract\s*$') }
    'TOC'               = @{ Label = 'Table of Contents'                ; Patterns = @('^Table of Contents\s*$', '^Contents\s*$') }
    'ListOfFigures'     = @{ Label = 'List of Figures'                  ; Patterns = @('^List of Figures\s*$') }
    'ListOfTables'      = @{ Label = 'List of Tables'                   ; Patterns = @('^List of Tables\s*$') }
    'ListOfAcronymsFM'  = @{ Label = 'List of Acronyms and Abbreviations'; Patterns = @('^List of (Acronyms|Symbols|Abbreviations)') }
    'Preface'           = @{ Label = 'Preface'                          ; Patterns = @('^Preface\s*$') }
    'Acknowledgments'   = @{ Label = 'Acknowledgments'                  ; Patterns = @('^Acknowledg(e?)ments\s*$') }
    'ChapterAuthors'    = @{ Label = 'Chapter Authors'                  ; Patterns = @('^Chapter Authors\s*$', '^Author Chapter Credits\s*$', '^Author Contributions\s*$') }
    'ExecutiveSummary'  = @{ Label = 'Executive Summary'                ; Patterns = @('^Executive Summary\s*$') }
}

$projectSequence = @(
    'CoverPage','TitlePage','Disclaimers','Dedication','AbstractKeywords',
    'TOC','ListOfFigures','ListOfTables','Preface','Acknowledgments','ChapterAuthors'
)
$summarySequence = @(
    'CoverPage','TitlePage','Disclaimers','NCSTRoster','Contributors','Dedication',
    'AbstractKeywords','TOC','ListOfFigures','ListOfTables','ListOfAcronymsFM',
    'Preface','Acknowledgments','ExecutiveSummary'
)

$expected = if ($ReportType -eq 'Project') { $projectSequence } else { $summarySequence }

# ---------------------------------------------------------------------------
# Extract paragraph sequence from document.xml
# ---------------------------------------------------------------------------

function Get-DocXml {
    param([string]$Path)
    $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, 'Read')
        try {
            $entry = $zip.GetEntry('word/document.xml')
            if (-not $entry) { throw "No word/document.xml in $Path" }
            $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
}

$docXml = Get-DocXml -Path $targetFile

# Walk paragraphs, extract (index, style, text) for any styled paragraph.
$paraMatches = [regex]::Matches($docXml, '(?s)<w:p\b[^>]*>(?<inner>.*?)</w:p>')
$paragraphs = New-Object System.Collections.ArrayList
$i = 0
foreach ($m in $paraMatches) {
    $inner = $m.Groups['inner'].Value
    $styleM = [regex]::Match($inner, '<w:pStyle w:val="([^"]+)"')
    if (-not $styleM.Success) { $i++; continue }
    $style = $styleM.Groups[1].Value
    $textRuns = [regex]::Matches($inner, '<w:t[^>]*>([^<]*)</w:t>')
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $textRuns) { [void]$sb.Append($t.Groups[1].Value) }
    $text = $sb.ToString().Trim()
    [void]$paragraphs.Add([pscustomobject]@{ Index = $i; Style = $style; Text = $text })
    $i++
}

# ---------------------------------------------------------------------------
# Locate each landmark
# ---------------------------------------------------------------------------

# Real style ids observed in the NIST/NWIRP templates.
# Cover page = the all-caps banner block at the very top of Chapter-00.
#   pStyle:    Title (used for the cover banner text), HMCoverDate (date area)
#   tableSty:  HM-Cover (the cover layout table)
# Title page = the formal title page block that follows the cover.
#   pStyle:    Series, ReportTitle, ReportSubtitle, DraftStage, Authors
# Disclaimer paragraphs use Disclaimer or frontmattertext.
$coverPageStyles  = '^(Title|HMCoverDate)$'
$titlePageStyles  = '^(Series|ReportTitle|ReportSubtitle|DraftStage|Authors)$'
$disclaimerStyles = '^(Disclaimer|frontmattertext)$'

function Find-FirstParagraphMatching {
    param([scriptblock]$Predicate)
    foreach ($p in $paragraphs) {
        if (& $Predicate $p) { return $p }
    }
    return $null
}

$observed = @{}

foreach ($key in $landmarkDefs.Keys) {
    $def = $landmarkDefs[$key]
    $hit = $null
    switch ($key) {
        'CoverPage' {
            # Match the FIRST occurrence of a cover paragraph style. The cover
            # banner sits at the very top of the doc, before any title-page style.
            $hit = Find-FirstParagraphMatching { param($p) $p.Style -match $coverPageStyles }
            if (-not $hit) {
                # Fall back to detecting the HM-Cover table style being referenced
                # anywhere in document.xml (cover may be table-only with no titled paragraphs).
                if ($docXml -match '<w:tblStyle\s+w:val="HM-Cover"') {
                    $hit = [pscustomobject]@{ Index = 0; Style = 'HM-Cover (table)'; Text = '' }
                }
            }
        }
        'TitlePage' {
            $hit = Find-FirstParagraphMatching { param($p) $p.Style -match $titlePageStyles }
        }
        'Disclaimers' {
            $hit = Find-FirstParagraphMatching {
                param($p)
                ($p.Style -match $disclaimerStyles) -and ($p.Text -imatch '(commercial entities|International System of Units|^Disclaimer)')
            }
        }
        default {
            foreach ($pat in $def.Patterns) {
                $hit = Find-FirstParagraphMatching {
                    param($p)
                    $p.Text -imatch $pat
                }
                if ($hit) { break }
            }
        }
    }
    if ($hit) { $observed[$key] = $hit.Index }
}

# ---------------------------------------------------------------------------
# Build the audit
# ---------------------------------------------------------------------------

# Status per expected landmark
$rows = New-Object System.Collections.ArrayList
$prevDocIndex = -1
$outOfOrderKeys = @{}

foreach ($key in $expected) {
    $def = $landmarkDefs[$key]
    $row = [ordered]@{
        Key      = $key
        Label    = $def.Label
        Expected = $expected.IndexOf($key) + 1
        Status   = ''
        DocIndex = $null
        Note     = ''
    }
    if ($observed.ContainsKey($key)) {
        $row.DocIndex = $observed[$key]
        if ($row.DocIndex -lt $prevDocIndex) {
            $row.Status = 'OUT-OF-ORDER'
            $outOfOrderKeys[$key] = $true
        } else {
            $row.Status = 'PRESENT'
        }
        $prevDocIndex = [Math]::Max($prevDocIndex, $row.DocIndex)
    } else {
        $row.Status = 'MISSING'
    }
    [void]$rows.Add([pscustomobject]$row)
}

# Find extras: observed landmarks that are not in the expected sequence for this report type.
$extras = New-Object System.Collections.ArrayList
foreach ($key in $observed.Keys) {
    if ($expected -notcontains $key) {
        [void]$extras.Add([pscustomobject]@{
            Key      = $key
            Label    = $landmarkDefs[$key].Label
            DocIndex = $observed[$key]
            Note     = "Not part of canonical $ReportType Report front matter."
        })
    }
}

$missingCount    = ($rows | Where-Object { $_.Status -eq 'MISSING' }).Count
$outOfOrderCount = ($rows | Where-Object { $_.Status -eq 'OUT-OF-ORDER' }).Count
$extraCount      = $extras.Count
$cleanRun        = ($missingCount + $outOfOrderCount + $extraCount) -eq 0

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

if ($Json) {
    $result = [ordered]@{
        File           = $targetFile
        ReportType     = $ReportType
        Clean          = $cleanRun
        MissingCount   = $missingCount
        OutOfOrderCount = $outOfOrderCount
        ExtraCount     = $extraCount
        Rows           = $rows
        Extras         = $extras
    }
    $result | ConvertTo-Json -Depth 5
} else {
    Write-Host ""
    Write-Host "Front matter audit" -ForegroundColor Cyan
    Write-Host "  File       : $targetFile"
    Write-Host "  ReportType : $ReportType"
    Write-Host ""
    $fmt = "  {0,3}. {1,-38} {2,-14} {3}"
    Write-Host ($fmt -f '#', 'Expected landmark', 'Status', 'Note')
    Write-Host ($fmt -f '---', '--------------------------------------', '--------------', '----')
    foreach ($r in $rows) {
        $color = switch ($r.Status) {
            'PRESENT'       { 'Green' }
            'MISSING'       { 'Red' }
            'OUT-OF-ORDER'  { 'Yellow' }
            default         { 'White' }
        }
        Write-Host ($fmt -f $r.Expected, $r.Label, $r.Status, $r.Note) -ForegroundColor $color
    }
    if ($extras.Count -gt 0) {
        Write-Host ""
        Write-Host "Unexpected sections found:" -ForegroundColor Yellow
        foreach ($x in $extras) {
            Write-Host ("  - {0,-38} (doc paragraph index {1})" -f $x.Label, $x.DocIndex) -ForegroundColor Yellow
            Write-Host ("    {0}" -f $x.Note) -ForegroundColor DarkGray
        }
    }
    Write-Host ""
    if ($cleanRun) {
        Write-Host "Clean. Front matter matches canonical $ReportType Report sequence." -ForegroundColor Green
    } else {
        Write-Host "Drift detected:" -ForegroundColor Red
        Write-Host ("  {0} missing, {1} out-of-order, {2} extra" -f $missingCount, $outOfOrderCount, $extraCount) -ForegroundColor Red
    }
    Write-Host ""
}

if ($cleanRun) { exit 0 } else { exit 1 }
