<#
.SYNOPSIS
    Converts HTML, LaTeX, and Markdown source into Word documents that carry
    the NIST template: pandoc against NIST-reference.docx, a remap pass that
    folds pandoc's styles into NIST styles, then the full NIST build pipeline.

.DESCRIPTION
    Five stages, in this order:

      0. prep    - for HTML sources only, rewrites CSS-only inline formatting
                   into real tags before pandoc sees it. Generators like R
                   flextable emit no <strong>, <em>, or <sup> at all: every run
                   is a <span class="cl-xxxx"> whose formatting lives only in
                   the generated stylesheet, and pandoc reads tags, not CSS, so
                   bold and superscripts are otherwise flattened without a
                   warning. The source file is not touched - the repair is done
                   on a working copy that is deleted afterwards.

      1. pandoc  - converts each source file to .docx using -ReferenceDoc, so
                   the output starts with the NIST theme, page setup, headers
                   and footers, numbering, and styles.

      2. remap   - rewrites the pandoc-specific style references in the output
                   to their NIST equivalents (Image Caption and Table Caption
                   to Caption, Block Text to Quote, First Paragraph and Compact
                   to Body Text, and so on), assigns the NIST-Data table style
                   to every table, and deletes the now-unreferenced pandoc
                   style definitions.

      3. build   - runs Build-NISTDocuments.ps1 on the output folder, which
                   applies theme, styles, page layout, table look, 3-column
                   widths, typography, and font de-embedding.

      4. tables  - finishes the tables, which the generic pipeline cannot do
                   because its settings are tuned for hand-authored NCSTAR
                   documents rather than converted statistical output:

                     - Table Style Options are set: Header Row on, Banded
                       Rows on, everything else off. First Column stays off -
                       turning it on is what bolds and shades column 1. Each
                       option is a parameter. Every table is also pinned to the
                       width its own grid adds up to, with a fixed layout,
                       because a converted table arrives on autofit and two
                       tables with an identical grid otherwise render at
                       different widths. The statistical tables - the ones
                       whose header is Characteristic / N / % - are set to
                       2.7 / 0.6 / 0.6, at all three levels Word reads:
                       tblGrid, tblW, and every cell's tcW. The width is what
                       it is so the full-width green banner rows do not wrap.
                     - cell paragraphs are moved onto Table Text with
                       Set-TableParagraphStyle.ps1, which rewrites w:pStyle and
                       nothing else, so bold, superscripts, subscripts and
                       w:ind all survive. The one exception is -UnboldColumns,
                       which strips direct bold from column 1 of the body rows:
                       converted group labels arrive bold, and because it is
                       direct run formatting no Table Style Option can turn it
                       off. The header row keeps its bold.
                     - the row-hierarchy indentation pandoc drops is rebuilt
                       from the source's leading whitespace with
                       Restore-TableIndents.ps1.
                     - the group-label banner rows inside each table - the
                       full-width merged rows that title a group rather than
                       carry data - are shaded and bolded by
                       Set-TableBannerRows.ps1. It runs after the footnote
                       merge, which is what separates the two kinds of
                       full-width row from each other.
                     - the note rows at the foot of each table are collapsed
                       into one unshaded full-width row with
                       Merge-TableFootnotes.ps1. -SkipTableFootnotes turns it
                       off.
                     - the header sitting above each table becomes a
                       Caption_Table paragraph carrying the NIST number
                       (Table 5-1.  Title), with the table number as a SEQ
                       field, via Set-NISTTableCaptions.ps1. Any number the
                       source already put in the title is dropped, so the
                       caption does not read "Table 5-1.  1 Title".
                       -SkipTableCaptions turns it off.

                   Both of those last two decide by shape what a thing means, so
                   check a new document set once: a table ending in a banner row
                   looks like a footnote, and a section heading sitting above a
                   table looks like a caption. Run each script alone with
                   -ListOnly to see what it would take.

    The order matters. Stage 3's style step deletes styles it does not
    recognize but never touches document.xml, so a paragraph still pointing at
    a deleted pandoc style would silently fall back to Normal. Stage 2 runs
    first so that by the time the pipeline runs, document.xml references only
    NIST styles.

    Source Code and Verbatim Char are deliberately NOT remapped. Code has no
    NIST equivalent, so those two styles are meant to survive into the finished
    document.

    Requires pandoc on PATH (or -PandocPath). Everything else operates on the
    ZIP package directly, so Word is not needed.

.PARAMETER File
    Path to a single source file (.html, .htm, .xhtml, .tex, .md, .markdown,
    .docx). A .docx source is read back against the reference doc, which
    retargets a pandoc- or Quarto-produced Word file at the NIST template while
    keeping its tables, merged cells, and images.

.PARAMETER Folder
    Folder of source files. Every supported extension in it is converted, except
    .docx: a folder scan would convert its own output on a re-run and would
    overwrite the sources in place. Pass Word sources with -File.

.PARAMETER ReferenceDoc
    The pandoc reference document. Defaults to NIST-reference.docx beside this
    script. Build it with Build-NISTReferenceDoc.ps1.

.PARAMETER OutputFolder
    Where the .docx files are written. Defaults to the source folder. Stage 3
    processes EVERY Word file in this folder, so point it at a dedicated
    folder when the source folder already holds unrelated documents.

.PARAMETER TableStyle
    Table style to assign to converted tables. Defaults to NIST-Data. Pass an
    empty string to leave pandoc's table styling alone.

.PARAMETER HeaderRow
    Header Row table style option on the converted tables. Default $false.

.PARAMETER TotalRow
    Total Row table style option. Default $false.

.PARAMETER FirstColumn
    First Column table style option. Default $false.

.PARAMETER LastColumn
    Last Column table style option. Default $false.

.PARAMETER BandedRows
    Banded Rows table style option. Default $false.

.PARAMETER BandedColumns
    Banded Columns table style option. Default $true.

.PARAMETER RepeatHeaderRow
    Repeat the first row across page breaks. Default $true. This is a page-break
    behavior, separate from the Header Row checkbox, so it stays on even with
    -HeaderRow $false.

.PARAMETER TableTextStyle
    StyleId put on every paragraph inside a table. Default TableText. Pass an
    empty string to leave cell paragraph styles alone.

.PARAMETER UnboldColumns
    Column numbers to strip direct bold from in stage 4, body rows only.
    Defaults to column 1, where converted statistical tables put their group
    labels. Pass an empty array to leave all bold alone.

.PARAMETER AutoFitTables
    Leave the tables on Word's autofit instead of pinning them to a width.
    By default each table is pinned to the width its own grid adds up to, with
    a fixed layout, so two tables with the same grid render the same width.

.PARAMETER TableWidthInches
    Pin every table to this width instead of to its own grid, rescaling the
    columns in proportion. The NIST text column is 6.5 inches.

.PARAMETER ColumnWidthsInches
    Exact per-column widths for the statistical tables, defaulting to the
    1.5 / 0.6 / 0.6 the NCSTAR and NWIRP chapters use. Applied only to tables
    whose column count matches and whose header row matches -MatchHeader, so
    every other table keeps its own grid. Pass @() to leave them all alone.

.PARAMETER MatchHeader
    The header row that identifies a statistical table, cell for cell.
    Defaults to Characteristic, N, %.

.PARAMETER SkipTableBanners
    Do not shade and bold the group-label banner rows inside each table.

.PARAMETER BannerFill
    Cell fill for a banner row, six hex digits. Default CCECE3.

.PARAMETER SkipTableFootnotes
    Do not collapse each table's footnote rows into a single full-width row.

.PARAMETER FootnoteStyle
    Paragraph style for the merged table notes. Default TableandFigureFootnotes.

.PARAMETER SkipTableCaptions
    Do not turn the header above each table into a numbered Caption_Table
    paragraph.

.PARAMETER KeepSourceNumbers
    Keep any number the source already put at the front of a table title.
    By default it is dropped, and only when it matches the number the caption
    is being given.

.PARAMETER CaptionFromStyle
    StyleIds the table headers are on. Defaults to Heading1. Takes a list:
    converted HTML often mixes levels, and an HTML <caption> element comes
    through pandoc on Caption. -CaptionFromStyle Heading1,Heading2,Caption

.PARAMETER CaptionStyle
    Caption style to apply. Default CaptionTable (Caption_Table in the template).

.PARAMETER CaptionChapter
    Write the chapter number as this literal value instead of as a field.
    Passing it at all is what selects a literal, so a multi-chapter batch
    converted in one run should leave it alone. Digits, or a single letter for
    an appendix.

.PARAMETER CaptionUseStyleRef
    Write the chapter number as a STYLEREF field. This is the default, and the
    parameter is only needed to force a field alongside an explicit
    -CaptionChapter. The field reads the Heading 1 the caption sits under, so
    the number is right wherever the content is later merged; before the merge,
    a standalone extract with no Heading 1 shows a field error until it lands
    in a chapter.

.PARAMETER SkipInlineRepair
    Skip stage 0. Use for HTML that already carries real <strong>/<em>/<sup>.

.PARAMETER SkipTableIndents
    Skip the indent rebuild in stage 4.

.PARAMETER RemapLists
    Also convert pandoc's direct list numbering to the NIST List Bullet and
    List Number ladders. Off by default: it gives bullets the NIST ladder
    cleanly, but ordered lists lose pandoc's explicit start values and
    restarts, so separate numbered lists may run on as one sequence.

.PARAMETER Bibliography
    Bibliography file (.bib, .json, .yaml). Turns on --citeproc.

.PARAMETER Csl
    CSL style file for citeproc.

.PARAMETER NoMedia
    Skip --extract-media. Images are embedded in the .docx either way; the
    extract folder just keeps the originals on disk.

.PARAMETER PandocArgs
    Extra arguments passed through to pandoc verbatim.

.PARAMETER PandocPath
    Path to the pandoc executable. Defaults to pandoc on PATH.

.PARAMETER SkipRemap
    Stop after stage 1 and 3, leaving pandoc's own styles in place.

.PARAMETER SkipPipeline
    Stop after stage 2. Use to inspect the converted files before the NIST
    build runs. Stage 4 is part of the build, so it is skipped too.

.PARAMETER LegacyStyleFix
    Run the pipeline's style step in full legacy-fix mode instead of -Clean.
    Rarely needed: converted output is clean by construction.

.PARAMETER Force
    Skip prompts and overwrite existing output files.

.EXAMPLE
    .\Convert-ToNISTWord.ps1 -File "chapter-03.html"

.EXAMPLE
    .\Convert-ToNISTWord.ps1 -Folder "C:\src\report" -OutputFolder "C:\out\report" -Force

.EXAMPLE
    .\Convert-ToNISTWord.ps1 -File "paper.tex" -Bibliography "refs.bib" -SkipPipeline

.EXAMPLE
    # Retarget an existing Quarto/pandoc Word file at the NIST template.
    .\Convert-ToNISTWord.ps1 -File "C:\src\analysis.docx" -OutputFolder "C:\out"
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$ReferenceDoc,

    [string]$OutputFolder,

    [string]$TableStyle = 'NIST-Data',

    # Table Style Options for converted tables: the standard report look, same
    # as Set-TableLook.ps1's own defaults. First Column stays off - turning it
    # on is what bolds and shades the first column.
    [bool]$HeaderRow       = $true,
    [bool]$TotalRow        = $false,
    [bool]$FirstColumn     = $false,
    [bool]$LastColumn      = $false,
    [bool]$BandedRows      = $true,
    [bool]$BandedColumns   = $false,
    [bool]$RepeatHeaderRow = $true,

    [string]$TableTextStyle = 'TableText',

    # Off by default. The bold full-width rows in a converted statistical table
    # are group-label banners, and they are meant to be bold - see
    # Set-TableBannerRows.ps1, which shades them to match.
    [int[]]$UnboldColumns = @(),

    # Converted tables arrive as tblW auto, so Word sizes them from their
    # contents and identical grids still render at different widths.
    [switch]$AutoFitTables,

    [double]$TableWidthInches = 0,

    # The NCSTAR/NWIRP statistical table shape. Gated on the header, so a table
    # that is not one of these is never touched.
    #
    # The two small columns are equal and stay at 0.6, which clears the widest
    # value in either of them (%, at 0.402 including cell margins). Column 1
    # takes the rest, and the rest is set by the GREEN BANNER rows rather than
    # by column 1's own labels: a banner spans all three columns, and the widest
    # one measures 3.685 in of bold Source Sans Pro 9pt, so the table has to
    # reach 3.745 in for it to stay on one line. 2.7/0.6/0.6 gives 3.9 in, which
    # clears it with room to spare. Column 1's own labels only need 1.513 in.
    [double[]]$ColumnWidthsInches = @(2.7, 0.6, 0.6),

    [string[]]$MatchHeader = @('Characteristic', 'N', '%'),

    [switch]$SkipTableBanners,

    [string]$BannerFill = 'CCECE3',

    [switch]$SkipTableFootnotes,

    [string]$FootnoteStyle = 'TableandFigureFootnotes',

    [switch]$SkipTableCaptions,

    [string[]]$CaptionFromStyle = @('Heading1'),

    [string]$CaptionStyle = 'CaptionTable',

    [string]$CaptionChapter = '1',

    [switch]$CaptionUseStyleRef,

    [switch]$KeepSourceNumbers,

    [switch]$SkipInlineRepair,

    [switch]$SkipTableIndents,

    [switch]$RemapLists,

    [string]$Bibliography,

    [string]$Csl,

    [switch]$NoMedia,

    [string[]]$PandocArgs,

    [string]$PandocPath = 'pandoc',

    [switch]$SkipRemap,

    [switch]$SkipPipeline,

    [switch]$LegacyStyleFix,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ooXmlModule = Join-Path $PSScriptRoot 'OoXml\OoXml.psd1'
Import-Module $ooXmlModule -Force

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# ---------------------------------------------------------------------------
# Argument sanity, before anything slow runs
#
# -CaptionChapter takes a value, so "-CaptionChapter 1-Force" binds the whole
# string "1-Force" as the chapter and silently drops -Force. Left unchecked
# that reaches stage 4 and writes "Table 1-Force-1." into every caption in the
# file, after pandoc has already done all the work.
# ---------------------------------------------------------------------------

if ($PSBoundParameters.ContainsKey('CaptionChapter') -and
    $CaptionChapter -notmatch '^([0-9]+|[A-Za-z])$') {
    $msg = "-CaptionChapter must be digits, or a single letter for an appendix. Got: '$CaptionChapter'."
    if ($CaptionChapter -match '^([0-9A-Za-z]+)-([A-Za-z]+)$') {
        $msg += "`n  That looks like a missing space: -CaptionChapter $($Matches[1]) -$($Matches[2])"
    }
    throw $msg
}

# ---------------------------------------------------------------------------
# Source format per extension
# ---------------------------------------------------------------------------
# raw_tex on the LaTeX reader keeps unknown macros as raw blocks instead of
# dropping them silently, so what pandoc could not interpret is visible in the
# output rather than missing from it.
#
# .docx is an input format too. Quarto, R Markdown, and pandoc itself often
# produce a Word file that is correct but carries pandoc's own styles instead of
# the template's. Reading it back against the reference doc retargets the whole
# document at the NIST template while keeping tables, merged cells, and images
# intact. -Folder never picks these up (see $folderExtensions below); pass -File.
$formats = @{
    '.html'     = 'html'
    '.htm'      = 'html'
    '.xhtml'    = 'html'
    '.tex'      = 'latex+raw_tex'
    '.md'       = 'markdown'
    '.markdown' = 'markdown'
    '.docx'     = 'docx'
}

# Extensions -Folder will collect. .docx is excluded on purpose: a folder scan
# that picked up Word files would try to convert its own output on a re-run, and
# would overwrite every source in place when -OutputFolder is left unset.
$folderExtensions = @('.html', '.htm', '.xhtml', '.tex', '.md', '.markdown')

# ---------------------------------------------------------------------------
# Stage 2 mapping: pandoc styleId -> ordered NIST fallbacks
# ---------------------------------------------------------------------------
# These are the ids Build-NISTReferenceDoc.ps1 creates, which are also the ids
# pandoc generates on its own when a name is missing from the reference doc, so
# the map holds either way. Styles absent from this table are left alone.
$styleRemap = [ordered]@{
    'FirstParagraph'  = @('BodyText')
    'Compact'         = @('BodyText')
    'BlockText'       = @('Quote', 'BodyText')
    'ImageCaption'    = @('Caption')
    'TableCaption'    = @('Caption')
    'Figure'          = @('graphic', 'BodyText')
    'CaptionedFigure' = @('graphic', 'BodyText')
    'Author'          = @('BodyText')
    'Date'            = @('BodyText')
    'Abstract'        = @('BodyText')
    'AbstractTitle'   = @('Frontmatterhead1', 'Heading1')
}

# Paragraph styles whose list paragraphs -RemapLists is allowed to restyle.
$listSourceStyles = @('Compact', 'FirstParagraph', 'BodyText', 'ListParagraph')

# ---------------------------------------------------------------------------
# Resolve inputs
# ---------------------------------------------------------------------------

if (-not $ReferenceDoc) { $ReferenceDoc = Join-Path $PSScriptRoot 'NIST-reference.docx' }
if (-not (Test-Path -LiteralPath $ReferenceDoc -PathType Leaf)) {
    throw "Reference doc not found: $ReferenceDoc`nBuild it first: .\Build-NISTReferenceDoc.ps1 -Template <NIST template.docx>"
}
$ReferenceDoc = (Resolve-Path -LiteralPath $ReferenceDoc).ProviderPath

$pandoc = Get-Command $PandocPath -ErrorAction SilentlyContinue
if (-not $pandoc) { throw "pandoc not found: $PandocPath. Install pandoc or pass -PandocPath." }
$pandocExe = $pandoc.Source

$sources = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $sources += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $sources = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object { $folderExtensions -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object Name |
        ForEach-Object { $_.FullName }
}
if ($sources.Count -eq 0) { throw "No supported source files found ($($folderExtensions -join ', ')). Word sources need -File." }

foreach ($s in $sources) {
    $ext = [System.IO.Path]::GetExtension($s).ToLowerInvariant()
    if (-not $formats.ContainsKey($ext)) { throw "Unsupported source type: $s" }
}

# Two sources with the same stem (report.html and report.tex) would both want
# report.docx and the second would silently overwrite the first. Catch it here,
# before any work is done, rather than after half the folder is converted.
$byOutput = @{}
foreach ($s in $sources) {
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($s) + '.docx'
    if (-not $byOutput.ContainsKey($stem)) { $byOutput[$stem] = @() }
    $byOutput[$stem] += $s
}
$clashes = @($byOutput.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 })
if ($clashes.Count -gt 0) {
    $detail = ($clashes | ForEach-Object { "  $($_.Key) <- " + (($_.Value | ForEach-Object { Split-Path -Leaf $_ }) -join ', ') }) -join "`n"
    throw "Source files collide on the same output name:`n$detail`nRename them or convert them one at a time with -File."
}

if (-not $OutputFolder) {
    $OutputFolder = Split-Path -Parent $sources[0]
} elseif (-not (Test-Path -LiteralPath $OutputFolder -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath

# A .docx source lands on its own name in the output folder. If that resolves
# back to the source itself, pandoc would read and write the same file and the
# original would be gone. Nothing else can hit this, since every other source
# type changes extension.
foreach ($s in $sources) {
    if ([System.IO.Path]::GetExtension($s).ToLowerInvariant() -ne '.docx') { continue }
    $dest = Join-Path $OutputFolder ([System.IO.Path]::GetFileName($s))
    if ([string]::Equals($dest, $s, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Word source and its output are the same file: $s`nPass -OutputFolder pointing somewhere else, so the original survives."
    }
}

if ($Bibliography) {
    if (-not (Test-Path -LiteralPath $Bibliography -PathType Leaf)) { throw "Bibliography not found: $Bibliography" }
    $Bibliography = (Resolve-Path -LiteralPath $Bibliography).ProviderPath
}
if ($Csl) {
    if (-not (Test-Path -LiteralPath $Csl -PathType Leaf)) { throw "CSL file not found: $Csl" }
    $Csl = (Resolve-Path -LiteralPath $Csl).ProviderPath
}

# ---------------------------------------------------------------------------
# Confirm
# ---------------------------------------------------------------------------

$htmlSources = @($sources | Where-Object { [System.IO.Path]::GetExtension($_).ToLowerInvariant() -match '^\.(html|htm|xhtml)$' })

$stages = @()
if (-not $SkipInlineRepair -and $htmlSources.Count -gt 0) { $stages += '0. inline formatting' }
$stages += '1. pandoc convert'
if (-not $SkipRemap)    { $stages += '2. remap pandoc styles to NIST' }
if (-not $SkipPipeline) { $stages += '3. Build-NISTDocuments' + $(if ($LegacyStyleFix) { ' (full style fix)' } else { ' (-Clean)' }) }
if (-not $SkipPipeline) { $stages += '4. tables' }

Write-Host "Convert to NIST Word" -ForegroundColor Cyan
Write-Host "  Reference doc: $ReferenceDoc" -ForegroundColor Cyan
Write-Host "  Output folder: $OutputFolder" -ForegroundColor Cyan
Write-Host "  pandoc:        $pandocExe" -ForegroundColor Cyan
Write-Host "  Stages:        $($stages -join '  ->  ')" -ForegroundColor Cyan
Write-Host "  Sources:" -ForegroundColor Cyan
$sources | ForEach-Object { Write-Host "    $_" }
if (-not $SkipPipeline) {
    Write-Host ""
    Write-Host "  Note: stage 3 processes EVERY Word file in $OutputFolder, not just the converted ones." -ForegroundColor Yellow
}
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Package read / write helpers
# ---------------------------------------------------------------------------

function Read-Package {
    param([string]$Path)
    $dict = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $dict[$e.FullName] = $ms.ToArray()
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
    # Comma so PowerShell hands back the collection instead of enumerating it.
    return ,$dict
}

function Write-Package {
    param($Entries, [string]$Path)
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "nistconv-$([guid]::NewGuid()).zip")
    $outStream = [System.IO.File]::Create($tempZip)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $Entries.Keys) {
                $ne = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try { $os.Write($Entries[$name], 0, $Entries[$name].Length) } finally { $os.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $outStream.Dispose() }
    Move-Item -LiteralPath $tempZip -Destination $Path -Force
}

# styleIds present in this styles.xml, as a case-insensitive lookup that hands
# back the id in its ORIGINAL casing. Templates differ on case (Graphic vs
# graphic) and a reference written in the wrong case is a reference to nothing,
# so every id this script emits comes from here rather than from a literal.
function Get-StyleIdMap {
    param([string]$Xml)
    $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in [regex]::Matches($Xml, 'w:styleId="([^"]+)"')) {
        $id = $m.Groups[1].Value
        if (-not $map.ContainsKey($id)) { $map[$id] = $id }
    }
    return ,$map
}

# ---------------------------------------------------------------------------
# Stage 1 and 2, per file
# ---------------------------------------------------------------------------

$converted = @()
$sourceOf  = @{}

foreach ($src in $sources) {
    $ext     = [System.IO.Path]::GetExtension($src).ToLowerInvariant()
    $from    = $formats[$ext]
    $srcDir  = Split-Path -Parent $src
    $outFile = Join-Path $OutputFolder ([System.IO.Path]::GetFileNameWithoutExtension($src) + '.docx')

    Write-Host ""
    Write-Host "===== $([System.IO.Path]::GetFileName($src)) ($from) =====" -ForegroundColor Green

    if ((Test-Path -LiteralPath $outFile) -and -not $Force) {
        $reply = Read-Host "  $outFile exists. Overwrite? (y/N)"
        if ($reply -notmatch '^[yY]') { Write-Host "  Skipped." -ForegroundColor DarkGray; continue }
    }

    # LaTeX figures in PDF or EPS cannot be embedded in a .docx; Word shows a
    # broken placeholder. Flag them rather than let them fail silently.
    if ($ext -eq '.tex') {
        $texBody = Get-Content -LiteralPath $src -Raw
        $badFigs = [regex]::Matches($texBody, '(?i)\\includegraphics[^\{]*\{([^\}]+\.(?:pdf|eps|ps))\}') |
            ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
        if ($badFigs.Count -gt 0) {
            Write-Warning "  PDF/EPS figures cannot be embedded in Word. Convert these to PNG or SVG first:"
            $badFigs | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
        }
    }

    # ---- Stage 0: HTML inline-formatting repair ----
    # pandoc reads tags, not CSS. Rewrite class-only bold/italic/super/sub into
    # real tags first, on a working copy in the source folder so that relative
    # image paths still resolve. $pandocSrc is what pandoc actually reads.
    $pandocSrc = $src
    $prepFile  = $null
    if (-not $SkipInlineRepair -and $ext -match '^\.(html|htm|xhtml)$') {
        $repair = Join-Path $PSScriptRoot 'Repair-HtmlInlineFormatting.ps1'
        if (-not (Test-Path -LiteralPath $repair -PathType Leaf)) {
            Write-Warning "  Repair-HtmlInlineFormatting.ps1 not found; CSS-only bold and superscripts will be lost."
        } else {
            $prepFile = Join-Path $srcDir ([System.IO.Path]::GetFileNameWithoutExtension($src) + '.nistprep' + $ext)
            Copy-Item -LiteralPath $src -Destination $prepFile -Force
            Write-Host "  ---- Stage 0: inline formatting ----" -ForegroundColor DarkGray
            & $repair -File $prepFile -Force | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
            Remove-Item -LiteralPath "$prepFile.inline.bak" -Force -ErrorAction SilentlyContinue
            $pandocSrc = $prepFile
        }
    }

    # ---- Stage 1: pandoc ----
    $argList = @(
        '--from',          $from,
        '--to',            'docx',
        '--reference-doc', $ReferenceDoc,
        '--resource-path', $srcDir,
        '--output',        $outFile
    )
    if (-not $NoMedia)  { $argList += @('--extract-media', (Join-Path $OutputFolder 'media')) }
    if ($Bibliography)  { $argList += @('--citeproc', '--bibliography', $Bibliography) }
    if ($Csl)           { $argList += @('--csl', $Csl) }
    if ($PandocArgs)    { $argList += $PandocArgs }
    $argList += $pandocSrc

    Write-Host "  pandoc $($argList -join ' ')" -ForegroundColor DarkGray
    # Native stderr redirected into the output stream surfaces as ErrorRecords,
    # which $ErrorActionPreference='Stop' would turn into a terminating error
    # before the exit code can be read. Relax it just around the call.
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $pandocOut = & $pandocExe @argList 2>&1
    } finally {
        $ErrorActionPreference = $prevEAP
    }
    if ($prepFile) { Remove-Item -LiteralPath $prepFile -Force -ErrorAction SilentlyContinue }
    if ($LASTEXITCODE -ne 0) {
        $pandocOut | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
        throw "pandoc failed on $src (exit $LASTEXITCODE)."
    }
    if ($pandocOut) { $pandocOut | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow } }
    Write-Host "  Converted -> $outFile" -ForegroundColor Green
    $converted += $outFile
    $sourceOf[$outFile] = $src

    if ($SkipRemap) { continue }

    # ---- Stage 2: remap ----
    $entries = Read-Package -Path $outFile
    if (-not $entries.ContainsKey('word/styles.xml')) { throw "No word/styles.xml in $outFile" }

    $stylesXml = $utf8NoBom.GetString($entries['word/styles.xml'])
    $styleIds  = Get-StyleIdMap -Xml $stylesXml

    # Resolve each pandoc style to the first NIST fallback that exists here.
    # Both sides are canonicalized to the casing styles.xml actually uses.
    $resolved = [ordered]@{}
    foreach ($pandocId in $styleRemap.Keys) {
        if (-not $styleIds.ContainsKey($pandocId)) { continue }
        $actualId = $styleIds[$pandocId]
        foreach ($candidate in $styleRemap[$pandocId]) {
            if ($styleIds.ContainsKey($candidate)) { $resolved[$actualId] = $styleIds[$candidate]; break }
        }
        if (-not $resolved.Contains($actualId)) {
            Write-Warning "  No NIST target present for $actualId; leaving it as-is."
        }
    }

    # Assigning a table style that styles.xml does not define leaves Word with a
    # dangling reference. Set-NISTStyleArchitecture installs NIST-Data as a
    # canonical block, so stage 3 supplies it; without stage 3 it has to already
    # be in the reference doc.
    $tableStyleForFile = $TableStyle
    if ($TableStyle -and $styleIds.ContainsKey($TableStyle)) {
        $tableStyleForFile = $styleIds[$TableStyle]
    } elseif ($TableStyle) {
        if ($SkipPipeline) {
            Write-Warning "  Table style $TableStyle is not defined in this document and the NIST build is skipped; leaving tables unstyled."
            $tableStyleForFile = ''
        } else {
            Write-Host "  Table style $TableStyle not in the reference doc; the NIST build step installs it." -ForegroundColor DarkGray
        }
    }

    # Numbering formats, for -RemapLists: numId -> numFmt at level 0.
    $numFmtByNumId = @{}
    if ($RemapLists -and $entries.ContainsKey('word/numbering.xml')) {
        $numXml = $utf8NoBom.GetString($entries['word/numbering.xml'])
        $absFmt = @{}
        foreach ($m in [regex]::Matches($numXml, '(?s)<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"[^>]*>(.*?)</w:abstractNum>')) {
            $lvl0 = [regex]::Match($m.Groups[2].Value, '(?s)<w:lvl\b[^>]*w:ilvl="0"[^>]*>(.*?)</w:lvl>')
            if ($lvl0.Success -and $lvl0.Groups[1].Value -match '<w:numFmt w:val="([^"]+)"') {
                $absFmt[$m.Groups[1].Value] = $Matches[1]
            }
        }
        foreach ($m in [regex]::Matches($numXml, '(?s)<w:num\b[^>]*w:numId="(\d+)"[^>]*>(.*?)</w:num>')) {
            if ($m.Groups[2].Value -match '<w:abstractNumId w:val="(\d+)"') {
                $aid = $Matches[1]
                if ($absFmt.ContainsKey($aid)) { $numFmtByNumId[$m.Groups[1].Value] = $absFmt[$aid] }
            }
        }
    }

    $textParts = @($entries.Keys | Where-Object {
        $_ -match '^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$'
    })

    # The table and list rewrites run inside MatchEvaluator scriptblocks, which
    # execute in a child scope, so their counters live at script scope.
    $remapCount = 0
    $script:tableHits = 0
    $script:listHits  = 0

    foreach ($part in $textParts) {
        $xml = $utf8NoBom.GetString($entries[$part])

        # Style references: paragraph styles and character styles alike.
        foreach ($pandocId in $resolved.Keys) {
            $target  = $resolved[$pandocId]
            $pattern = '<w:(pStyle|rStyle) w:val="' + [regex]::Escape($pandocId) + '"\s*/>'
            $hits    = [regex]::Matches($xml, $pattern).Count
            if ($hits -gt 0) {
                $xml = [regex]::Replace($xml, $pattern, ('<w:$1 w:val="' + $target + '"/>'))
                $remapCount += $hits
            }
        }

        # Lists: give NIST's ladder control by moving to the List styles and
        # dropping pandoc's direct numbering.
        if ($RemapLists -and $numFmtByNumId.Count -gt 0) {
            $xml = [regex]::Replace($xml, '(?s)<w:p\b[^>]*>.*?</w:p>', {
                param($m)
                $p = $m.Value
                if ($p -notmatch '<w:numPr>') { return $p }
                $numId = ''
                if ($p -match '<w:numId w:val="(\d+)"\s*/>') { $numId = $Matches[1] }
                if (-not $numId -or -not $numFmtByNumId.ContainsKey($numId)) { return $p }
                $ilvl = 0
                if ($p -match '<w:ilvl w:val="(\d+)"\s*/>') { $ilvl = [int]$Matches[1] }
                $current = ''
                if ($p -match '<w:pStyle w:val="([^"]+)"\s*/>') { $current = $Matches[1] }
                if ($current -and ($listSourceStyles -notcontains $current)) { return $p }

                $family = 'ListNumber'
                if ($numFmtByNumId[$numId] -eq 'bullet') { $family = 'ListBullet' }
                $level = [Math]::Min($ilvl + 1, 5)
                $newId = $family
                if ($level -gt 1) { $newId = "$family$level" }
                # Only move the paragraph if that rung of the ladder exists, and
                # spell the id the way this styles.xml spells it.
                if (-not $styleIds.ContainsKey($newId)) { return $p }
                $newId = $styleIds[$newId]

                $p = [regex]::Replace($p, '(?s)<w:numPr>.*?</w:numPr>', '', 1)
                if ($current) {
                    $p = [regex]::Replace($p, '<w:pStyle w:val="[^"]+"\s*/>', "<w:pStyle w:val=`"$newId`"/>", 1)
                } elseif ($p -match '<w:pPr>') {
                    $p = [regex]::Replace($p, '<w:pPr>', "<w:pPr><w:pStyle w:val=`"$newId`"/>", 1)
                } else {
                    $p = [regex]::Replace($p, '(<w:p\b[^>]*>)', "`$1<w:pPr><w:pStyle w:val=`"$newId`"/></w:pPr>", 1)
                }
                $script:listHits++
                return $p
            })
        }

        # Tables: pandoc's table style is not in the NIST allowlist, so point
        # every table at the NIST table style instead.
        if ($tableStyleForFile -and $part -eq 'word/document.xml') {
            $xml = [regex]::Replace($xml, '(?s)<w:tblPr>(.*?)</w:tblPr>', {
                param($m)
                $inner = $m.Groups[1].Value
                $script:tableHits++
                if ($inner -match '<w:tblStyle w:val="[^"]*"\s*/>') {
                    $inner = [regex]::Replace($inner, '<w:tblStyle w:val="[^"]*"\s*/>', "<w:tblStyle w:val=`"$tableStyleForFile`"/>", 1)
                } else {
                    # Schema order puts tblStyle first inside tblPr.
                    $inner = "<w:tblStyle w:val=`"$tableStyleForFile`"/>" + $inner
                }
                return "<w:tblPr>$inner</w:tblPr>"
            })
        }

        $entries[$part] = $utf8NoBom.GetBytes($xml)
    }

    $listCount  = $script:listHits
    $tableCount = $script:tableHits

    # Drop the style definitions that nothing references any more.
    $removed = 0
    foreach ($pandocId in $resolved.Keys) {
        $r = Remove-StyleBlock -Xml $stylesXml -Id $pandocId
        $stylesXml = $r[0]
        if ($r[1]) { $removed++ }
    }
    $entries['word/styles.xml'] = $utf8NoBom.GetBytes($stylesXml)

    Write-Package -Entries $entries -Path $outFile

    Write-Host ("  Remapped {0} style reference(s), removed {1} pandoc style definition(s)." -f $remapCount, $removed) -ForegroundColor Gray
    if ($tableStyleForFile) { Write-Host "  Tables set to ${tableStyleForFile}: $tableCount" -ForegroundColor Gray }
    if ($RemapLists) { Write-Host "  List paragraphs moved to the NIST ladder: $listCount" -ForegroundColor Gray }
    if ($resolved.Count -gt 0) {
        foreach ($pandocId in $resolved.Keys) { Write-Host ("    {0,-16} -> {1}" -f $pandocId, $resolved[$pandocId]) -ForegroundColor DarkGray }
    }
}

if ($converted.Count -eq 0) { throw "Nothing was converted." }

# ---------------------------------------------------------------------------
# Stage 3: the NIST build pipeline
# ---------------------------------------------------------------------------

if ($SkipPipeline) {
    Write-Host ""
    Write-Host "Converted $($converted.Count) file(s). Pipeline skipped (-SkipPipeline)." -ForegroundColor Green
    Write-Host "Run it when ready: .\Build-NISTDocuments.ps1 -Folder `"$OutputFolder`" -Clean" -ForegroundColor Cyan
    return
}

$pipeline = Join-Path $PSScriptRoot 'Build-NISTDocuments.ps1'
if (-not (Test-Path -LiteralPath $pipeline -PathType Leaf)) { throw "Pipeline script not found: $pipeline" }

Write-Host ""
Write-Host "===== Stage 3: NIST build pipeline =====" -ForegroundColor Green
$pipeArgs = @{ Folder = $OutputFolder; Force = $true }
if (-not $LegacyStyleFix) { $pipeArgs['Clean'] = $true }
& $pipeline @pipeArgs

# ---------------------------------------------------------------------------
# Stage 4: finish the tables
#
# Every step here is optional polish on a document that is already converted and
# built. One of them failing must not abandon the others, or a file with no
# table notes loses its captions too, so each is caught and reported.
#
# Build-NISTDocuments ran Set-TableLook with ITS defaults, which are the
# hand-authored NCSTAR look (Header Row on, Banded Rows on). Converted
# statistical tables want the other set, so this re-runs it with the
# parameters, then puts the cell text on Table Text and rebuilds the row
# indentation pandoc dropped. Order matters: the style move has to happen
# before the indents, because Set-TableParagraphStyle preserves w:ind but only
# the paragraphs it leaves in place.
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "===== Stage 4: tables =====" -ForegroundColor Green

$tableLook    = Join-Path $PSScriptRoot 'Set-TableLook.ps1'
$tableParaSty = Join-Path $PSScriptRoot 'Set-TableParagraphStyle.ps1'
$tableIndents = Join-Path $PSScriptRoot 'Restore-TableIndents.ps1'
$tableFootnts = Join-Path $PSScriptRoot 'Merge-TableFootnotes.ps1'
$tableCaption = Join-Path $PSScriptRoot 'Set-NISTTableCaptions.ps1'
$tableBanners = Join-Path $PSScriptRoot 'Set-TableBannerRows.ps1'

foreach ($outFile in $converted) {
    Write-Host ""
    Write-Host "  $([System.IO.Path]::GetFileName($outFile))" -ForegroundColor Yellow

    $lookArgs = @{
        File            = $outFile
        Force           = $true
        HeaderRow       = $HeaderRow
        TotalRow        = $TotalRow
        FirstColumn     = $FirstColumn
        LastColumn      = $LastColumn
        BandedRows      = $BandedRows
        BandedColumns   = $BandedColumns
        RepeatHeaderRow = $RepeatHeaderRow
    }
    if (-not $AutoFitTables) {
        $lookArgs['FixedWidth'] = $true
        if ($TableWidthInches -gt 0) { $lookArgs['TableWidthInches'] = $TableWidthInches }
        if ($ColumnWidthsInches.Count -gt 0) {
            $lookArgs['ColumnWidthsInches'] = $ColumnWidthsInches
            if ($MatchHeader.Count -gt 0) { $lookArgs['MatchHeader'] = $MatchHeader }
        }
    }
    & $tableLook @lookArgs

    if ($TableTextStyle) {
        if (Test-Path -LiteralPath $tableParaSty -PathType Leaf) {
            $paraArgs = @{ File = $outFile; Style = $TableTextStyle; Force = $true }
            if ($UnboldColumns.Count -gt 0) { $paraArgs['UnboldColumns'] = $UnboldColumns }
            & $tableParaSty @paraArgs
        } else {
            Write-Warning "  Set-TableParagraphStyle.ps1 not found; cell paragraphs left on the converter's styles."
        }
    }

    if (-not $SkipTableIndents -and (Test-Path -LiteralPath $tableIndents -PathType Leaf)) {
        # A .docx source still has the leading tabs to map from; HTML output
        # carries the hierarchy as leading spaces in the target itself.
        $origin = $sourceOf[$outFile]
        $indentArgs = @{ Target = $outFile; Force = $true }
        if ($origin -and [System.IO.Path]::GetExtension($origin).ToLowerInvariant() -eq '.docx') {
            $indentArgs['Source'] = $origin
        }
        try {
            & $tableIndents @indentArgs
        } catch {
            Write-Warning "  Indent rebuild skipped: $($_.Exception.Message)"
        }
    }

    # Footnote rows last: it rewrites whole rows, so it wants the cell style
    # and the indents already settled. Off unless asked for - the rows are
    # found by shape, and a table ending in a banner row matches the same rule.
    if (-not $SkipTableFootnotes) {
        if (Test-Path -LiteralPath $tableFootnts -PathType Leaf) {
            $fnArgs = @{ File = $outFile; Style = $FootnoteStyle; Force = $true }
            # An HTML original carries the notes in <tfoot>, which identifies
            # them exactly instead of by row shape.
            $origin = $sourceOf[$outFile]
            if ($origin -and [System.IO.Path]::GetExtension($origin).ToLowerInvariant() -match '^\.(html|htm|xhtml)$') {
                $fnArgs['SourceHtml'] = $origin
            }
            try { & $tableFootnts @fnArgs }
            catch { Write-Warning "  Footnote merge skipped: $($_.Exception.Message)" }
        } else {
            Write-Warning "  Merge-TableFootnotes.ps1 not found; table note rows left as separate rows."
        }
    }

    # Banners after the footnote merge, which is what tells the two kinds of
    # full-width merged row apart: the notes are on the footnote style by now,
    # so whatever is still full-width and unclaimed is a group label.
    if (-not $SkipTableBanners) {
        if (Test-Path -LiteralPath $tableBanners -PathType Leaf) {
            try { & $tableBanners -File $outFile -Fill $BannerFill -Force }
            catch { Write-Warning "  Banner shading skipped: $($_.Exception.Message)" }
        } else {
            Write-Warning "  Set-TableBannerRows.ps1 not found; group-label rows left unshaded."
        }
    }

    # Captions last. The caption style only exists once stage 3 has applied the
    # style architecture, and the script refuses to run without it rather than
    # write a reference that renders as Normal.
    if (-not $SkipTableCaptions) {
        if (Test-Path -LiteralPath $tableCaption -PathType Leaf) {
            $capArgs = @{
                File      = $outFile
                FromStyle = $CaptionFromStyle
                Style     = $CaptionStyle
                Chapter   = $CaptionChapter
                Force     = $true
            }
            # The chapter number is a STYLEREF field by default, so each
            # caption reads the chapter it actually lands under instead of
            # carrying a number baked in at conversion time. Passing an
            # explicit -CaptionChapter is what asks for a literal.
            $literalChapter = $PSBoundParameters.ContainsKey('CaptionChapter')
            if ($CaptionUseStyleRef -or -not $literalChapter) { $capArgs['UseStyleRef'] = $true }
            # Converted sources normally number their own tables in the title.
            if (-not $KeepSourceNumbers) { $capArgs['StripLeadingNumber'] = $true }
            try { & $tableCaption @capArgs }
            catch { Write-Warning "  Caption step skipped: $($_.Exception.Message)" }
        } else {
            Write-Warning "  Set-NISTTableCaptions.ps1 not found; table headers left as headings."
        }
    }
}

Write-Host ""
Write-Host "Done. $($converted.Count) file(s) converted and built:" -ForegroundColor Green
$converted | ForEach-Object { Write-Host "  $_" }
Write-Host ""
Write-Host "Still manual: header and footer content (Set-NISTHeaderFooterContent.ps1) and document titles (Set-DocTitle.ps1)." -ForegroundColor Cyan
