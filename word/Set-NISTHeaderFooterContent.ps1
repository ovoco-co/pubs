<#
.SYNOPSIS
    Installs NIST NCSTAR/NWIRP-spec header and footer parts in a Word file.

.DESCRIPTION
    Per comment 4 (NIST review of FedWriters template):

      Header (after first page of each chapter):
        - Even pages: chapter number, left-aligned (STYLEREF "Heading 1" \n)
        - Odd pages:  chapter title,  right-aligned (STYLEREF "Heading 1")
        - First page of each section: blank
      Footer:
        - Even pages: PAGE on left (outer), TITLE on right (inner near binding)
        - Odd pages:  TITLE on left (inner near binding), PAGE on right (outer)
        - First page: TITLE + PAGE (recto layout) on chapters; blank on the
          front matter (pass -FrontMatter). The first-page HEADER stays blank
          in both cases - the chapter title is already on that page.
      Header paragraph applies style HeaderEven (left-aligned) or HeaderOdd
      (right-aligned). Footer paragraph applies style FooterEven or FooterOdd.
      All four styles are created by Set-NISTStyleArchitecture.ps1 and carry
      the horizontal rule (Resolution Blue, 0.75pt) plus footer tab stops.
      Run that script before this one.

    The TITLE field pulls from docProps/core.xml <dc:title>; set that with
    Set-NISTDocTitle.ps1 first. The chapter STYLEREF fields read from
    Heading 1 paragraphs.

    Removes ALL existing header*.xml / footer*.xml parts and replaces them
    with six fresh ones (nistHeaderFirst/Even/Default, nistFooterFirst/
    Even/Default). Wires references into every sectPr, adds <w:titlePg/>,
    enables <w:evenAndOddHeaders/> and <w:mirrorMargins/> in settings.xml.

    Idempotent: re-running replaces the same nist* parts.

    Finally, recalculates the header and footer fields through Word. Everything
    before that step edits the ZIP directly and needs no Word, but STYLEREF and
    PAGE fields cache their result and only Word can compute a new one. Ctrl+A
    then F9 is not an alternative: Ctrl+A selects the main text story, so F9
    never reaches the header and footer stories. Fields inside text boxes are
    handled too, since those belong to the shape's text range rather than the
    header's own range. Pass -SkipFieldUpdate to leave it; if Word is not
    installed the script says so and the files are still correct, just showing
    cached values.

.PARAMETER File
    Path to a single Word file.

.PARAMETER Folder
    Folder of Word files. All .docx/.dotx/.docm/.dotm in folder are processed.

.PARAMETER FrontMatter
    Treat the target as front matter: leave the first-page footer blank.
    Without this switch (the chapter default) the first page carries a
    TITLE + PAGE footer. Run the front-matter file with -FrontMatter and the
    chapter files without it.

.PARAMETER Force
    Skip the confirmation prompt.

.PARAMETER SkipFieldUpdate
    Do not open the files in Word to recalculate header/footer fields. Use when
    Word is unavailable, or when a later step will refresh them anyway.

.EXAMPLE
    # Chapters: first page gets a footer
    .\Set-NISTHeaderFooterContent.ps1 -Folder "C:\chapters"

.EXAMPLE
    # Front matter: first page footer stays blank
    .\Set-NISTHeaderFooterContent.ps1 -File "NCSTAR-4-3_Chapter-00_Front-Matter.docx" -FrontMatter
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$FrontMatter,

    [switch]$SkipFieldUpdate,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# -------- Target list --------
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

# -------- Constants --------
$HDR_TYPE_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
$FTR_TYPE_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'
$HDR_CT       = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
$FTR_CT       = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# -------- XML builders for the six new parts --------
function New-HdrFirst {
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:pStyle w:val="Header"/></w:pPr>
  </w:p>
</w:hdr>
'@
}

function New-HdrEven {
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:pStyle w:val="HeaderEven"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> STYLEREF &quot;Heading 1&quot; \n \* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Chapter</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:hdr>
'@
}

function New-HdrDefault {
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:pStyle w:val="HeaderOdd"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> STYLEREF &quot;Heading 1&quot; \* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Chapter Title</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:hdr>
'@
}

function New-FtrFirst {
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:pStyle w:val="Footer"/></w:pPr>
  </w:p>
</w:ftr>
'@
}

function New-FtrFirstContent {
    # Chapter opening (first) page footer: recto layout - TITLE on inner (left),
    # PAGE on outer (right) - same as the odd/default footer, using FooterOdd so
    # it carries the horizontal rule and tab stops. Front matter uses the blank
    # New-FtrFirst instead (pass -FrontMatter).
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:pStyle w:val="FooterOdd"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> TITLE \* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Report Title</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:tab/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE \* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>
'@
}

function New-FtrEven {
    # Even (left-hand) page: PAGE on outer (left), TITLE on inner (right)
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:pStyle w:val="FooterEven"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE \* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:tab/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> TITLE \* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Report Title</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>
'@
}

function New-FtrDefault {
    # Odd (right-hand) page: TITLE on inner (left), PAGE on outer (right)
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:pStyle w:val="FooterOdd"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> TITLE \* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Report Title</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:tab/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE \* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>
'@
}

# -------- Confirm --------
Write-Host "NIST Header/Footer Content Installer" -ForegroundColor Cyan
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
$firstFooterMode = if ($FrontMatter) { 'blank (front matter)' } else { 'TITLE + PAGE (chapter)' }
Write-Host "  First-page footer: $firstFooterMode" -ForegroundColor Cyan
Write-Host "  This will REPLACE all existing header*.xml and footer*.xml parts." -ForegroundColor Yellow
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# -------- Per-file processing --------
foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    # Read all entries into memory
    $entries  = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
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

    # Decode the XML parts we need
    function Read-Text([string]$path) {
        if (-not $entries.ContainsKey($path)) { return $null }
        return $utf8NoBom.GetString($entries[$path])
    }

    $contentTypes = Read-Text '[Content_Types].xml'
    $documentRels = Read-Text 'word/_rels/document.xml.rels'
    $documentXml  = Read-Text 'word/document.xml'
    $settingsXml  = Read-Text 'word/settings.xml'

    if (-not $contentTypes) { Write-Warning "  No [Content_Types].xml. Skipping."; continue }
    if (-not $documentRels) { Write-Warning "  No word/_rels/document.xml.rels. Skipping."; continue }
    if (-not $documentXml)  { Write-Warning "  No word/document.xml. Skipping."; continue }

    # 1. Remove existing header*.xml / footer*.xml parts (and their .rels) from the entry set
    $hfPattern = '^word/(_rels/)?(header|footer)\d*\.xml(\.rels)?$'
    $toRemove = @($entries.Keys | Where-Object { $_ -match $hfPattern })
    foreach ($k in $toRemove) { [void]$entries.Remove($k) }
    Write-Host "  Removed $($toRemove.Count) existing header/footer part(s)." -ForegroundColor Gray

    # 2. Strip their <Override> entries from [Content_Types].xml.
    #    Match by ContentType (header+xml / footer+xml), NOT by part name, so
    #    this removes the original header#/footer# overrides AND any nist*
    #    overrides a previous run added - which keeps the step idempotent.
    #    Matching only header\d*/footer\d* names would leave the nist* overrides
    #    in place and a re-run would append a second copy, producing duplicate
    #    PartName entries that Word rejects as "unreadable content".
    #    The ContentType value contains "/", so the attribute skip must be
    #    [^>]* here - [^/]* stops at the first slash and matches nothing.
    $contentTypes = [regex]::Replace(
        $contentTypes,
        '<Override\b[^>]*ContentType="[^"]*wordprocessingml\.(?:header|footer)\+xml"[^>]*/>',
        ''
    )

    # 3. Strip their <Relationship> entries from document.xml.rels
    $documentRels = [regex]::Replace(
        $documentRels,
        '<Relationship\b[^>]*Type="[^"]*/(?:header|footer)"[^>]*/>',
        ''
    )

    # 4. Add <Override> entries for the six new parts
    $newOverrides =
        "<Override PartName=`"/word/nistHeaderFirst.xml`" ContentType=`"$HDR_CT`"/>" +
        "<Override PartName=`"/word/nistHeaderEven.xml`" ContentType=`"$HDR_CT`"/>" +
        "<Override PartName=`"/word/nistHeaderDefault.xml`" ContentType=`"$HDR_CT`"/>" +
        "<Override PartName=`"/word/nistFooterFirst.xml`" ContentType=`"$FTR_CT`"/>" +
        "<Override PartName=`"/word/nistFooterEven.xml`" ContentType=`"$FTR_CT`"/>" +
        "<Override PartName=`"/word/nistFooterDefault.xml`" ContentType=`"$FTR_CT`"/>"
    $contentTypes = [regex]::Replace($contentTypes, '</Types>', "$newOverrides</Types>", 1)

    # 5. Add <Relationship> entries for the six new parts (descriptive IDs)
    $newRels =
        "<Relationship Id=`"rIdNistHdrFirst`" Type=`"$HDR_TYPE_URI`" Target=`"nistHeaderFirst.xml`"/>" +
        "<Relationship Id=`"rIdNistHdrEven`" Type=`"$HDR_TYPE_URI`" Target=`"nistHeaderEven.xml`"/>" +
        "<Relationship Id=`"rIdNistHdrDefault`" Type=`"$HDR_TYPE_URI`" Target=`"nistHeaderDefault.xml`"/>" +
        "<Relationship Id=`"rIdNistFtrFirst`" Type=`"$FTR_TYPE_URI`" Target=`"nistFooterFirst.xml`"/>" +
        "<Relationship Id=`"rIdNistFtrEven`" Type=`"$FTR_TYPE_URI`" Target=`"nistFooterEven.xml`"/>" +
        "<Relationship Id=`"rIdNistFtrDefault`" Type=`"$FTR_TYPE_URI`" Target=`"nistFooterDefault.xml`"/>"
    $documentRels = [regex]::Replace($documentRels, '</Relationships>', "$newRels</Relationships>", 1)

    # 6. Update every sectPr in document.xml
    #    a. Strip existing <w:headerReference .../> and <w:footerReference .../> and <w:titlePg/>
    #    b. Insert new references at start, titlePg before close
    $sectPrInjectStart =
        '<w:headerReference w:type="first" r:id="rIdNistHdrFirst"/>' +
        '<w:headerReference w:type="even" r:id="rIdNistHdrEven"/>' +
        '<w:headerReference w:type="default" r:id="rIdNistHdrDefault"/>' +
        '<w:footerReference w:type="first" r:id="rIdNistFtrFirst"/>' +
        '<w:footerReference w:type="even" r:id="rIdNistFtrEven"/>' +
        '<w:footerReference w:type="default" r:id="rIdNistFtrDefault"/>'

    $sectPrEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $attrs = $m.Groups[1].Value
        $body  = $m.Groups[2].Value
        $body  = [regex]::Replace($body, '<w:headerReference\b[^/]*/>', '')
        $body  = [regex]::Replace($body, '<w:footerReference\b[^/]*/>', '')
        $body  = [regex]::Replace($body, '<w:titlePg\s*/>', '')
        "<w:sectPr$attrs>$sectPrInjectStart$body<w:titlePg/></w:sectPr>"
    }.GetNewClosure()
    $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
        $documentXml,
        '(?s)<w:sectPr\b([^>]*)>(.*?)</w:sectPr>',
        $sectPrEvaluator
    )
    # Convert any self-closing <w:sectPr/> into a populated one
    $documentXml = [System.Text.RegularExpressions.Regex]::Replace(
        $documentXml,
        '<w:sectPr\b([^>]*)/>',
        "<w:sectPr`$1>$sectPrInjectStart<w:titlePg/></w:sectPr>"
    )

    # 7. Update settings.xml (add evenAndOddHeaders + mirrorMargins; idempotent)
    if ($settingsXml) {
        if ($settingsXml -notmatch '<w:evenAndOddHeaders\s*/>') {
            $settingsXml = [regex]::Replace(
                $settingsXml,
                '(<w:settings\b[^>]*>)',
                '$1<w:evenAndOddHeaders/>',
                1
            )
        }
        if ($settingsXml -notmatch '<w:mirrorMargins\s*/>') {
            $settingsXml = [regex]::Replace(
                $settingsXml,
                '(<w:settings\b[^>]*>)',
                '$1<w:mirrorMargins/>',
                1
            )
        }
    }

    # 8. Stage the modified XML parts back into the entry dictionary
    $entries['[Content_Types].xml']            = $utf8NoBom.GetBytes($contentTypes)
    $entries['word/_rels/document.xml.rels']   = $utf8NoBom.GetBytes($documentRels)
    $entries['word/document.xml']              = $utf8NoBom.GetBytes($documentXml)
    if ($settingsXml) { $entries['word/settings.xml'] = $utf8NoBom.GetBytes($settingsXml) }

    # 9. Write the six new parts
    $entries['word/nistHeaderFirst.xml']    = $utf8NoBom.GetBytes((New-HdrFirst))
    $entries['word/nistHeaderEven.xml']     = $utf8NoBom.GetBytes((New-HdrEven))
    $entries['word/nistHeaderDefault.xml']  = $utf8NoBom.GetBytes((New-HdrDefault))
    $firstFooter = if ($FrontMatter) { New-FtrFirst } else { New-FtrFirstContent }
    $entries['word/nistFooterFirst.xml']    = $utf8NoBom.GetBytes($firstFooter)
    $entries['word/nistFooterEven.xml']     = $utf8NoBom.GetBytes((New-FtrEven))
    $entries['word/nistFooterDefault.xml']  = $utf8NoBom.GetBytes((New-FtrDefault))

    # 10. Write zip back
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "hfcontent-$([guid]::NewGuid()).zip")
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
    Write-Host "  Header/footer content installed." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Recalculate the header/footer fields
# ---------------------------------------------------------------------------
# Everything above edits the ZIP directly and needs no Word. This step does,
# unavoidably: STYLEREF and PAGE fields store a cached result, and only Word can
# compute a new one. Ctrl+A then F9 is not an alternative - Ctrl+A selects the
# main text story, so F9 never reaches the header and footer stories.
#
# Shapes are walked separately because a field inside a text box belongs to that
# shape's text range, not to the header's own range, so Range.Fields.Update()
# skips it.
if ($SkipFieldUpdate) {
    Write-Host ""
    Write-Host "Field update skipped (-SkipFieldUpdate). The headers and footers will show" -ForegroundColor Yellow
    Write-Host "stale values until Word recalculates them." -ForegroundColor Yellow
    return
}

$word = $null
try {
    $word = New-Object -ComObject Word.Application
} catch {
    Write-Host ""
    Write-Warning "Word is not available on this machine, so the fields were not recalculated."
    Write-Host "  The files are correct; their headers and footers just show cached values." -ForegroundColor Cyan
    Write-Host "  Re-run this script from Windows PowerShell, or open each file in Word and" -ForegroundColor Cyan
    Write-Host "  right-click in the header/footer > Update Field." -ForegroundColor Cyan
    return
}

Write-Host ""
Write-Host "Recalculating header/footer fields in Word..." -ForegroundColor Cyan
$word.Visible = $false
$word.DisplayAlerts = 0
$updatedFiles = 0
try {
    foreach ($target in $targets) {
        $doc = $null
        try {
            $doc = $word.Documents.Open($target, $false, $false)
            # Recalculating a field is not an editorial change. With track
            # changes on, Word would otherwise record every updated field as a
            # revision under the last author's name and bury it in the markup
            # of a document that is out for review.
            $wasTracking = $doc.TrackRevisions
            $doc.TrackRevisions = $false
            $n = 0
            foreach ($sec in $doc.Sections) {
                foreach ($collection in @($sec.Headers, $sec.Footers)) {
                    foreach ($hf in $collection) {
                        $n += $hf.Range.Fields.Count
                        $hf.Range.Fields.Update() | Out-Null
                        foreach ($shp in $hf.Shapes) {
                            if ($shp.TextFrame.HasText) {
                                $n += $shp.TextFrame.TextRange.Fields.Count
                                $shp.TextFrame.TextRange.Fields.Update() | Out-Null
                            }
                        }
                    }
                }
            }
            $doc.TrackRevisions = $wasTracking
            $doc.Save()
            $tracked = if ($wasTracking) { ', track changes left on' } else { '' }
            Write-Host ("  {0}: {1} field(s) across {2} section(s){3}" -f (Split-Path -Leaf $target), $n, $doc.Sections.Count, $tracked) -ForegroundColor Gray
            $updatedFiles++
        } catch {
            Write-Warning "  $(Split-Path -Leaf $target): field update failed - $($_.Exception.Message)"
        } finally {
            if ($doc) { $doc.Close($false) | Out-Null }
        }
    }
} finally {
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Fields recalculated in $updatedFiles of $($targets.Count) file(s)." -ForegroundColor Green
Write-Host "Title comes from File > Info > Properties > Title (set with Set-NISTDocTitle.ps1)." -ForegroundColor Cyan
