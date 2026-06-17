<#
.SYNOPSIS
    Combines a NIST NCSTAR or NWIRP series of chapter files into a single
    deliverable docx with front matter, body, and appendices in publication
    order.

.DESCRIPTION
    Discovers files matching <Series>*Chapter-*.docx and <Series>*Appendix-*.docx
    in -Folder, sorts them in publication order, and concatenates their body
    content into a single document.

    Sort order:
        Chapter-00 (front matter)
        Chapter-01, Chapter-02, ..., Chapter-NN
        Appendix-A, Appendix-B, ..., Appendix-Z

    The Chapter-00 file is treated as the host package. Its theme, styles,
    numbering, font table, headers and footers, and settings carry into the
    output. Every other chapter and appendix contributes body content only.

    Per-chapter section properties become section breaks between chapters.
    Front matter retains the page numbering format declared in Chapter-00's
    sectPr (typically lowerRoman). Chapter-01 is forced to pgNumType
    decimal start=1 so the body starts at page 1.

    Each section break in the output uses the host's headerReference and
    footerReference rIds. Per-chapter header / footer references are stripped
    so the deliverable has consistent headers and footers throughout.

    Image and hyperlink relationships are renumbered per chapter to avoid
    rId collisions, and embedded media files are renamed with a chapter
    prefix (ch01_, ch02_, ax_a_, etc.). Footnote and endnote ids are
    offset to avoid collisions.

    Comments are stripped from the output. Tracked revisions are accepted
    (insertions kept, deletions removed). TOC, LOF, and LOT field
    instructions are marked dirty so Word refreshes them when the file is
    next opened.

.PARAMETER Folder
    Source folder containing the chapter files.

.PARAMETER Series
    Publication series filename prefix. Match is case-insensitive and
    treats any non-word character after the prefix as a separator.
    Examples: NCSTAR_4-3, NWIRP_2-3.

.PARAMETER Output
    Path to write the combined docx. If a folder, the output filename is
    <Series>_Combined.docx inside that folder. If a .docx path, used as-is.

.PARAMETER Title
    Optional override for the dc:title metadata in docProps/core.xml.
    If omitted, the host's existing title is preserved.

.PARAMETER SkipStyleAudit
    Skip the Audit-NISTStyleSet.ps1 pre-assembly check. By default every
    source file is audited for the toolkit's required style ids first,
    and assembly refuses to start if any required style is missing in
    any input file. Use -SkipStyleAudit to bypass the audit (assembly
    will still produce output but downstream rendering may be wrong).

.PARAMETER Pdf
    After writing the combined docx, drive Microsoft Word via COM to refresh
    all fields (TOC, LOF, LOT, page numbers, cross-references), repaginate,
    and ExportAsFixedFormat to PDF next to the combined docx. Requires Word
    installed on this machine. The intermediate docx is preserved.

.PARAMETER KeepCombinedDocx
    Used with -Pdf. By default the intermediate combined docx is preserved.
    Pass -KeepCombinedDocx:$false to delete it after the PDF is produced.

.PARAMETER Force
    Overwrite the output without prompting.

.EXAMPLE
    .\Build-NISTReport.ps1 -Folder "C:\Users\Admin\Documents\260512" `
        -Series NCSTAR_4-3 -Output "C:\Users\Admin\Documents\260512\_assembled" -Pdf

.EXAMPLE
    .\Build-NISTReport.ps1 -Folder "C:\Users\Admin\Documents\260512" `
        -Series NWIRP_2-3 -Output "C:\out\NWIRP_2-3_Combined.docx" -Pdf -Force

.NOTES
    Source chapter files are read-only to this script. The combined docx
    and PDF are intermediate / output artifacts; the source chapters stay
    editable for future review rounds.

    Run Set-NISTStyleArchitecture.ps1 and Set-NISTHeaderFooterContent.ps1
    on every chapter file BEFORE running this script. Assembly assumes
    every chapter file is style-conformant; mismatched styles between
    chapters will combine awkwardly.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Folder,

    [Parameter(Mandatory=$true)]
    [string]$Series,

    [Parameter(Mandatory=$true)]
    [string]$Output,

    [string]$Title,

    [switch]$SkipStyleAudit,

    [switch]$Pdf,

    [switch]$KeepCombinedDocx = $true,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Path resolution
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
    throw "Folder not found: $Folder"
}
$folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath

if (Test-Path -LiteralPath $Output -PathType Container) {
    $outputPath = Join-Path $Output "${Series}_Combined.docx"
} else {
    $outputDir = Split-Path -LiteralPath $Output -Parent
    if (-not $outputDir) { $outputDir = (Get-Location).ProviderPath }
    if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    $outputPath = $Output
    if ($outputPath -notmatch '\.docx$') { $outputPath += '.docx' }
}

if ((Test-Path -LiteralPath $outputPath) -and -not $Force) {
    Write-Host "Output exists: $outputPath" -ForegroundColor Yellow
    $reply = Read-Host "Overwrite? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Discover and sort source files
# ---------------------------------------------------------------------------

$seriesEsc = [regex]::Escape($Series)

$chapterFiles = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
    Where-Object { $_.Extension -ieq '.docx' -and $_.Name -imatch "^${seriesEsc}\W" } |
    ForEach-Object {
        $kind = 'unknown'
        $key  = $null
        $sub  = 0
        if ($_.Name -imatch 'Chapter-(\d{1,2})') {
            $kind = 'chapter'
            $key  = [int]$matches[1]
        } elseif ($_.Name -imatch 'Appendix-([A-Z])') {
            $kind = 'appendix'
            $key  = [int][char]([string]$matches[1]).ToUpper()[0]
        }
        [pscustomobject]@{
            Path = $_.FullName
            Name = $_.Name
            Kind = $kind
            Key  = $key
        }
    } |
    Where-Object { $_.Kind -ne 'unknown' } |
    Sort-Object @{Expression={if ($_.Kind -eq 'chapter') {0} else {1}}}, Key

if (-not $chapterFiles -or $chapterFiles.Count -eq 0) {
    throw "No files matching '${Series}*Chapter-*.docx' or '${Series}*Appendix-*.docx' in $folderResolved"
}

# Chapter-00 must be present and first
$front = $chapterFiles | Where-Object { $_.Kind -eq 'chapter' -and $_.Key -eq 0 } | Select-Object -First 1
if (-not $front) {
    throw "Front matter file ('${Series}*Chapter-00*.docx') not found in $folderResolved"
}

Write-Host ""
Write-Host "Assembly plan:" -ForegroundColor Cyan
$idx = 0
foreach ($f in $chapterFiles) {
    $tag = if ($f.Kind -eq 'chapter' -and $f.Key -eq 0) { 'FRONT MATTER (host)' }
           elseif ($f.Kind -eq 'chapter') { "Chapter $($f.Key.ToString('D2'))" }
           else { "Appendix $([char]$f.Key)" }
    Write-Host ("  {0,2}. {1,-22} {2}" -f (++$idx), $tag, $f.Name)
}
Write-Host ""

# ---------------------------------------------------------------------------
# Pre-assembly style-set audit
# ---------------------------------------------------------------------------

if (-not $SkipStyleAudit) {
    $auditScript = Join-Path $PSScriptRoot 'Audit-NISTStyleSet.ps1'
    if (Test-Path -LiteralPath $auditScript) {
        Write-Host "Pre-assembly style audit..." -ForegroundColor Yellow
        $auditFailed = $false
        foreach ($cf in $chapterFiles) {
            $json   = & $auditScript -File $cf.Path -Json
            $result = $json | ConvertFrom-Json
            # Audit-NISTStyleSet wraps results in .Files[]; with a single file, take index 0.
            $fileResult     = $result.Files[0]
            $missCount      = @($fileResult.MissingRequired).Count
            $typeMissCount  = @($fileResult.TypeMismatch).Count
            $leaf           = Split-Path -Leaf $cf.Path
            if ($missCount -gt 0 -or $typeMissCount -gt 0) {
                $auditFailed = $true
                Write-Host ("  FAIL  {0}" -f $leaf) -ForegroundColor Red
                foreach ($mr in $fileResult.MissingRequired) {
                    Write-Host ("    missing required: {0,-26}  (needed by: {1})" -f $mr.StyleId, $mr.UsedBy) -ForegroundColor Red
                }
                foreach ($tm in $fileResult.TypeMismatch) {
                    Write-Host ("    type mismatch:    {0,-26}  expected {1}, found {2}" -f $tm.StyleId, $tm.Expected, $tm.Actual) -ForegroundColor Yellow
                }
            } else {
                Write-Host ("  OK    {0}" -f $leaf) -ForegroundColor DarkGray
            }
        }
        if ($auditFailed) {
            if ($Force) {
                Write-Warning "Style audit reported missing required styles. Continuing because -Force is set."
            } else {
                throw "Style audit failed. Fix the missing required styles, run with -Force to override, or pass -SkipStyleAudit."
            }
        } else {
            Write-Host "  Audit clean." -ForegroundColor Green
        }
        Write-Host ""
    } else {
        Write-Warning "Audit-NISTStyleSet.ps1 not found next to Build-NISTReport.ps1; skipping pre-assembly audit."
    }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Read-Entry {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if (-not $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-MaxRId {
    param([string]$RelsXml)
    if (-not $RelsXml) { return 0 }
    $matches = [regex]::Matches($RelsXml, 'Id="rId(\d+)"')
    if ($matches.Count -eq 0) { return 0 }
    return ($matches | ForEach-Object { [int]$_.Groups[1].Value } | Measure-Object -Maximum).Maximum
}

function Get-MaxAttrId {
    # Largest w:id="N" attribute (used for footnotes/endnotes/comments).
    param([string]$Xml, [string]$Pattern = 'w:id="(\d+)"')
    if (-not $Xml) { return 0 }
    $matches = [regex]::Matches($Xml, $Pattern)
    if ($matches.Count -eq 0) { return 0 }
    return ($matches | ForEach-Object { [int]$_.Groups[1].Value } | Measure-Object -Maximum).Maximum
}

function Get-BodyContent {
    # Returns (innerBody, terminalSectPr) for a document.xml.
    # innerBody = everything inside <w:body> EXCEPT the body-level terminal sectPr.
    # terminalSectPr = the body-level final <w:sectPr>...</w:sectPr>, or $null.
    param([string]$DocXml)
    $m = [regex]::Match($DocXml, '(?s)<w:body[^>]*>(.*)</w:body>')
    if (-not $m.Success) { throw "Could not locate <w:body> in document.xml" }
    $inner = $m.Groups[1].Value
    # Body-level sectPr is the LAST <w:sectPr>...</w:sectPr> not nested in <w:pPr>.
    $sectMatches = [regex]::Matches($inner, '(?s)<w:sectPr\b[^>]*>.*?</w:sectPr>')
    $terminal = $null
    if ($sectMatches.Count -gt 0) {
        $last = $sectMatches[$sectMatches.Count - 1]
        # Determine if it sits inside a <w:pPr> by checking what precedes it within a 200-char window.
        $startIdx = $last.Index
        $window = $inner.Substring([Math]::Max(0, $startIdx - 200), [Math]::Min(200, $startIdx))
        if ($window -notmatch '<w:pPr[^>]*>\s*$') {
            $terminal = $last.Value
            $inner = $inner.Remove($startIdx, $last.Length)
        }
    }
    return @{ Inner = $inner.Trim(); Terminal = $terminal }
}

function Strip-HeaderFooterRefs {
    param([string]$SectPrXml)
    if (-not $SectPrXml) { return $SectPrXml }
    return [regex]::Replace($SectPrXml, '<w:(headerReference|footerReference)[^/]*/>', '')
}

function Make-SectionBreakParagraph {
    # Wraps a sectPr inside an empty paragraph so it becomes an in-body section break.
    param([string]$SectPrXml)
    if (-not $SectPrXml) { return '' }
    return "<w:p><w:pPr>$SectPrXml</w:pPr></w:p>"
}

function Force-DecimalPage1 {
    param([string]$SectPrXml)
    if (-not $SectPrXml) { return $SectPrXml }
    $new = '<w:pgNumType w:fmt="decimal" w:start="1"/>'
    if ($SectPrXml -match '<w:pgNumType[^/]*/>') {
        return [regex]::Replace($SectPrXml, '<w:pgNumType[^/]*/>', $new)
    }
    # Insert before </w:sectPr>
    return $SectPrXml -replace '</w:sectPr>', "$new</w:sectPr>"
}

function Strip-Comments {
    # Removes all comment range markers and references from body XML.
    param([string]$Xml)
    if (-not $Xml) { return $Xml }
    $x = $Xml
    $x = [regex]::Replace($x, '<w:commentRangeStart[^/]*/>', '')
    $x = [regex]::Replace($x, '<w:commentRangeEnd[^/]*/>', '')
    $x = [regex]::Replace($x, '(?s)<w:r\b[^>]*>\s*(<w:rPr>.*?</w:rPr>\s*)?<w:commentReference[^/]*/>\s*</w:r>', '')
    $x = [regex]::Replace($x, '<w:commentReference[^/]*/>', '')
    return $x
}

function Accept-Revisions {
    # Strip insertion wrappers (keep their inner runs) and drop deletions entirely.
    # Also drops paragraph-level rPrChange / pPrChange / sectPrChange / tblPrChange tracking.
    param([string]$Xml)
    if (-not $Xml) { return $Xml }
    $x = $Xml
    # Drop entire deletion blocks (w:del with their contents).
    $x = [regex]::Replace($x, '(?s)<w:del\b[^>]*>.*?</w:del>', '')
    # Unwrap insertion blocks (keep inner content).
    $x = [regex]::Replace($x, '(?s)<w:ins\b[^>]*>(.*?)</w:ins>', '$1')
    # Strip change-tracking sub-elements.
    foreach ($tag in 'rPrChange','pPrChange','sectPrChange','tblPrChange','tcPrChange','trPrChange','tblGridChange','cellIns','cellDel','cellMerge') {
        $x = [regex]::Replace($x, "(?s)<w:$tag\b[^/]*?/>", '')
        $x = [regex]::Replace($x, "(?s)<w:$tag\b[^>]*>.*?</w:$tag>", '')
    }
    return $x
}

function Mark-FieldsDirty {
    # Mark TOC, LOF, LOT field instructions so Word refreshes them on open.
    # Adds w:dirty="true" attribute to the parent <w:fldChar w:fldCharType="begin"/>.
    param([string]$Xml)
    if (-not $Xml) { return $Xml }
    # Find every fldChar begin paired with an instrText that starts with TOC, LOF, LOT, TABLE OF, etc.
    return [regex]::Replace(
        $Xml,
        '(?s)<w:fldChar w:fldCharType="begin"(?![^>]*w:dirty)([^/]*)/>(\s*<w:[^>]+>)*\s*<w:instrText[^>]*>\s*(TOC|TABLE OF|LIST OF)\b',
        '<w:fldChar w:fldCharType="begin" w:dirty="true"$1/>$2 <w:instrText xml:space="preserve"> $3'
    )
}

function Renumber-Ids {
    # Replaces every occurrence of `w:id="N"` in $Xml using $map (hashtable N->newN).
    # Also handles footnoteReference, endnoteReference, and bookmark id forms.
    param([string]$Xml, [hashtable]$Map)
    if (-not $Xml -or $Map.Count -eq 0) { return $Xml }
    # Use callback so we only touch ids actually in the map.
    return [regex]::Replace($Xml, 'w:id="(\d+)"', {
        param($m)
        $old = [int]$m.Groups[1].Value
        if ($Map.ContainsKey($old)) { return 'w:id="' + $Map[$old] + '"' }
        return $m.Value
    })
}

function Renumber-Rids {
    # Replaces r:id="rIdN", r:embed="rIdN", r:link="rIdN" in $Xml using $map (hashtable oldRid->newRid).
    param([string]$Xml, [hashtable]$Map)
    if (-not $Xml -or $Map.Count -eq 0) { return $Xml }
    return [regex]::Replace($Xml, '(r:(?:id|embed|link)|r:dataPipe|r:saveThroughXslt)="(rId\d+)"', {
        param($m)
        $attr = $m.Groups[1].Value
        $old  = $m.Groups[2].Value
        if ($Map.ContainsKey($old)) { return "$attr=`"$($Map[$old])`"" }
        return $m.Value
    })
}

# ---------------------------------------------------------------------------
# Load host (Chapter-00) into memory
# ---------------------------------------------------------------------------

Write-Host "Loading host: $($front.Name)" -ForegroundColor Yellow

$hostEntries = @{}   # name -> byte[] (will be written as-is unless overridden later)
$hostStream = [System.IO.File]::Open($front.Path, 'Open', 'Read', 'Read')
try {
    $hostZip = New-Object System.IO.Compression.ZipArchive($hostStream, 'Read')
    try {
        foreach ($e in $hostZip.Entries) {
            $ms = New-Object System.IO.MemoryStream
            $es = $e.Open()
            try { $es.CopyTo($ms) } finally { $es.Dispose() }
            $hostEntries[$e.FullName] = $ms.ToArray()
        }
    } finally { $hostZip.Dispose() }
} finally { $hostStream.Dispose() }

function Get-HostText {
    param([string]$Name)
    if (-not $hostEntries.ContainsKey($Name)) { return $null }
    return [System.Text.Encoding]::UTF8.GetString($hostEntries[$Name])
}
function Set-HostText {
    param([string]$Name, [string]$Text)
    $hostEntries[$Name] = [System.Text.Encoding]::UTF8.GetBytes($Text)
}

$hostDocXml       = Get-HostText 'word/document.xml'
$hostRelsXml      = Get-HostText 'word/_rels/document.xml.rels'
$hostFootnotesXml = Get-HostText 'word/footnotes.xml'
$hostEndnotesXml  = Get-HostText 'word/endnotes.xml'

if (-not $hostDocXml)  { throw "Host has no word/document.xml" }
if (-not $hostRelsXml) {
    $hostRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
}

# Carry over header / footer references from host's body-level sectPr so each
# inserted section break uses the same headers and footers.
$hostBody = Get-BodyContent -DocXml $hostDocXml
$hostHfRefs = ''
if ($hostBody.Terminal) {
    $hfMatches = [regex]::Matches($hostBody.Terminal, '<w:(?:headerReference|footerReference)[^/]*/>')
    foreach ($hf in $hfMatches) { $hostHfRefs += $hf.Value }
}

# Running counters for id collisions.
$nextRId  = (Get-MaxRId  -RelsXml $hostRelsXml) + 1
$nextFnId = (Get-MaxAttrId -Xml $hostFootnotesXml) + 1
$nextEnId = (Get-MaxAttrId -Xml $hostEndnotesXml)  + 1
$nextBookmarkId = (Get-MaxAttrId -Xml $hostBody.Inner -Pattern 'w:bookmarkStart[^>]*\bw:id="(\d+)"') + 1

# Strip comments / accept revisions in the host body too.
$hostBodyInner = Accept-Revisions -Xml $hostBody.Inner
$hostBodyInner = Strip-Comments  -Xml $hostBodyInner

# Build the running body. Start with host content; for each subsequent file
# we append a section-break paragraph carrying the PREVIOUS file's terminal
# sectPr, then the next file's body content.
$assembledBody = $hostBodyInner
$prevTerminalSectPr = $hostBody.Terminal

# Accumulators that turn into host file entries at the end.
$newRels       = New-Object System.Collections.ArrayList
$newFootnotes  = New-Object System.Text.StringBuilder
$newEndnotes   = New-Object System.Text.StringBuilder
$newMedia      = @{}   # outputPathInPackage -> source byte[]

# ---------------------------------------------------------------------------
# Append each subsequent chapter / appendix
# ---------------------------------------------------------------------------

$followers = $chapterFiles | Where-Object { -not ($_.Kind -eq 'chapter' -and $_.Key -eq 0) }

foreach ($f in $followers) {
    Write-Host "Merging: $($f.Name)" -ForegroundColor Yellow

    $tagPrefix = if ($f.Kind -eq 'chapter') { "ch{0:D2}_" -f $f.Key } else { "ax_{0}_" -f ([char]$f.Key).ToString().ToLower() }

    # Load all entries from this file.
    $entries = @{}
    $fs = [System.IO.File]::Open($f.Path, 'Open', 'Read', 'Read')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, 'Read')
        try {
            foreach ($e in $zip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }

    function Get-EntryText { param([string]$Name) if ($entries.ContainsKey($Name)) { return [System.Text.Encoding]::UTF8.GetString($entries[$Name]) } else { return $null } }

    $docXml     = Get-EntryText 'word/document.xml'
    $relsXml    = Get-EntryText 'word/_rels/document.xml.rels'
    $fnXml      = Get-EntryText 'word/footnotes.xml'
    $enXml      = Get-EntryText 'word/endnotes.xml'
    if (-not $docXml) { Write-Warning "  No document.xml. Skipping."; continue }

    # 1. Build rId remap for this chapter.
    $rIdMap = @{}
    $skipRids = @{}   # rels we drop (header/footer/comments/people) — references in body get scrubbed
    if ($relsXml) {
        $relMatches = [regex]::Matches($relsXml, '<Relationship\s[^/]*/>')
        foreach ($rel in $relMatches) {
            $relXml = $rel.Value
            if ($relXml -notmatch 'Id="(rId\d+)"') { continue }
            $oldRid = $matches[1]
            if ($relXml -notmatch 'Type="([^"]+)"') { continue }
            $type = $matches[1]
            if ($relXml -notmatch 'Target="([^"]+)"') { continue }
            $target = $matches[1]
            $mode = $null
            if ($relXml -match 'TargetMode="([^"]+)"') { $mode = $matches[1] }

            # Drop header/footer/comments/people/glossary/numbering/styles/theme/fontTable/settings/webSettings/customXml refs:
            # the host already has these. Image and hyperlink rels remap.
            $shortType = $type.Split('/')[-1]
            switch -Wildcard ($shortType) {
                'header'        { $skipRids[$oldRid] = $true; continue }
                'footer'        { $skipRids[$oldRid] = $true; continue }
                'comments'      { $skipRids[$oldRid] = $true; continue }
                'commentsExtended' { $skipRids[$oldRid] = $true; continue }
                'commentsIds'   { $skipRids[$oldRid] = $true; continue }
                'people'        { $skipRids[$oldRid] = $true; continue }
                'theme'         { $skipRids[$oldRid] = $true; continue }
                'fontTable'     { $skipRids[$oldRid] = $true; continue }
                'styles'        { $skipRids[$oldRid] = $true; continue }
                'numbering'     { $skipRids[$oldRid] = $true; continue }
                'settings'      { $skipRids[$oldRid] = $true; continue }
                'webSettings'   { $skipRids[$oldRid] = $true; continue }
                'glossaryDocument' { $skipRids[$oldRid] = $true; continue }
                'customXml'     { $skipRids[$oldRid] = $true; continue }
                'footnotes'     { $skipRids[$oldRid] = $true; continue }
                'endnotes'      { $skipRids[$oldRid] = $true; continue }
            }

            $newRid = "rId$nextRId"
            $nextRId++
            $rIdMap[$oldRid] = $newRid

            switch -Wildcard ($shortType) {
                'image' {
                    # Copy + rename media file.
                    $srcEntry = "word/$target"
                    if ($target.StartsWith('../')) { $srcEntry = $target.Substring(3) }
                    if ($entries.ContainsKey($srcEntry)) {
                        $leaf = $target.Split('/')[-1]
                        $newLeaf = "${tagPrefix}${leaf}"
                        $newTarget = ($target.Substring(0, $target.Length - $leaf.Length)) + $newLeaf
                        $newMedia["word/$newTarget"] = $entries[$srcEntry]
                        $null = $newRels.Add("<Relationship Id=`"$newRid`" Type=`"$type`" Target=`"$newTarget`"/>")
                    } else {
                        Write-Warning "  Image referenced but not found in package: $srcEntry"
                        $null = $newRels.Add("<Relationship Id=`"$newRid`" Type=`"$type`" Target=`"$target`"/>")
                    }
                }
                'hyperlink' {
                    $tm = if ($mode) { " TargetMode=`"$mode`"" } else { '' }
                    $null = $newRels.Add("<Relationship Id=`"$newRid`" Type=`"$type`" Target=`"$target`"$tm/>")
                }
                default {
                    $tm = if ($mode) { " TargetMode=`"$mode`"" } else { '' }
                    $null = $newRels.Add("<Relationship Id=`"$newRid`" Type=`"$type`" Target=`"$target`"$tm/>")
                }
            }
        }
    }

    # 2. Extract body and remap rIds inside it.
    $body = Get-BodyContent -DocXml $docXml
    $inner = $body.Inner
    $terminal = $body.Terminal

    # Drop runs whose only payload was a header/footer/comments/etc. reference.
    if ($skipRids.Count -gt 0) {
        foreach ($drid in $skipRids.Keys) {
            $inner = [regex]::Replace($inner, "<w:(?:headerReference|footerReference|commentReference)[^>]*r:id=`"$drid`"[^/]*/>", '')
        }
    }
    $inner = Renumber-Rids -Xml $inner -Map $rIdMap
    if ($terminal) { $terminal = Renumber-Rids -Xml $terminal -Map $rIdMap }

    # 3. Renumber bookmark ids.
    $bookmarkMap = @{}
    $bmMatches = [regex]::Matches($inner, '<w:bookmarkStart[^>]*\bw:id="(\d+)"')
    foreach ($bm in $bmMatches) {
        $oldId = [int]$bm.Groups[1].Value
        if (-not $bookmarkMap.ContainsKey($oldId)) {
            $bookmarkMap[$oldId] = $nextBookmarkId
            $nextBookmarkId++
        }
    }
    if ($bookmarkMap.Count -gt 0) {
        $inner = [regex]::Replace($inner, '(<w:bookmark(?:Start|End)[^>]*\bw:id=")(\d+)(")', {
            param($m)
            $oldId = [int]$m.Groups[2].Value
            if ($bookmarkMap.ContainsKey($oldId)) {
                return $m.Groups[1].Value + $bookmarkMap[$oldId] + $m.Groups[3].Value
            }
            return $m.Value
        })
    }

    # 4. Footnote / endnote merge.
    if ($fnXml) {
        $fnNoteMatches = [regex]::Matches($fnXml, '(?s)<w:footnote\b(?![^>]*w:type)[^>]*\bw:id="(\d+)"[^>]*>.*?</w:footnote>')
        $fnMap = @{}
        $fnAdded = New-Object System.Text.StringBuilder
        foreach ($fn in $fnNoteMatches) {
            $oldId = [int]$fn.Groups[1].Value
            $newId = $nextFnId
            $nextFnId++
            $fnMap[$oldId] = $newId
            $noteXml = $fn.Value -replace ('w:id="' + $oldId + '"'), ('w:id="' + $newId + '"')
            # Remap r:embed / r:id inside the note.
            $noteXml = Renumber-Rids -Xml $noteXml -Map $rIdMap
            [void]$fnAdded.Append($noteXml)
        }
        if ($fnMap.Count -gt 0) {
            $inner = [regex]::Replace($inner, '<w:footnoteReference[^>]*\bw:id="(\d+)"', {
                param($m)
                $oldId = [int]$m.Groups[1].Value
                if ($fnMap.ContainsKey($oldId)) {
                    return ($m.Value -replace ('w:id="' + $oldId + '"'), ('w:id="' + $fnMap[$oldId] + '"'))
                }
                return $m.Value
            })
        }
        [void]$newFootnotes.Append($fnAdded.ToString())
    }
    if ($enXml) {
        $enNoteMatches = [regex]::Matches($enXml, '(?s)<w:endnote\b(?![^>]*w:type)[^>]*\bw:id="(\d+)"[^>]*>.*?</w:endnote>')
        $enMap = @{}
        $enAdded = New-Object System.Text.StringBuilder
        foreach ($en in $enNoteMatches) {
            $oldId = [int]$en.Groups[1].Value
            $newId = $nextEnId
            $nextEnId++
            $enMap[$oldId] = $newId
            $noteXml = $en.Value -replace ('w:id="' + $oldId + '"'), ('w:id="' + $newId + '"')
            $noteXml = Renumber-Rids -Xml $noteXml -Map $rIdMap
            [void]$enAdded.Append($noteXml)
        }
        if ($enMap.Count -gt 0) {
            $inner = [regex]::Replace($inner, '<w:endnoteReference[^>]*\bw:id="(\d+)"', {
                param($m)
                $oldId = [int]$m.Groups[1].Value
                if ($enMap.ContainsKey($oldId)) {
                    return ($m.Value -replace ('w:id="' + $oldId + '"'), ('w:id="' + $enMap[$oldId] + '"'))
                }
                return $m.Value
            })
        }
        [void]$newEndnotes.Append($enAdded.ToString())
    }

    # 5. Strip revisions and comments from this chapter's body.
    $inner = Accept-Revisions -Xml $inner
    $inner = Strip-Comments  -Xml $inner

    # 6. Force Chapter-01 to restart at decimal page 1.
    if ($f.Kind -eq 'chapter' -and $f.Key -eq 1 -and $terminal) {
        $terminal = Force-DecimalPage1 -SectPrXml $terminal
    }

    # 7. Section break before this chapter's content using the PREVIOUS terminal
    #    sectPr (with host header/footer refs). Then append this chapter's inner.
    $sectForBreak = $prevTerminalSectPr
    $sectForBreak = Strip-HeaderFooterRefs -SectPrXml $sectForBreak
    if ($hostHfRefs) {
        $sectForBreak = $sectForBreak -replace '</w:sectPr>', "$hostHfRefs</w:sectPr>"
    }
    $assembledBody += Make-SectionBreakParagraph -SectPrXml $sectForBreak
    $assembledBody += $inner

    # 8. Carry this chapter's terminal sectPr forward.
    $prevTerminalSectPr = $terminal
}

# ---------------------------------------------------------------------------
# Final assembly of document.xml
# ---------------------------------------------------------------------------

Write-Host "Stitching document.xml..." -ForegroundColor Yellow

# The very last terminal sectPr stays as the body-level terminal sectPr.
# Make sure its header/footer references match the host's.
if ($prevTerminalSectPr) {
    $finalSect = Strip-HeaderFooterRefs -SectPrXml $prevTerminalSectPr
    if ($hostHfRefs) {
        $finalSect = $finalSect -replace '</w:sectPr>', "$hostHfRefs</w:sectPr>"
    }
    $assembledBody += $finalSect
}

# Mark TOC / LOF / LOT field instructions dirty.
$assembledBody = Mark-FieldsDirty -Xml $assembledBody

# Rebuild full document.xml: keep host's <w:document> envelope and <w:body open tag>.
$envelopeMatch = [regex]::Match($hostDocXml, '(?s)^(.*?<w:body[^>]*>)')
if (-not $envelopeMatch.Success) { throw "Could not find <w:body> opening in host document.xml" }
$envelope = $envelopeMatch.Groups[1].Value
$newDocXml = "$envelope$assembledBody</w:body></w:document>"

Set-HostText 'word/document.xml' $newDocXml

# ---------------------------------------------------------------------------
# Rebuild document.xml.rels (host rels + accumulated new rels)
# ---------------------------------------------------------------------------

if ($newRels.Count -gt 0) {
    Write-Host ("Adding {0} new relationships..." -f $newRels.Count) -ForegroundColor Yellow
    $insert = ($newRels -join '')
    $hostRelsXml = $hostRelsXml -replace '</Relationships>', "$insert</Relationships>"
    Set-HostText 'word/_rels/document.xml.rels' $hostRelsXml
}

# ---------------------------------------------------------------------------
# Merge footnotes / endnotes
# ---------------------------------------------------------------------------

if ($newFootnotes.Length -gt 0) {
    if (-not $hostFootnotesXml) {
        $hostFootnotesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '</w:footnotes>'
    }
    $hostFootnotesXml = $hostFootnotesXml -replace '</w:footnotes>', ($newFootnotes.ToString() + '</w:footnotes>')
    Set-HostText 'word/footnotes.xml' $hostFootnotesXml
}
if ($newEndnotes.Length -gt 0) {
    if (-not $hostEndnotesXml) {
        $hostEndnotesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '</w:endnotes>'
    }
    $hostEndnotesXml = $hostEndnotesXml -replace '</w:endnotes>', ($newEndnotes.ToString() + '</w:endnotes>')
    Set-HostText 'word/endnotes.xml' $hostEndnotesXml
}

# ---------------------------------------------------------------------------
# Add merged media files
# ---------------------------------------------------------------------------

if ($newMedia.Count -gt 0) {
    Write-Host ("Adding {0} media files..." -f $newMedia.Count) -ForegroundColor Yellow
    foreach ($k in $newMedia.Keys) { $hostEntries[$k] = $newMedia[$k] }
}

# ---------------------------------------------------------------------------
# Drop comments parts from the package (we stripped references already)
# ---------------------------------------------------------------------------

$commentParts = @(
    'word/comments.xml',
    'word/commentsExtended.xml',
    'word/commentsIds.xml',
    'word/commentsExtensible.xml',
    'word/people.xml'
)
foreach ($p in $commentParts) {
    if ($hostEntries.ContainsKey($p)) { $hostEntries.Remove($p) | Out-Null }
}
# Also drop their entries from document.xml.rels if they survived (they shouldn't, but be safe).
$relsX = Get-HostText 'word/_rels/document.xml.rels'
if ($relsX) {
    $relsX = [regex]::Replace($relsX, '<Relationship[^/]*Target="(comments\w*\.xml|people\.xml)"[^/]*/>', '')
    Set-HostText 'word/_rels/document.xml.rels' $relsX
}
# And from [Content_Types].xml.
$ctText = Get-HostText '[Content_Types].xml'
if ($ctText) {
    foreach ($p in $commentParts) {
        $partPath = '/' + $p
        $ctText = [regex]::Replace($ctText, "<Override\s+PartName=`"$partPath`"[^/]*/>", '')
    }
    Set-HostText '[Content_Types].xml' $ctText
}

# ---------------------------------------------------------------------------
# Optional: override docProps/core.xml title
# ---------------------------------------------------------------------------

if ($Title) {
    $core = Get-HostText 'docProps/core.xml'
    if ($core) {
        if ($core -match '<dc:title>') {
            $core = [regex]::Replace($core, '(?s)<dc:title>.*?</dc:title>', "<dc:title>$Title</dc:title>")
        } else {
            $core = $core -replace '</cp:coreProperties>', "<dc:title>$Title</dc:title></cp:coreProperties>"
        }
        Set-HostText 'docProps/core.xml' $core
    }
}

# ---------------------------------------------------------------------------
# Write output zip
# ---------------------------------------------------------------------------

Write-Host "Writing $outputPath..." -ForegroundColor Yellow

$tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "nist-assemble-$([guid]::NewGuid()).zip")
$outStream = [System.IO.File]::Create($tempPath)
try {
    $outZip = New-Object System.IO.Compression.ZipArchive($outStream, 'Create')
    try {
        # Stable entry order: put [Content_Types].xml first if present, then everything else sorted.
        $names = @()
        if ($hostEntries.ContainsKey('[Content_Types].xml')) { $names += '[Content_Types].xml' }
        $names += ($hostEntries.Keys | Where-Object { $_ -ne '[Content_Types].xml' } | Sort-Object)
        foreach ($n in $names) {
            $newEntry = $outZip.CreateEntry($n, 'Optimal')
            $stream = $newEntry.Open()
            try { $stream.Write($hostEntries[$n], 0, $hostEntries[$n].Length) } finally { $stream.Dispose() }
        }
    } finally { $outZip.Dispose() }
} finally { $outStream.Dispose() }

Move-Item -LiteralPath $tempPath -Destination $outputPath -Force

Write-Host ""
Write-Host "Wrote: $outputPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Optional: Word COM -> PDF export
# ---------------------------------------------------------------------------

if ($Pdf) {
    $pdfPath = [System.IO.Path]::ChangeExtension($outputPath, '.pdf')
    Write-Host ""
    Write-Host "Driving Word to refresh fields and export PDF..." -ForegroundColor Yellow
    Write-Host "  Output PDF: $pdfPath" -ForegroundColor Cyan

    $word = $null
    $doc  = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0  # wdAlertsNone

        $doc = $word.Documents.Open($outputPath, $false, $false)  # ConfirmConversions=false, ReadOnly=false

        # Repaginate first so page numbers are stable before TOC/LOF/LOT build.
        $doc.Repaginate()

        # Update every field everywhere: body, headers, footers, footnotes, endnotes,
        # text frames, shapes, and every TOC, LOF, LOT explicitly.
        foreach ($story in $doc.StoryRanges) {
            $r = $story
            while ($null -ne $r) {
                try { $null = $r.Fields.Update() } catch {}
                $r = $r.NextStoryRange
            }
        }
        foreach ($toc in $doc.TablesOfContents) { try { $toc.Update() } catch {} }
        foreach ($tof in $doc.TablesOfFigures) { try { $tof.Update() } catch {} }
        # Repaginate again so post-TOC page changes settle.
        $doc.Repaginate()

        # ExportAsFixedFormat: wdExportFormatPDF=17, wdExportOptimizeForPrint=0,
        # wdExportAllDocument=0, wdExportDocumentWithMarkup=7, wdExportCreateHeadingBookmarks=1.
        $doc.ExportAsFixedFormat(
            $pdfPath,         # OutputFileName
            17,               # ExportFormat   (wdExportFormatPDF)
            $false,           # OpenAfterExport
            0,                # OptimizeFor    (wdExportOptimizeForPrint)
            0,                # Range          (wdExportAllDocument)
            0,                # From
            0,                # To
            7,                # Item           (wdExportDocumentWithMarkup) -- 7=DocumentContent; markup not exported because we already stripped revisions
            $true,            # IncludeDocProps
            $true,            # KeepIRM
            1,                # CreateBookmarks (wdExportCreateHeadingBookmarks)
            $true,            # DocStructureTags (accessibility)
            $true,            # BitmapMissingFonts
            $false            # UseISO19005_1 (PDF/A); set $true if PDF/A required
        )

        # Save the docx so refreshed fields persist for future reviews.
        $doc.Save()

        Write-Host "  PDF written." -ForegroundColor Green
    } finally {
        if ($doc)  { try { $doc.Close([ref]$false) } catch {} }
        if ($word) { try { $word.Quit() } catch {} }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc)  | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    if (-not $KeepCombinedDocx) {
        Remove-Item -LiteralPath $outputPath -Force
        Write-Host "  Removed intermediate docx." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host ""
Write-Host "Notes:" -ForegroundColor Cyan
Write-Host "  - Source chapter files were not modified. Continue editing them as separate docs." -ForegroundColor Cyan
Write-Host "  - For each review round, re-run this script to regenerate the combined deliverable." -ForegroundColor Cyan
if (-not $Pdf) {
    Write-Host "  - Pass -Pdf to also emit a PDF (requires Word installed)." -ForegroundColor Cyan
}
