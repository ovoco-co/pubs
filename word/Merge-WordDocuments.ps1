<#
.SYNOPSIS
    Combines several Word files into one document with a section break
    between each, without touching the originals.

.DESCRIPTION
    Takes every .docx in a folder (or an explicit ordered list of files),
    uses the FIRST file as the host package, and appends each following
    file's body content after a section break. The result is written to a
    new .docx; the source files are read-only to this script.

    The first file supplies the theme, styles, numbering, font table,
    settings, and header/footer parts for the whole output. Each following
    file contributes body content only; its own section geometry (page
    size, margins, orientation) is preserved as the properties of its
    section, but its header/footer references are unified to the host's so
    the combined document has consistent running heads throughout.

    Between documents, the previous document's section properties are
    emitted as an in-body section break. -SectionBreak sets the break type
    (nextPage by default, so each following document starts on a new page).

    To keep the output valid, the script:
      - renumbers relationship ids (rId), image/embed references, bookmark
        ids, and footnote/endnote ids per file to avoid collisions
      - copies and renames each file's media and embedded parts with a
        per-file prefix (d02_, d03_, ...)
      - registers a content-type Default for any new media extension the
        host did not already declare
      - accepts tracked revisions (insertions kept, deletions removed)
      - strips comments
      - marks TOC / list-of-figures / list-of-tables fields dirty so Word
        refreshes them on open

    Pure OOXML via System.IO.Compression - no Word installation required.

    Limitation: parts with their own nested relationships (for example
    charts, which reference further parts through _rels/chartN.xml.rels)
    are copied at the top level only; their sub-parts are not followed.
    Documents whose figures are pasted pictures merge cleanly.

.PARAMETER Folder
    Folder of Word files. Every .docx in it is combined, sorted by file
    name. Word lock files (~$*) and the output file are ignored.

.PARAMETER Files
    Explicit, ordered list of Word files to combine. Use this when you need
    a specific order or a subset. The first entry is the host.

.PARAMETER Output
    Path to write the combined .docx. If a folder, the file is written as
    <folder>\Combined.docx. If omitted with -Folder, defaults to
    <Folder>\Combined.docx.

.PARAMETER Title
    Optional dc:title metadata for the combined file. If omitted, the
    host's existing title is kept.

.PARAMETER SectionBreak
    Section break type between documents: nextPage (default), oddPage,
    evenPage, or continuous.

.PARAMETER Force
    Overwrite the output without prompting.

.EXAMPLE
    .\Merge-WordDocuments.ps1 -Folder "C:\Users\Artist\Documents\260729"

.EXAMPLE
    .\Merge-WordDocuments.ps1 -Folder $docs -Output "$docs\Full-Draft.docx" -Force

.EXAMPLE
    .\Merge-WordDocuments.ps1 -Files a.docx,b.docx,c.docx -Output out.docx -SectionBreak oddPage

.NOTES
    Source files are never modified. Re-run any time to regenerate the
    combined document.
#>

[CmdletBinding()]
param(
    [string]$Folder,

    [string[]]$Files,

    [string]$Output,

    [string]$Title,

    [ValidateSet('nextPage','oddPage','evenPage','continuous')]
    [string]$SectionBreak = 'nextPage',

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# -Folder and -Files are validated by hand rather than with parameter sets:
# with two mandatory-in-set params the binder is fragile, so pick the mode here.
if ($Folder -and $Files)             { throw "Specify either -Folder or -Files, not both." }
if (-not $Folder -and -not $Files)   { throw "Specify -Folder <dir> or -Files <file1,file2,...>." }
$mode = if ($Folder) { 'Folder' } else { 'Files' }

# ---------------------------------------------------------------------------
# Resolve output path first (so a folder scan can exclude it)
# ---------------------------------------------------------------------------

if ($mode -eq 'Folder') {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    if (-not $Output) { $Output = Join-Path $folderResolved 'Combined.docx' }
}

if (-not $Output) { throw "-Output is required when using -Files." }

if (Test-Path -LiteralPath $Output -PathType Container) {
    $outputPath = Join-Path ((Resolve-Path -LiteralPath $Output).ProviderPath) 'Combined.docx'
} else {
    $outputDir = Split-Path -LiteralPath $Output   # -Parent is the default; it is illegal to combine with -LiteralPath
    if (-not $outputDir) { $outputDir = (Get-Location).ProviderPath }
    if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    $outputDir = (Resolve-Path -LiteralPath $outputDir).ProviderPath
    $leafName  = Split-Path -Leaf $Output
    if ($leafName -notmatch '\.docx$') { $leafName += '.docx' }
    $outputPath = Join-Path $outputDir $leafName
}

# ---------------------------------------------------------------------------
# Build the ordered input list
# ---------------------------------------------------------------------------

if ($mode -eq 'Folder') {
    $inputs = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object {
            $_.Extension -ieq '.docx' -and
            $_.Name -notlike '~$*' -and
            $_.FullName -ne $outputPath
        } |
        Sort-Object Name |
        ForEach-Object { $_.FullName }
} else {
    $inputs = @()
    foreach ($f in $Files) {
        if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { throw "File not found: $f" }
        $inputs += (Resolve-Path -LiteralPath $f).ProviderPath
    }
}

$inputs = @($inputs)
if ($inputs.Count -eq 0) { throw "No .docx files to combine." }

Write-Host ""
Write-Host "Combine plan (section break: $SectionBreak):" -ForegroundColor Cyan
$idx = 0
foreach ($p in $inputs) {
    $tag = if ($idx -eq 0) { 'HOST' } else { '    ' }
    Write-Host ("  {0,2}. [{1}] {2}" -f (++$idx), $tag, (Split-Path -Leaf $p))
}
Write-Host "  -> $outputPath" -ForegroundColor Cyan
Write-Host ""

if ($inputs.Count -eq 1) { Write-Warning "Only one input file; the output will be a copy of it." }

if ((Test-Path -LiteralPath $outputPath) -and -not $Force) {
    $reply = Read-Host "Output exists. Overwrite? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Content types for media extensions (used when a follower brings a new one)
# ---------------------------------------------------------------------------

$ctByExt = @{
    'png'='image/png'; 'jpg'='image/jpeg'; 'jpeg'='image/jpeg'; 'jpe'='image/jpeg'
    'gif'='image/gif'; 'bmp'='image/bmp'; 'tif'='image/tiff'; 'tiff'='image/tiff'
    'emf'='image/x-emf'; 'wmf'='image/x-wmf'; 'svg'='image/svg+xml'; 'ico'='image/x-icon'
    'bin'='application/vnd.openxmlformats-officedocument.oleObject'
    'xlsx'='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    'docx'='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    'pptx'='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    'pdf'='application/pdf'
}

# ---------------------------------------------------------------------------
# Helpers (shared approach with Build-NISTReport.ps1)
# ---------------------------------------------------------------------------

function Get-MaxRId {
    param([string]$RelsXml)
    if (-not $RelsXml) { return 0 }
    $m = [regex]::Matches($RelsXml, 'Id="rId(\d+)"')
    if ($m.Count -eq 0) { return 0 }
    return ($m | ForEach-Object { [int]$_.Groups[1].Value } | Measure-Object -Maximum).Maximum
}

function Get-MaxAttrId {
    param([string]$Xml, [string]$Pattern = 'w:id="(\d+)"')
    if (-not $Xml) { return 0 }
    $m = [regex]::Matches($Xml, $Pattern)
    if ($m.Count -eq 0) { return 0 }
    return ($m | ForEach-Object { [int]$_.Groups[1].Value } | Measure-Object -Maximum).Maximum
}

function Get-BodyContent {
    # Returns @{ Inner; Terminal } for a document.xml. Terminal = body-level
    # final <w:sectPr> not nested in a <w:pPr>, or $null.
    param([string]$DocXml)
    $m = [regex]::Match($DocXml, '(?s)<w:body[^>]*>(.*)</w:body>')
    if (-not $m.Success) { throw "Could not locate <w:body> in document.xml" }
    $inner = $m.Groups[1].Value
    $sectMatches = [regex]::Matches($inner, '(?s)<w:sectPr\b[^>]*>.*?</w:sectPr>')
    $terminal = $null
    if ($sectMatches.Count -gt 0) {
        $last = $sectMatches[$sectMatches.Count - 1]
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

function Add-HeaderFooterRefs {
    # Insert header/footer references as the FIRST children of the sectPr - where
    # the OOXML schema requires them (headerReference*, footerReference* come
    # before type/pgSz/...). Appending them before </w:sectPr> puts them after
    # pgSz and Word rejects the file as "unreadable content".
    param([string]$SectPrXml, [string]$Refs)
    if (-not $SectPrXml -or -not $Refs) { return $SectPrXml }
    return [regex]::Replace($SectPrXml, '(<w:sectPr\b[^>]*>)', "`$1$Refs", 1)
}

function Set-SectBreakType {
    param([string]$SectPrXml, [string]$Type)
    if (-not $SectPrXml -or -not $Type) { return $SectPrXml }
    $s = [regex]::Replace($SectPrXml, '<w:type\b[^/]*/>', '')
    $typeEl = "<w:type w:val=`"$Type`"/>"
    if ($s -match '<w:pgSz\b') {
        return [regex]::Replace($s, '(<w:pgSz\b)', "$typeEl`$1", 1)
    }
    return [regex]::Replace($s, '(<w:sectPr\b[^>]*>)', "`$1$typeEl", 1)
}

function Make-SectionBreakParagraph {
    param([string]$SectPrXml)
    if (-not $SectPrXml) { return '' }
    return "<w:p><w:pPr>$SectPrXml</w:pPr></w:p>"
}

function Strip-Comments {
    param([string]$Xml)
    if (-not $Xml) { return $Xml }
    $x = $Xml
    $x = [regex]::Replace($x, '<w:commentRangeStart[^/]*/>', '')
    $x = [regex]::Replace($x, '<w:commentRangeEnd[^/]*/>', '')
    # Remove only the reference element, leaving a harmless (possibly empty) run.
    # Do NOT try to match the whole enclosing run: a pattern like
    # <w:r ...>(<w:rPr>.*?</w:rPr>)?<w:commentReference/></w:r> lets the rPr .*?
    # bridge from one run across entire table cells to a distant comment run,
    # deleting thousands of characters and breaking table structure.
    $x = [regex]::Replace($x, '<w:commentReference[^/]*/>', '')
    return $x
}

function Accept-Revisions {
    param([string]$Xml)
    if (-not $Xml) { return $Xml }
    $x = $Xml
    # Paragraph-mark change markers are self-closing (<w:ins .../>, <w:del .../>)
    # sitting inside an rPr. Strip these FIRST: otherwise the block unwrap below
    # mistakes a self-closing <w:ins .../> for an opening tag and swallows content
    # up to the next real </w:ins>, corrupting the XML.
    $x = [regex]::Replace($x, '<w:ins\b[^>]*?/>', '')
    $x = [regex]::Replace($x, '<w:del\b[^>]*?/>', '')
    $x = [regex]::Replace($x, '(?s)<w:del\b[^>]*>.*?</w:del>', '')
    $x = [regex]::Replace($x, '(?s)<w:ins\b[^>]*>(.*?)</w:ins>', '$1')
    foreach ($tag in 'rPrChange','pPrChange','sectPrChange','tblPrChange','tcPrChange','trPrChange','tblGridChange','cellIns','cellDel','cellMerge') {
        $x = [regex]::Replace($x, "(?s)<w:$tag\b[^/]*?/>", '')
        $x = [regex]::Replace($x, "(?s)<w:$tag\b[^>]*>.*?</w:$tag>", '')
    }
    return $x
}

function Mark-FieldsDirty {
    param([string]$Xml)
    if (-not $Xml) { return $Xml }
    return [regex]::Replace(
        $Xml,
        '(?s)<w:fldChar w:fldCharType="begin"(?![^>]*w:dirty)([^/]*)/>(\s*<w:[^>]+>)*\s*<w:instrText[^>]*>\s*(TOC|TABLE OF|LIST OF)\b',
        '<w:fldChar w:fldCharType="begin" w:dirty="true"$1/>$2 <w:instrText xml:space="preserve"> $3'
    )
}

function Renumber-Rids {
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
# Load host (first file) into memory
# ---------------------------------------------------------------------------

$hostPath = $inputs[0]
Write-Host "Loading host: $(Split-Path -Leaf $hostPath)" -ForegroundColor Yellow

$hostEntries = @{}
$hostStream = [System.IO.File]::Open($hostPath, 'Open', 'Read', 'Read')
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

function Get-HostText { param([string]$Name)
    if (-not $hostEntries.ContainsKey($Name)) { return $null }
    return [System.Text.Encoding]::UTF8.GetString($hostEntries[$Name])
}
function Set-HostText { param([string]$Name, [string]$Text)
    $hostEntries[$Name] = [System.Text.Encoding]::UTF8.GetBytes($Text)
}

$hostDocXml       = Get-HostText 'word/document.xml'
$hostRelsXml      = Get-HostText 'word/_rels/document.xml.rels'
$hostFootnotesXml = Get-HostText 'word/footnotes.xml'
$hostEndnotesXml  = Get-HostText 'word/endnotes.xml'

if (-not $hostDocXml) { throw "Host has no word/document.xml" }
if (-not $hostRelsXml) {
    $hostRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
}

$hostBody = Get-BodyContent -DocXml $hostDocXml
$hostHfRefs = ''
if ($hostBody.Terminal) {
    foreach ($hf in [regex]::Matches($hostBody.Terminal, '<w:(?:headerReference|footerReference)[^/]*/>')) {
        $hostHfRefs += $hf.Value
    }
}

$nextRId        = (Get-MaxRId  -RelsXml $hostRelsXml) + 1
$nextFnId       = (Get-MaxAttrId -Xml $hostFootnotesXml) + 1
$nextEnId       = (Get-MaxAttrId -Xml $hostEndnotesXml)  + 1
$nextBookmarkId = (Get-MaxAttrId -Xml $hostBody.Inner -Pattern 'w:bookmarkStart[^>]*\bw:id="(\d+)"') + 1

$hostBodyInner = Accept-Revisions -Xml $hostBody.Inner
$hostBodyInner = Strip-Comments  -Xml $hostBodyInner

$assembledBody = $hostBodyInner
$prevTerminalSectPr = $hostBody.Terminal

$newRels      = New-Object System.Collections.ArrayList
$newFootnotes = New-Object System.Text.StringBuilder
$newEndnotes  = New-Object System.Text.StringBuilder
$newMedia     = @{}   # package path -> byte[]

# ---------------------------------------------------------------------------
# Append each following file
# ---------------------------------------------------------------------------

for ($fi = 1; $fi -lt $inputs.Count; $fi++) {
    $fPath = $inputs[$fi]
    Write-Host "Merging: $(Split-Path -Leaf $fPath)" -ForegroundColor Yellow
    $tagPrefix = "d{0:D2}_" -f ($fi + 1)

    $entries = @{}
    $fs = [System.IO.File]::Open($fPath, 'Open', 'Read', 'Read')
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

    function Get-EntryText { param([string]$Name)
        if ($entries.ContainsKey($Name)) { return [System.Text.Encoding]::UTF8.GetString($entries[$Name]) } else { return $null }
    }

    $docXml  = Get-EntryText 'word/document.xml'
    $relsXml = Get-EntryText 'word/_rels/document.xml.rels'
    $fnXml   = Get-EntryText 'word/footnotes.xml'
    $enXml   = Get-EntryText 'word/endnotes.xml'
    if (-not $docXml) { Write-Warning "  No document.xml. Skipping."; continue }

    # 1. Remap relationships. Skip parts the host already provides; copy media
    #    and embedded parts with a per-file prefix; keep hyperlinks/externals.
    $rIdMap = @{}
    $skipRids = @{}
    if ($relsXml) {
        # [^>]* not [^/]*: every Relationship Type is an http:// URL full of
        # slashes, so [^/]* stops at the first slash and matches nothing.
        foreach ($rel in [regex]::Matches($relsXml, '<Relationship\s[^>]*/>')) {
            $relXml = $rel.Value
            if ($relXml -notmatch 'Id="(rId\d+)"')  { continue }; $oldRid = $matches[1]
            if ($relXml -notmatch 'Type="([^"]+)"') { continue }; $type   = $matches[1]
            if ($relXml -notmatch 'Target="([^"]+)"'){ continue }; $target = $matches[1]
            $mode = $null; if ($relXml -match 'TargetMode="([^"]+)"') { $mode = $matches[1] }
            $shortType = $type.Split('/')[-1]

            switch -Wildcard ($shortType) {
                'header'           { $skipRids[$oldRid] = $true }
                'footer'           { $skipRids[$oldRid] = $true }
                'comments'         { $skipRids[$oldRid] = $true }
                'commentsExtended' { $skipRids[$oldRid] = $true }
                'commentsIds'      { $skipRids[$oldRid] = $true }
                'commentsExtensible' { $skipRids[$oldRid] = $true }
                'people'           { $skipRids[$oldRid] = $true }
                'theme'            { $skipRids[$oldRid] = $true }
                'fontTable'        { $skipRids[$oldRid] = $true }
                'styles'           { $skipRids[$oldRid] = $true }
                'numbering'        { $skipRids[$oldRid] = $true }
                'settings'         { $skipRids[$oldRid] = $true }
                'webSettings'      { $skipRids[$oldRid] = $true }
                'glossaryDocument' { $skipRids[$oldRid] = $true }
                'customXml'        { $skipRids[$oldRid] = $true }
                'footnotes'        { $skipRids[$oldRid] = $true }
                'endnotes'         { $skipRids[$oldRid] = $true }
            }
            if ($skipRids.ContainsKey($oldRid)) { continue }

            $newRid = "rId$nextRId"; $nextRId++
            $rIdMap[$oldRid] = $newRid

            if ($mode -eq 'External' -or $shortType -eq 'hyperlink') {
                $tm = if ($mode) { " TargetMode=`"$mode`"" } else { '' }
                $null = $newRels.Add("<Relationship Id=`"$newRid`" Type=`"$type`" Target=`"$target`"$tm/>")
                continue
            }

            # Internal part: copy+rename if it lives under word/ and exists.
            if (-not $target.StartsWith('../')) {
                $srcEntry = "word/$target"
                if ($entries.ContainsKey($srcEntry)) {
                    $leaf      = $target.Split('/')[-1]
                    $newLeaf   = "$tagPrefix$leaf"
                    $newTarget = ($target.Substring(0, $target.Length - $leaf.Length)) + $newLeaf
                    $newMedia["word/$newTarget"] = $entries[$srcEntry]
                    $null = $newRels.Add("<Relationship Id=`"$newRid`" Type=`"$type`" Target=`"$newTarget`"/>")
                    continue
                }
            }
            # Shared part (../) or not found: reference as-is; host is expected to provide it.
            $null = $newRels.Add("<Relationship Id=`"$newRid`" Type=`"$type`" Target=`"$target`"/>")
        }
    }

    # 2. Body: strip skipped refs, remap rIds.
    $body     = Get-BodyContent -DocXml $docXml
    $inner    = $body.Inner
    $terminal = $body.Terminal
    foreach ($drid in $skipRids.Keys) {
        $inner = [regex]::Replace($inner, "<w:(?:headerReference|footerReference|commentReference)[^>]*r:id=`"$drid`"[^/]*/>", '')
    }
    $inner = Renumber-Rids -Xml $inner -Map $rIdMap
    if ($terminal) { $terminal = Renumber-Rids -Xml $terminal -Map $rIdMap }

    # 3. Renumber bookmark ids.
    $bookmarkMap = @{}
    foreach ($bm in [regex]::Matches($inner, '<w:bookmarkStart[^>]*\bw:id="(\d+)"')) {
        $oldId = [int]$bm.Groups[1].Value
        if (-not $bookmarkMap.ContainsKey($oldId)) { $bookmarkMap[$oldId] = $nextBookmarkId; $nextBookmarkId++ }
    }
    if ($bookmarkMap.Count -gt 0) {
        $inner = [regex]::Replace($inner, '(<w:bookmark(?:Start|End)[^>]*\bw:id=")(\d+)(")', {
            param($m)
            $oldId = [int]$m.Groups[2].Value
            if ($bookmarkMap.ContainsKey($oldId)) { return $m.Groups[1].Value + $bookmarkMap[$oldId] + $m.Groups[3].Value }
            return $m.Value
        })
    }

    # 4. Footnote / endnote merge.
    if ($fnXml) {
        $fnMap = @{}; $fnAdded = New-Object System.Text.StringBuilder
        foreach ($fn in [regex]::Matches($fnXml, '(?s)<w:footnote\b(?![^>]*w:type)[^>]*\bw:id="(\d+)"[^>]*>.*?</w:footnote>')) {
            $oldId = [int]$fn.Groups[1].Value; $newId = $nextFnId; $nextFnId++; $fnMap[$oldId] = $newId
            $noteXml = $fn.Value -replace ('w:id="' + $oldId + '"'), ('w:id="' + $newId + '"')
            $noteXml = Renumber-Rids -Xml $noteXml -Map $rIdMap
            [void]$fnAdded.Append($noteXml)
        }
        if ($fnMap.Count -gt 0) {
            $inner = [regex]::Replace($inner, '<w:footnoteReference[^>]*\bw:id="(\d+)"', {
                param($m)
                $oldId = [int]$m.Groups[1].Value
                if ($fnMap.ContainsKey($oldId)) { return ($m.Value -replace ('w:id="' + $oldId + '"'), ('w:id="' + $fnMap[$oldId] + '"')) }
                return $m.Value
            })
        }
        [void]$newFootnotes.Append($fnAdded.ToString())
    }
    if ($enXml) {
        $enMap = @{}; $enAdded = New-Object System.Text.StringBuilder
        foreach ($en in [regex]::Matches($enXml, '(?s)<w:endnote\b(?![^>]*w:type)[^>]*\bw:id="(\d+)"[^>]*>.*?</w:endnote>')) {
            $oldId = [int]$en.Groups[1].Value; $newId = $nextEnId; $nextEnId++; $enMap[$oldId] = $newId
            $noteXml = $en.Value -replace ('w:id="' + $oldId + '"'), ('w:id="' + $newId + '"')
            $noteXml = Renumber-Rids -Xml $noteXml -Map $rIdMap
            [void]$enAdded.Append($noteXml)
        }
        if ($enMap.Count -gt 0) {
            $inner = [regex]::Replace($inner, '<w:endnoteReference[^>]*\bw:id="(\d+)"', {
                param($m)
                $oldId = [int]$m.Groups[1].Value
                if ($enMap.ContainsKey($oldId)) { return ($m.Value -replace ('w:id="' + $oldId + '"'), ('w:id="' + $enMap[$oldId] + '"')) }
                return $m.Value
            })
        }
        [void]$newEndnotes.Append($enAdded.ToString())
    }

    # 5. Accept revisions, strip comments.
    $inner = Accept-Revisions -Xml $inner
    $inner = Strip-Comments  -Xml $inner

    # 6. Section break carrying the PREVIOUS file's section properties, with
    #    the host's header/footer refs and the chosen break type.
    $sectForBreak = if ($prevTerminalSectPr) { $prevTerminalSectPr } else { '<w:sectPr></w:sectPr>' }
    $sectForBreak = Strip-HeaderFooterRefs -SectPrXml $sectForBreak
    $sectForBreak = Set-SectBreakType -SectPrXml $sectForBreak -Type $SectionBreak
    $sectForBreak = Add-HeaderFooterRefs -SectPrXml $sectForBreak -Refs $hostHfRefs
    $assembledBody += Make-SectionBreakParagraph -SectPrXml $sectForBreak
    $assembledBody += $inner

    $prevTerminalSectPr = $terminal
}

# ---------------------------------------------------------------------------
# Final assembly
# ---------------------------------------------------------------------------

Write-Host "Stitching document.xml..." -ForegroundColor Yellow

if ($prevTerminalSectPr) {
    $finalSect = Strip-HeaderFooterRefs -SectPrXml $prevTerminalSectPr
    $finalSect = Set-SectBreakType -SectPrXml $finalSect -Type $SectionBreak
    $finalSect = Add-HeaderFooterRefs -SectPrXml $finalSect -Refs $hostHfRefs
    $assembledBody += $finalSect
}

$assembledBody = Mark-FieldsDirty -Xml $assembledBody

$envelopeMatch = [regex]::Match($hostDocXml, '(?s)^(.*?<w:body[^>]*>)')
if (-not $envelopeMatch.Success) { throw "Could not find <w:body> opening in host document.xml" }
$newDocXml = "$($envelopeMatch.Groups[1].Value)$assembledBody</w:body></w:document>"
Set-HostText 'word/document.xml' $newDocXml

if ($newRels.Count -gt 0) {
    Write-Host ("Adding {0} new relationships..." -f $newRels.Count) -ForegroundColor Yellow
    $hostRelsXml = $hostRelsXml -replace '</Relationships>', (($newRels -join '') + '</Relationships>')
    Set-HostText 'word/_rels/document.xml.rels' $hostRelsXml
}

if ($newFootnotes.Length -gt 0) {
    if (-not $hostFootnotesXml) {
        $hostFootnotesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:footnotes>'
    }
    $hostFootnotesXml = $hostFootnotesXml -replace '</w:footnotes>', ($newFootnotes.ToString() + '</w:footnotes>')
    Set-HostText 'word/footnotes.xml' $hostFootnotesXml
}
if ($newEndnotes.Length -gt 0) {
    if (-not $hostEndnotesXml) {
        $hostEndnotesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:endnotes>'
    }
    $hostEndnotesXml = $hostEndnotesXml -replace '</w:endnotes>', ($newEndnotes.ToString() + '</w:endnotes>')
    Set-HostText 'word/endnotes.xml' $hostEndnotesXml
}

if ($newMedia.Count -gt 0) {
    Write-Host ("Adding {0} media/embedded part(s)..." -f $newMedia.Count) -ForegroundColor Yellow
    foreach ($k in $newMedia.Keys) { $hostEntries[$k] = $newMedia[$k] }

    # Register a content-type Default for any new extension the host lacks.
    $ct = Get-HostText '[Content_Types].xml'
    if ($ct) {
        $have = @{}
        foreach ($d in [regex]::Matches($ct, '<Default\s+Extension="([^"]+)"')) { $have[$d.Groups[1].Value.ToLower()] = $true }
        $addDefaults = ''; $seenExt = @{}
        foreach ($k in $newMedia.Keys) {
            $ext = [System.IO.Path]::GetExtension($k).TrimStart('.').ToLower()
            if (-not $ext -or $have.ContainsKey($ext) -or $seenExt.ContainsKey($ext)) { continue }
            $seenExt[$ext] = $true
            $ctype = if ($ctByExt.ContainsKey($ext)) { $ctByExt[$ext] } else { 'application/octet-stream' }
            $addDefaults += "<Default Extension=`"$ext`" ContentType=`"$ctype`"/>"
        }
        if ($addDefaults) {
            $ct = [regex]::Replace($ct, '(<Types\b[^>]*>)', "`$1$addDefaults", 1)
            Set-HostText '[Content_Types].xml' $ct
        }
    }
}

# ---------------------------------------------------------------------------
# Drop comment parts (references were stripped) so the package stays valid
# ---------------------------------------------------------------------------

$commentParts = @('word/comments.xml','word/commentsExtended.xml','word/commentsIds.xml','word/commentsExtensible.xml','word/people.xml')
foreach ($p in $commentParts) { if ($hostEntries.ContainsKey($p)) { $hostEntries.Remove($p) | Out-Null } }
$relsX = Get-HostText 'word/_rels/document.xml.rels'
if ($relsX) {
    $relsX = [regex]::Replace($relsX, '<Relationship\b[^>]*Target="(comments\w*\.xml|people\.xml)"[^>]*/>', '')
    Set-HostText 'word/_rels/document.xml.rels' $relsX
}
$ctText = Get-HostText '[Content_Types].xml'
if ($ctText) {
    foreach ($p in $commentParts) {
        # [^>]* not [^/]*: PartName is fine but ContentType is a slash-filled URL.
        $ctText = [regex]::Replace($ctText, "<Override\s+PartName=`"/$p`"[^>]*/>", '')
    }
    Set-HostText '[Content_Types].xml' $ctText
}

# ---------------------------------------------------------------------------
# Optional title override
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
# Write output
# ---------------------------------------------------------------------------

Write-Host "Writing $outputPath..." -ForegroundColor Yellow

$tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "word-merge-$([guid]::NewGuid()).zip")
$outStream = [System.IO.File]::Create($tempPath)
try {
    $outZip = New-Object System.IO.Compression.ZipArchive($outStream, 'Create')
    try {
        $names = @()
        if ($hostEntries.ContainsKey('[Content_Types].xml')) { $names += '[Content_Types].xml' }
        $names += ($hostEntries.Keys | Where-Object { $_ -ne '[Content_Types].xml' } | Sort-Object)
        foreach ($n in $names) {
            $ne = $outZip.CreateEntry($n, 'Optimal')
            $os = $ne.Open()
            try { $os.Write($hostEntries[$n], 0, $hostEntries[$n].Length) } finally { $os.Dispose() }
        }
    } finally { $outZip.Dispose() }
} finally { $outStream.Dispose() }

Move-Item -LiteralPath $tempPath -Destination $outputPath -Force

Write-Host ""
Write-Host "Wrote: $outputPath" -ForegroundColor Green
Write-Host ""
Write-Host "Notes:" -ForegroundColor Cyan
Write-Host "  - Source files were not modified." -ForegroundColor Cyan
Write-Host "  - Headers/footers come from the first file; comments were removed and" -ForegroundColor Cyan
Write-Host "    tracked changes accepted." -ForegroundColor Cyan
Write-Host "  - Open in Word and press Ctrl+A then F9 to refresh the table of contents" -ForegroundColor Cyan
Write-Host "    and any other fields." -ForegroundColor Cyan
