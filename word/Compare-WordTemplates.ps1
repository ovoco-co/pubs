<#
.SYNOPSIS
    Compares Word templates (or documents) in a folder for drift, by reading
    the underlying OOXML directly. No Word installation required.

.DESCRIPTION
    Opens every .docx, .dotx, .docm, .dotm file in a folder as a zip archive,
    extracts styles (paragraph, character, table, list/numbering), header and
    footer content, building blocks, content controls, fields, bookmarks, and
    custom XML parts directly from the OOXML. Writes a per-file JSON snapshot,
    plus a drift report (CSV + HTML) showing where the files disagree.

    Designed for verifying that several templates intended to be identical
    have not drifted apart. Works on Windows, WSL, Linux, or macOS with
    PowerShell 5.1+ or PowerShell 7+; no Word installation needed.

.PARAMETER Folder
    Path to the folder containing the templates to compare. Required.

.PARAMETER OutputDir
    Where to write the snapshots and drift reports. Defaults to a subfolder
    named "_drift" inside Folder.

.PARAMETER Reference
    Optional file name (just the base name, no path) to use as the reference
    template. When set, drift is reported relative to this template. When
    omitted, drift is reported as "any value differs across the set".

.PARAMETER IncludeIdentical
    Include identical-everywhere properties in the drift CSV. Default is to
    suppress them so the report focuses on drift only.

.EXAMPLE
    .\Compare-WordTemplates.ps1 -Folder "\\blue\E\00-clients\nist-HM\nist-HM\HM-Template-250903"

.EXAMPLE
    .\Compare-WordTemplates.ps1 -Folder "C:\templates" -Reference "master.dotx"

.NOTES
    OOXML reference: ECMA-376 / ISO 29500. Namespaces used:
      w  = http://schemas.openxmlformats.org/wordprocessingml/2006/main
      r  = http://schemas.openxmlformats.org/officeDocument/2006/relationships
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Folder,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [Parameter(Mandatory = $false)]
    [string]$Reference,

    [switch]$IncludeIdentical
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
    throw "Folder not found: $Folder"
}

# Resolve to a raw filesystem path so file enumeration and the Extension property
# behave correctly when CWD is a UNC path (which PowerShell prefixes with the
# "Microsoft.PowerShell.Core\FileSystem::" provider name).
$Folder = (Resolve-Path -LiteralPath $Folder).ProviderPath

if (-not $OutputDir) {
    $OutputDir = Join-Path $Folder '_drift'
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$files = [System.IO.DirectoryInfo]::new($Folder).GetFiles() |
    Where-Object { $_.Extension -match '^\.(docx|dotx|docm|dotm)$' } |
    Sort-Object Name

if ($files.Count -lt 2) {
    throw "Need at least 2 .docx/.dotx/.docm/.dotm files in $Folder. Found: $($files.Count)."
}

Write-Host "Found $($files.Count) file(s) to compare:" -ForegroundColor Cyan
$files | ForEach-Object { Write-Host "  $($_.Name)" }

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

# OOXML namespaces
$script:wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$script:rNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Trunc {
    param([string]$Text, [int]$Max = 200)
    if ($null -eq $Text) { return '' }
    $clean = $Text -replace "`r`n", ' ' -replace "`r", ' ' -replace "`n", ' '
    $clean = $clean.Trim()
    if ($clean.Length -le $Max) { return $clean }
    return $clean.Substring(0, $Max) + '...'
}

function Get-PackagedXml {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        $text = $reader.ReadToEnd()
        $reader.Dispose()
    } finally {
        $stream.Dispose()
    }
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $false
    $xml.LoadXml($text)
    return $xml
}

function New-NsManager {
    param([System.Xml.XmlDocument]$Xml)
    $nsm = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    $nsm.AddNamespace('w', $script:wNs)
    $nsm.AddNamespace('r', $script:rNs)
    # XmlNamespaceManager implements IEnumerable, so PowerShell will unroll
    # it on return unless we wrap in a single-element array via the comma
    # operator. Without this, callers receive an array of namespace prefix
    # strings instead of the manager and SelectNodes overload resolution fails.
    return ,$nsm
}

# Read a w:val attribute (which can be in the w namespace OR unqualified
# depending on the producer)
function W-Val {
    param([System.Xml.XmlElement]$Node)
    if (-not $Node) { return $null }
    $v = $Node.GetAttribute('val', $script:wNs)
    if ($v) { return $v }
    return $Node.GetAttribute('val')
}

# OOXML toggle property: presence with no val (or val=1/true) means true,
# val=0/false means false, absence means inherit (returned as $null)
function Read-Toggle {
    param([System.Xml.XmlElement]$Parent, [string]$LocalName, [System.Xml.XmlNamespaceManager]$Nsm)
    $node = $Parent.SelectSingleNode("w:$LocalName", $Nsm)
    if (-not $node) { return $null }
    $v = W-Val $node
    if ($null -eq $v -or $v -eq '') { return $true }
    if ($v -eq '1' -or $v -eq 'true' -or $v -eq 'on') { return $true }
    return $false
}

# Parse run properties (rPr) into a flat hash
function Read-RunProps {
    param([System.Xml.XmlElement]$Rpr, [System.Xml.XmlNamespaceManager]$Nsm)
    if (-not $Rpr) { return $null }
    $r = [ordered]@{}

    $rFonts = $Rpr.SelectSingleNode('w:rFonts', $Nsm)
    if ($rFonts) {
        foreach ($attr in @('ascii','hAnsi','cs','eastAsia','asciiTheme','hAnsiTheme','csTheme','eastAsiaTheme')) {
            $v = $rFonts.GetAttribute($attr, $script:wNs)
            if ($v) { $r["Font.$attr"] = $v }
        }
    }
    $sz = $Rpr.SelectSingleNode('w:sz', $Nsm); if ($sz) { $r['SizeHalfPt'] = W-Val $sz }
    $szCs = $Rpr.SelectSingleNode('w:szCs', $Nsm); if ($szCs) { $r['SizeCsHalfPt'] = W-Val $szCs }
    foreach ($t in @('b','bCs','i','iCs','caps','smallCaps','strike','dstrike','vanish','outline','shadow','emboss','imprint','noProof','snapToGrid','spacing','kern')) {
        $val = Read-Toggle -Parent $Rpr -LocalName $t -Nsm $Nsm
        if ($null -ne $val) { $r["Toggle.$t"] = $val }
    }
    $u = $Rpr.SelectSingleNode('w:u', $Nsm)
    if ($u) {
        $r['Underline'] = W-Val $u
        $uc = $u.GetAttribute('color', $script:wNs); if ($uc) { $r['UnderlineColor'] = $uc }
    }
    $color = $Rpr.SelectSingleNode('w:color', $Nsm)
    if ($color) {
        $cv = W-Val $color; if ($cv) { $r['Color'] = $cv }
        $tc = $color.GetAttribute('themeColor', $script:wNs); if ($tc) { $r['ColorTheme'] = $tc }
        $tt = $color.GetAttribute('themeTint', $script:wNs); if ($tt) { $r['ColorThemeTint'] = $tt }
        $ts = $color.GetAttribute('themeShade', $script:wNs); if ($ts) { $r['ColorThemeShade'] = $ts }
    }
    $shd = $Rpr.SelectSingleNode('w:shd', $Nsm)
    if ($shd) {
        $r['ShadingFill'] = $shd.GetAttribute('fill', $script:wNs)
        $r['ShadingPattern'] = $shd.GetAttribute('val', $script:wNs)
    }
    $vert = $Rpr.SelectSingleNode('w:vertAlign', $Nsm); if ($vert) { $r['VertAlign'] = W-Val $vert }
    $lang = $Rpr.SelectSingleNode('w:lang', $Nsm)
    if ($lang) {
        $lv = $lang.GetAttribute('val', $script:wNs); if ($lv) { $r['LangVal'] = $lv }
        $le = $lang.GetAttribute('eastAsia', $script:wNs); if ($le) { $r['LangEastAsia'] = $le }
        $lc = $lang.GetAttribute('bidi', $script:wNs); if ($lc) { $r['LangBidi'] = $lc }
    }

    if ($r.Count -eq 0) { return $null }
    return $r
}

# Parse paragraph properties (pPr) into a flat hash
function Read-ParaProps {
    param([System.Xml.XmlElement]$Ppr, [System.Xml.XmlNamespaceManager]$Nsm)
    if (-not $Ppr) { return $null }
    $p = [ordered]@{}

    $jc = $Ppr.SelectSingleNode('w:jc', $Nsm); if ($jc) { $p['Alignment'] = W-Val $jc }

    $spacing = $Ppr.SelectSingleNode('w:spacing', $Nsm)
    if ($spacing) {
        foreach ($attr in @('before','after','line','lineRule','beforeLines','afterLines','beforeAutospacing','afterAutospacing')) {
            $v = $spacing.GetAttribute($attr, $script:wNs); if ($v) { $p["Spacing.$attr"] = $v }
        }
    }

    $ind = $Ppr.SelectSingleNode('w:ind', $Nsm)
    if ($ind) {
        foreach ($attr in @('left','right','firstLine','hanging','start','end')) {
            $v = $ind.GetAttribute($attr, $script:wNs); if ($v) { $p["Ind.$attr"] = $v }
        }
    }

    foreach ($t in @('keepNext','keepLines','pageBreakBefore','widowControl','suppressLineNumbers','suppressAutoHyphens')) {
        $val = Read-Toggle -Parent $Ppr -LocalName $t -Nsm $Nsm
        if ($null -ne $val) { $p["Toggle.$t"] = $val }
    }

    $outline = $Ppr.SelectSingleNode('w:outlineLvl', $Nsm); if ($outline) { $p['OutlineLevel'] = W-Val $outline }
    $numPr = $Ppr.SelectSingleNode('w:numPr', $Nsm)
    if ($numPr) {
        $ilvl = $numPr.SelectSingleNode('w:ilvl', $Nsm); if ($ilvl) { $p['Num.Ilvl'] = W-Val $ilvl }
        $numId = $numPr.SelectSingleNode('w:numId', $Nsm); if ($numId) { $p['Num.NumId'] = W-Val $numId }
    }

    $pBdr = $Ppr.SelectSingleNode('w:pBdr', $Nsm)
    if ($pBdr) {
        foreach ($side in @('top','bottom','left','right','between','bar')) {
            $b = $pBdr.SelectSingleNode("w:$side", $Nsm)
            if ($b) {
                foreach ($attr in @('val','sz','space','color','themeColor')) {
                    $v = $b.GetAttribute($attr, $script:wNs); if ($v) { $p["Border.$side.$attr"] = $v }
                }
            }
        }
    }

    $shd = $Ppr.SelectSingleNode('w:shd', $Nsm)
    if ($shd) {
        $p['Shading.fill'] = $shd.GetAttribute('fill', $script:wNs)
        $p['Shading.val']  = $shd.GetAttribute('val', $script:wNs)
    }

    if ($p.Count -eq 0) { return $null }
    return $p
}

function Read-TableProps {
    param([System.Xml.XmlElement]$TblPr, [System.Xml.XmlNamespaceManager]$Nsm)
    if (-not $TblPr) { return $null }
    $t = [ordered]@{}

    $tblW = $TblPr.SelectSingleNode('w:tblW', $Nsm)
    if ($tblW) {
        $t['Width.w'] = $tblW.GetAttribute('w', $script:wNs)
        $t['Width.type'] = $tblW.GetAttribute('type', $script:wNs)
    }

    $tblBorders = $TblPr.SelectSingleNode('w:tblBorders', $Nsm)
    if ($tblBorders) {
        foreach ($side in @('top','bottom','left','right','insideH','insideV')) {
            $b = $tblBorders.SelectSingleNode("w:$side", $Nsm)
            if ($b) {
                foreach ($attr in @('val','sz','space','color','themeColor')) {
                    $v = $b.GetAttribute($attr, $script:wNs); if ($v) { $t["Border.$side.$attr"] = $v }
                }
            }
        }
    }

    $tblCellMar = $TblPr.SelectSingleNode('w:tblCellMar', $Nsm)
    if ($tblCellMar) {
        foreach ($side in @('top','bottom','left','right','start','end')) {
            $b = $tblCellMar.SelectSingleNode("w:$side", $Nsm)
            if ($b) {
                $v = $b.GetAttribute('w', $script:wNs); if ($v) { $t["CellMar.$side.w"] = $v }
                $tp = $b.GetAttribute('type', $script:wNs); if ($tp) { $t["CellMar.$side.type"] = $tp }
            }
        }
    }

    $tblStyle = $TblPr.SelectSingleNode('w:tblStyle', $Nsm); if ($tblStyle) { $t['Style'] = W-Val $tblStyle }
    $tblLayout = $TblPr.SelectSingleNode('w:tblLayout', $Nsm); if ($tblLayout) { $t['Layout'] = $tblLayout.GetAttribute('type', $script:wNs) }
    $tblLook = $TblPr.SelectSingleNode('w:tblLook', $Nsm)
    if ($tblLook) {
        foreach ($attr in @('val','firstRow','lastRow','firstColumn','lastColumn','noHBand','noVBand')) {
            $v = $tblLook.GetAttribute($attr, $script:wNs); if ($v) { $t["Look.$attr"] = $v }
        }
    }

    if ($t.Count -eq 0) { return $null }
    return $t
}

function Get-StylesFromZip {
    param([System.IO.Compression.ZipArchive]$Zip)
    $xml = Get-PackagedXml -Zip $Zip -EntryName 'word/styles.xml'
    if (-not $xml) { return @() }
    $nsm = New-NsManager -Xml $xml

    $styles = @()
    foreach ($s in $xml.SelectNodes('//w:style', $nsm)) {
        $type = $s.GetAttribute('type', $script:wNs)
        $styleId = $s.GetAttribute('styleId', $script:wNs)
        $nameNode = $s.SelectSingleNode('w:name', $nsm)
        $aliasNode = $s.SelectSingleNode('w:aliases', $nsm)
        $basedOn = $s.SelectSingleNode('w:basedOn', $nsm)
        $next = $s.SelectSingleNode('w:next', $nsm)
        $link = $s.SelectSingleNode('w:link', $nsm)

        $entry = [ordered]@{
            StyleId     = $styleId
            Name        = if ($nameNode) { W-Val $nameNode } else { $styleId }
            Aliases     = if ($aliasNode) { W-Val $aliasNode } else { $null }
            Type        = $type
            BasedOn     = if ($basedOn) { W-Val $basedOn } else { $null }
            Next        = if ($next)    { W-Val $next }    else { $null }
            Link        = if ($link)    { W-Val $link }    else { $null }
            CustomStyle = ($s.GetAttribute('customStyle', $script:wNs) -eq '1')
            Default     = ($s.GetAttribute('default', $script:wNs) -eq '1')
        }

        $rPr = $s.SelectSingleNode('w:rPr', $nsm)
        $pPr = $s.SelectSingleNode('w:pPr', $nsm)
        $tblPr = $s.SelectSingleNode('w:tblPr', $nsm)
        $rp = Read-RunProps  -Rpr $rPr -Nsm $nsm; if ($rp) { $entry.RunProps = $rp }
        $pp = Read-ParaProps -Ppr $pPr -Nsm $nsm; if ($pp) { $entry.ParaProps = $pp }
        $tp = Read-TableProps -TblPr $tblPr -Nsm $nsm; if ($tp) { $entry.TableProps = $tp }

        # Conditional table style overrides (header row, first column, etc.)
        $condBlocks = $s.SelectNodes('w:tblStylePr', $nsm)
        if ($condBlocks.Count -gt 0) {
            $entry.TableConditional = [ordered]@{}
            foreach ($cb in $condBlocks) {
                $cType = $cb.GetAttribute('type', $script:wNs)
                $cEntry = [ordered]@{}
                $cRPr = $cb.SelectSingleNode('w:rPr', $nsm); if ($cRPr) { $cEntry.RunProps = Read-RunProps -Rpr $cRPr -Nsm $nsm }
                $cPPr = $cb.SelectSingleNode('w:pPr', $nsm); if ($cPPr) { $cEntry.ParaProps = Read-ParaProps -Ppr $cPPr -Nsm $nsm }
                $cTblPr = $cb.SelectSingleNode('w:tblPr', $nsm); if ($cTblPr) { $cEntry.TableProps = Read-TableProps -TblPr $cTblPr -Nsm $nsm }
                $entry.TableConditional[$cType] = $cEntry
            }
        }

        $styles += $entry
    }
    return $styles
}

# Helper: extract plain text from a Word XML node (concatenates w:t descendants)
function Get-XmlText {
    param([System.Xml.XmlNode]$Node, [System.Xml.XmlNamespaceManager]$Nsm)
    if (-not $Node) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $Node.SelectNodes('.//w:t', $Nsm)) {
        [void]$sb.Append($t.InnerText)
    }
    foreach ($br in $Node.SelectNodes('.//w:br', $Nsm)) {
        [void]$sb.Append(' ')
    }
    foreach ($tab in $Node.SelectNodes('.//w:tab', $Nsm)) {
        [void]$sb.Append("`t")
    }
    return $sb.ToString()
}

function Get-HeadersFootersFromZip {
    param([System.IO.Compression.ZipArchive]$Zip)
    $result = @()

    # Walk every word/headerN.xml and footerN.xml file; we don't try to resolve
    # which section they belong to (would require parsing rels and document.xml
    # sectPr) since we just want the content for drift comparison.
    $entries = $Zip.Entries | Where-Object { $_.FullName -match '^word/(header|footer)\d+\.xml$' } | Sort-Object FullName
    foreach ($e in $entries) {
        $xml = Get-PackagedXml -Zip $Zip -EntryName $e.FullName
        if (-not $xml) { continue }
        $nsm = New-NsManager -Xml $xml
        $kind = if ($e.FullName -match '^word/header') { 'header' } else { 'footer' }
        $text = Get-XmlText -Node $xml.DocumentElement -Nsm $nsm

        # Style references applied to paragraphs inside
        $pStyles = @()
        foreach ($pStyle in $xml.SelectNodes('//w:pStyle', $nsm)) {
            $v = W-Val $pStyle
            if ($v -and ($pStyles -notcontains $v)) { $pStyles += $v }
        }

        $result += [ordered]@{
            Part         = $e.FullName
            Kind         = $kind
            Text         = Trunc $text 800
            ParaStyles   = $pStyles
            FieldCount   = $xml.SelectNodes('//w:fldSimple', $nsm).Count + $xml.SelectNodes('//w:instrText', $nsm).Count
        }
    }
    return $result
}

function Get-BuildingBlocksFromZip {
    param([System.IO.Compression.ZipArchive]$Zip)
    $blocks = @()
    $xml = Get-PackagedXml -Zip $Zip -EntryName 'word/glossary/document.xml'
    if (-not $xml) { return $blocks }
    $nsm = New-NsManager -Xml $xml

    foreach ($dp in $xml.SelectNodes('//w:docPart', $nsm)) {
        $pr = $dp.SelectSingleNode('w:docPartPr', $nsm)
        if (-not $pr) { continue }
        $name = $pr.SelectSingleNode('w:name', $nsm)
        $cat = $pr.SelectSingleNode('w:category/w:name', $nsm)
        $catGallery = $pr.SelectSingleNode('w:category/w:gallery', $nsm)
        $type = $pr.SelectSingleNode('w:types/w:type', $nsm)
        $description = $pr.SelectSingleNode('w:description', $nsm)
        $behaviors = $pr.SelectNodes('w:behaviors/w:behavior', $nsm) | ForEach-Object { W-Val $_ }
        $contentText = Get-XmlText -Node ($dp.SelectSingleNode('w:docPartBody', $nsm)) -Nsm $nsm

        $blocks += [ordered]@{
            Name        = if ($name) { W-Val $name } else { $null }
            Category    = if ($cat) { W-Val $cat } else { $null }
            Gallery     = if ($catGallery) { W-Val $catGallery } else { $null }
            Type        = if ($type) { W-Val $type } else { $null }
            Description = if ($description) { Trunc (W-Val $description) 200 } else { $null }
            Behaviors   = ($behaviors -join ',')
            ContentSnippet = Trunc $contentText 200
        }
    }
    return $blocks
}

function Get-ContentControlsFromZip {
    param([System.IO.Compression.ZipArchive]$Zip)
    $controls = @()

    # Content controls live in document.xml and may also live in headers/footers
    # and the glossary. We scan all of them.
    $entries = $Zip.Entries | Where-Object {
        $_.FullName -eq 'word/document.xml' -or
        $_.FullName -match '^word/(header|footer)\d+\.xml$' -or
        $_.FullName -eq 'word/glossary/document.xml'
    }

    foreach ($e in $entries) {
        $xml = Get-PackagedXml -Zip $Zip -EntryName $e.FullName
        if (-not $xml) { continue }
        $nsm = New-NsManager -Xml $xml

        foreach ($sdt in $xml.SelectNodes('//w:sdt', $nsm)) {
            $pr = $sdt.SelectSingleNode('w:sdtPr', $nsm)
            if (-not $pr) { continue }
            $titleNode = $pr.SelectSingleNode('w:alias', $nsm)
            $tagNode = $pr.SelectSingleNode('w:tag', $nsm)
            $idNode = $pr.SelectSingleNode('w:id', $nsm)

            # Determine type by checking which child marker is present
            $typeMarkers = @('text','richText','picture','comboBox','dropDownList','date','docPartObj','docPartList','citation','group','bibliography','equation','repeatingSection','repeatingSectionItem','checkbox','dataBinding')
            $type = $null
            foreach ($m in $typeMarkers) {
                if ($pr.SelectSingleNode("w:$m", $nsm)) { $type = $m; break }
            }

            $lockNode = $pr.SelectSingleNode('w:lock', $nsm)
            $lockVal = if ($lockNode) { W-Val $lockNode } else { $null }

            $contentNode = $sdt.SelectSingleNode('w:sdtContent', $nsm)
            $contentText = Get-XmlText -Node $contentNode -Nsm $nsm

            $placeholderNode = $pr.SelectSingleNode('w:placeholder/w:docPart', $nsm)

            $controls += [ordered]@{
                Part        = $e.FullName
                Title       = if ($titleNode) { W-Val $titleNode } else { $null }
                Tag         = if ($tagNode) { W-Val $tagNode } else { $null }
                Id          = if ($idNode) { W-Val $idNode } else { $null }
                Type        = $type
                Lock        = $lockVal
                Placeholder = if ($placeholderNode) { W-Val $placeholderNode } else { $null }
                Content     = Trunc $contentText 200
            }
        }
    }
    return $controls
}

function Get-FieldsFromZip {
    param([System.IO.Compression.ZipArchive]$Zip)
    $fields = @()
    $entries = $Zip.Entries | Where-Object {
        $_.FullName -eq 'word/document.xml' -or
        $_.FullName -match '^word/(header|footer)\d+\.xml$' -or
        $_.FullName -eq 'word/glossary/document.xml'
    }
    foreach ($e in $entries) {
        $xml = Get-PackagedXml -Zip $Zip -EntryName $e.FullName
        if (-not $xml) { continue }
        $nsm = New-NsManager -Xml $xml

        # Simple fields
        foreach ($f in $xml.SelectNodes('//w:fldSimple', $nsm)) {
            $code = $f.GetAttribute('instr', $script:wNs)
            $fields += [ordered]@{
                Part = $e.FullName
                Kind = 'simple'
                Code = Trunc $code 300
            }
        }

        # Complex fields: instrText spans between fldChar begin/end markers
        $instrs = $xml.SelectNodes('//w:instrText', $nsm)
        foreach ($it in $instrs) {
            $text = $it.InnerText
            if ($text -and $text.Trim()) {
                $fields += [ordered]@{
                    Part = $e.FullName
                    Kind = 'complex'
                    Code = Trunc $text 300
                }
            }
        }
    }
    return $fields
}

function Get-BookmarksFromZip {
    param([System.IO.Compression.ZipArchive]$Zip)
    $bookmarks = @()
    $entries = $Zip.Entries | Where-Object {
        $_.FullName -eq 'word/document.xml' -or
        $_.FullName -match '^word/(header|footer)\d+\.xml$' -or
        $_.FullName -eq 'word/glossary/document.xml'
    }
    foreach ($e in $entries) {
        $xml = Get-PackagedXml -Zip $Zip -EntryName $e.FullName
        if (-not $xml) { continue }
        $nsm = New-NsManager -Xml $xml
        foreach ($b in $xml.SelectNodes('//w:bookmarkStart', $nsm)) {
            $name = $b.GetAttribute('name', $script:wNs)
            if (-not $name) { continue }
            # Skip Word's internal _Toc, _Hlk, _GoBack etc bookmarks
            if ($name -match '^_(Toc|Hlk|GoBack|Ref)') { continue }
            $bookmarks += [ordered]@{
                Part = $e.FullName
                Name = $name
            }
        }
    }
    return $bookmarks
}

function Get-CustomXmlPartsFromZip {
    param([System.IO.Compression.ZipArchive]$Zip)
    $parts = @()
    $items = $Zip.Entries | Where-Object { $_.FullName -match '^customXml/item\d+\.xml$' } | Sort-Object FullName
    foreach ($it in $items) {
        $xml = Get-PackagedXml -Zip $Zip -EntryName $it.FullName
        if (-not $xml) { continue }
        # Read paired itemPropsN.xml for namespace metadata
        $propsName = $it.FullName -replace '/item(\d+)\.xml$', '/itemProps$1.xml'
        $propsXml = Get-PackagedXml -Zip $Zip -EntryName $propsName

        $namespace = $xml.DocumentElement.NamespaceURI
        $rootName = $xml.DocumentElement.LocalName
        $preview = Trunc $xml.DocumentElement.OuterXml 500

        $parts += [ordered]@{
            Part         = $it.FullName
            RootElement  = $rootName
            NamespaceURI = $namespace
            Preview      = $preview
        }
    }
    return $parts
}

function Get-DocumentSnapshotFromZip {
    param([string]$FilePath, [string]$FileName)

    Write-Host "  Opening package..." -ForegroundColor DarkGray
    $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            Write-Host "  Reading styles..." -ForegroundColor DarkGray
            $styles = Get-StylesFromZip -Zip $zip
            Write-Host "  Reading headers and footers..." -ForegroundColor DarkGray
            $hf = Get-HeadersFootersFromZip -Zip $zip
            Write-Host "  Reading building blocks (glossary)..." -ForegroundColor DarkGray
            $bb = Get-BuildingBlocksFromZip -Zip $zip
            Write-Host "  Reading content controls..." -ForegroundColor DarkGray
            $cc = Get-ContentControlsFromZip -Zip $zip
            Write-Host "  Reading fields..." -ForegroundColor DarkGray
            $fields = Get-FieldsFromZip -Zip $zip
            Write-Host "  Reading bookmarks..." -ForegroundColor DarkGray
            $bm = Get-BookmarksFromZip -Zip $zip
            Write-Host "  Reading custom XML parts..." -ForegroundColor DarkGray
            $cxp = Get-CustomXmlPartsFromZip -Zip $zip

            return [ordered]@{
                FileName        = $FileName
                ExtractedAt     = (Get-Date).ToString('o')
                Styles          = $styles
                HeadersFooters  = $hf
                BuildingBlocks  = $bb
                ContentControls = $cc
                Fields          = $fields
                Bookmarks       = $bm
                CustomXmlParts  = $cxp
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

# ---------------------------------------------------------------------------
# Snapshot every file
# ---------------------------------------------------------------------------

$snapshots = [ordered]@{}
foreach ($file in $files) {
    Write-Host ""
    Write-Host "Processing $($file.Name)" -ForegroundColor Yellow
    $snap = Get-DocumentSnapshotFromZip -FilePath $file.FullName -FileName $file.Name
    $snapshots[$file.Name] = $snap

    $jsonPath = Join-Path $OutputDir ([System.IO.Path]::GetFileNameWithoutExtension($file.Name) + '.snapshot.json')
    $snap | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    Write-Host "  Wrote $jsonPath" -ForegroundColor DarkGreen
}

# ---------------------------------------------------------------------------
# Drift detection
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "Computing drift..." -ForegroundColor Cyan

function Flatten-Snapshot {
    param($Snap)
    $flat = @{}

    foreach ($s in $Snap.Styles) {
        $key = "style[$($s.Type)]::$($s.Name)"
        foreach ($prop in @('StyleId','BasedOn','Next','Link','CustomStyle','Default','Aliases')) {
            if ($s.Contains($prop) -and $null -ne $s[$prop]) { $flat["$key|$prop"] = $s[$prop] }
        }
        foreach ($bag in @('RunProps','ParaProps','TableProps')) {
            if ($s.Contains($bag) -and $s[$bag]) {
                foreach ($p in $s[$bag].Keys) { $flat["$key|$bag.$p"] = $s[$bag][$p] }
            }
        }
        if ($s.Contains('TableConditional') -and $s.TableConditional) {
            foreach ($cType in $s.TableConditional.Keys) {
                $cEntry = $s.TableConditional[$cType]
                foreach ($bag in @('RunProps','ParaProps','TableProps')) {
                    if ($cEntry.Contains($bag) -and $cEntry[$bag]) {
                        foreach ($p in $cEntry[$bag].Keys) { $flat["$key|TableCond.$cType.$bag.$p"] = $cEntry[$bag][$p] }
                    }
                }
            }
        }
    }

    foreach ($hf in $Snap.HeadersFooters) {
        $key = "$($hf.Kind)::$($hf.Part)"
        $flat["$key|Text"] = $hf.Text
        $flat["$key|FieldCount"] = $hf.FieldCount
        $flat["$key|ParaStyles"] = ($hf.ParaStyles -join ',')
    }

    foreach ($bb in $Snap.BuildingBlocks) {
        $bkey = "buildingBlock::$($bb.Gallery)::$($bb.Category)::$($bb.Name)"
        foreach ($p in @('Type','Description','Behaviors','ContentSnippet')) {
            if ($bb.Contains($p)) { $flat["$bkey|$p"] = $bb[$p] }
        }
    }

    foreach ($cc in $Snap.ContentControls) {
        $id = if ($cc.Tag) { "tag:$($cc.Tag)" } elseif ($cc.Title) { "title:$($cc.Title)" } else { "id:$($cc.Id)" }
        $key = "contentControl[$($cc.Part)]::$id"
        foreach ($p in @('Title','Type','Lock','Placeholder','Content')) {
            if ($cc.Contains($p)) { $flat["$key|$p"] = $cc[$p] }
        }
    }

    $fOrd = @{}
    foreach ($f in $Snap.Fields) {
        $code = $f.Code
        $okey = "$($f.Part)::$code"
        if (-not $fOrd.ContainsKey($okey)) { $fOrd[$okey] = 0 }
        $fOrd[$okey]++
        $flat["field[$($f.Part)]::$code#$($fOrd[$okey])|Kind"] = $f.Kind
    }

    foreach ($b in $Snap.Bookmarks) {
        $flat["bookmark[$($b.Part)]::$($b.Name)|Exists"] = $true
    }

    foreach ($cxp in $Snap.CustomXmlParts) {
        $key = "customXml::$($cxp.NamespaceURI)::$($cxp.RootElement)"
        $flat["$key|Preview"] = $cxp.Preview
    }

    return $flat
}

$flat = @{}
foreach ($fn in $snapshots.Keys) {
    $flat[$fn] = Flatten-Snapshot $snapshots[$fn]
}

$fileNames = @($snapshots.Keys)
$allKeys = New-Object System.Collections.Generic.HashSet[string]
foreach ($fn in $fileNames) {
    foreach ($k in $flat[$fn].Keys) { [void]$allKeys.Add($k) }
}

$rows = @()
foreach ($k in ($allKeys | Sort-Object)) {
    $parts = $k -split '\|', 2
    $itemKey = $parts[0]
    $propPath = if ($parts.Length -gt 1) { $parts[1] } else { '<root>' }

    $values = @{}
    foreach ($fn in $fileNames) {
        if ($flat[$fn].ContainsKey($k)) {
            $values[$fn] = $flat[$fn][$k]
        } else {
            $values[$fn] = '<MISSING>'
        }
    }

    if ($Reference) {
        if (-not $values.ContainsKey($Reference)) {
            throw "Reference file '$Reference' not found among snapshots: $($fileNames -join ', ')"
        }
        $refVal = [string]$values[$Reference]
        $drift = $false
        foreach ($fn in $fileNames) {
            if ($fn -eq $Reference) { continue }
            if (([string]$values[$fn]) -ne $refVal) { $drift = $true; break }
        }
    } else {
        $distinctVals = ($values.Values | ForEach-Object { [string]$_ } | Sort-Object -Unique)
        $drift = ($distinctVals.Count -gt 1)
    }

    if (-not $drift -and -not $IncludeIdentical) { continue }

    $row = [ordered]@{
        Item     = $itemKey
        Property = $propPath
        Drift    = if ($drift) { 'YES' } else { 'no' }
    }
    foreach ($fn in $fileNames) { $row[$fn] = $values[$fn] }
    $rows += [pscustomobject]$row
}

# ---------------------------------------------------------------------------
# CSV output
# ---------------------------------------------------------------------------

$csvPath = Join-Path $OutputDir 'drift.csv'
$rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "Drift CSV: $csvPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# HTML output
# ---------------------------------------------------------------------------

$htmlPath = Join-Path $OutputDir 'drift.html'
$summary = @{
    TotalRows  = $rows.Count
    ByCategory = @{}
}
foreach ($row in $rows) {
    $cat = ($row.Item -split '::', 2)[0] -replace '\[.*\]', ''
    if (-not $summary.ByCategory.ContainsKey($cat)) { $summary.ByCategory[$cat] = 0 }
    $summary.ByCategory[$cat]++
}

$html = New-Object System.Text.StringBuilder
[void]$html.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8'><title>Word Template Drift Report</title>")
[void]$html.AppendLine(@"
<style>
  body { font-family: -apple-system, Segoe UI, Roboto, sans-serif; padding: 20px; color: #222; }
  h1 { font-size: 22px; margin-bottom: 4px; }
  .meta { color: #666; margin-bottom: 18px; font-size: 13px; }
  table { border-collapse: collapse; width: 100%; font-size: 13px; }
  th, td { border: 1px solid #d4d4d4; padding: 6px 8px; vertical-align: top; text-align: left; }
  th { background: #f0f0f0; position: sticky; top: 0; }
  tr.drift td.val { font-weight: 500; }
  td.val.differ { background: #fff3cd; }
  td.val.missing { background: #ffd7d7; color: #a00; font-style: italic; }
  td.item { font-family: ui-monospace, Consolas, monospace; color: #444; }
  td.prop { font-family: ui-monospace, Consolas, monospace; color: #555; }
  .summary { background: #f8f8f8; padding: 12px; border-left: 4px solid #4a90e2; margin-bottom: 16px; }
  .summary h2 { margin: 0 0 6px 0; font-size: 16px; }
  .summary ul { margin: 4px 0 0 18px; padding: 0; }
</style>
"@)
[void]$html.AppendLine("</head><body>")
[void]$html.AppendLine("<h1>Word Template Drift Report</h1>")
[void]$html.AppendLine("<div class='meta'>Generated: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) | Folder: $Folder | Files: $($fileNames.Count)</div>")

[void]$html.AppendLine("<div class='summary'><h2>Summary</h2>")
[void]$html.AppendLine("<div>Total drifting properties: <strong>$($summary.TotalRows)</strong></div>")
[void]$html.AppendLine("<ul>")
foreach ($cat in ($summary.ByCategory.Keys | Sort-Object)) {
    [void]$html.AppendLine("<li>$cat`: $($summary.ByCategory[$cat])</li>")
}
[void]$html.AppendLine("</ul></div>")

[void]$html.AppendLine("<table><thead><tr><th>Item</th><th>Property</th>")
foreach ($fn in $fileNames) { [void]$html.AppendLine("<th>$([System.Web.HttpUtility]::HtmlEncode($fn))</th>") }
[void]$html.AppendLine("</tr></thead><tbody>")

foreach ($row in $rows) {
    $rowClass = if ($row.Drift -eq 'YES') { 'drift' } else { '' }
    [void]$html.AppendLine("<tr class='$rowClass'>")
    [void]$html.AppendLine("<td class='item'>$([System.Web.HttpUtility]::HtmlEncode($row.Item))</td>")
    [void]$html.AppendLine("<td class='prop'>$([System.Web.HttpUtility]::HtmlEncode($row.Property))</td>")

    $vals = @{}
    foreach ($fn in $fileNames) { $vals[$fn] = [string]$row.$fn }
    if ($Reference) {
        $ref = $vals[$Reference]
        foreach ($fn in $fileNames) {
            $v = $vals[$fn]
            $cls = 'val'
            if ($v -eq '<MISSING>') { $cls += ' missing' }
            elseif ($fn -ne $Reference -and $v -ne $ref) { $cls += ' differ' }
            [void]$html.AppendLine("<td class='$cls'>$([System.Web.HttpUtility]::HtmlEncode($v))</td>")
        }
    } else {
        $distinct = ($vals.Values | Sort-Object -Unique)
        $isDrift = $distinct.Count -gt 1
        foreach ($fn in $fileNames) {
            $v = $vals[$fn]
            $cls = 'val'
            if ($v -eq '<MISSING>') { $cls += ' missing' }
            elseif ($isDrift) { $cls += ' differ' }
            [void]$html.AppendLine("<td class='$cls'>$([System.Web.HttpUtility]::HtmlEncode($v))</td>")
        }
    }

    [void]$html.AppendLine("</tr>")
}

[void]$html.AppendLine("</tbody></table></body></html>")
$html.ToString() | Set-Content -LiteralPath $htmlPath -Encoding UTF8

Write-Host "Drift HTML: $htmlPath" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Drifting properties: $($summary.TotalRows)"
foreach ($cat in ($summary.ByCategory.Keys | Sort-Object)) {
    Write-Host "  $cat`: $($summary.ByCategory[$cat])"
}
Write-Host ""
Write-Host "Done." -ForegroundColor Green
