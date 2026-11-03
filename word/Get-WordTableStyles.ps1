<#
.SYNOPSIS
    Lists every table style in a single Word file with full detail,
    including every conditional formatting block (header row, banded rows,
    first column, etc.) as a Markdown document.

.DESCRIPTION
    Reads word/styles.xml and extracts only style records of type="table".
    For each table style emits whole-table defaults (font, paragraph,
    cell margins, table-level borders) plus every conditional formatting
    block with its overrides (font, paragraph, row props, cell props,
    shading, borders).

    Every relevant attribute is shown whether or not it is set. Unset
    attributes display as "(not set)" so it is clear what the file
    explicitly defines vs leaves to inheritance.

.PARAMETER File
    Path to the .docx, .dotx, .docm, or .dotm file. Required.

.PARAMETER OutputDir
    Where to write the output. Defaults to a sibling "_styles" folder.

.PARAMETER Html
    Also produce an HTML version for browser viewing.

.EXAMPLE
    .\Get-WordTableStyles.ps1 -File "master.dotx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [switch]$Html
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
$resolvedPath = (Resolve-Path -LiteralPath $File).ProviderPath
if ($resolvedPath -notmatch '\.(docx|dotx|docm|dotm)$') {
    throw "Unsupported extension. Path resolved to: $resolvedPath. Expected .docx, .dotx, .docm, or .dotm."
}
$dirName  = [System.IO.Path]::GetDirectoryName($resolvedPath)
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
$fileName = [System.IO.Path]::GetFileName($resolvedPath)

if (-not $OutputDir) { $OutputDir = Join-Path $dirName '_styles' }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# Conditional block names in the canonical order Word presents them
$ConditionLabels = [ordered]@{
    'wholeTable' = 'Whole Table (default)'
    'firstRow'   = 'First Row (header)'
    'lastRow'    = 'Last Row (total)'
    'firstCol'   = 'First Column'
    'lastCol'    = 'Last Column'
    'band1Horz'  = 'Banded Rows (odd)'
    'band2Horz'  = 'Banded Rows (even)'
    'band1Vert'  = 'Banded Columns (odd)'
    'band2Vert'  = 'Banded Columns (even)'
    'neCell'     = 'Top-Right Cell'
    'nwCell'     = 'Top-Left Cell'
    'seCell'     = 'Bottom-Right Cell'
    'swCell'     = 'Bottom-Left Cell'
}

$FontAttrs    = @('Font (ascii)','Font (hAnsi)','Font (theme)','Size','Bold','Italic','AllCaps','SmallCaps','Strikethrough','Underline','Color','ThemeColor','VerticalAlign')
$ParaAttrs    = @('Alignment','SpaceBefore','SpaceAfter','LineSpacing','IndentLeft','IndentRight','OutlineLevel','KeepWithNext','KeepLines')
$TableAttrs   = @('TableWidth','RowBandSize','ColBandSize','CellMargin.Top','CellMargin.Bottom','CellMargin.Left','CellMargin.Right','Border.Top','Border.Bottom','Border.Left','Border.Right','Border.InsideH','Border.InsideV','Layout','Shading')
$RowAttrs     = @('CantSplit','TblHeader','RowHeight')
$CellAttrs    = @('CellShading','CellBorder.Top','CellBorder.Bottom','CellBorder.Left','CellBorder.Right','VAlign','NoWrap','TcW')

# Helpers
function W-Val { param([System.Xml.XmlElement]$N)
    if (-not $N) { return $null }
    $v = $N.GetAttribute('val', $wNs); if ($v) { return $v }
    return $N.GetAttribute('val')
}
function Read-Toggle {
    param([System.Xml.XmlElement]$Parent, [string]$Local, [System.Xml.XmlNamespaceManager]$Nsm)
    $n = $Parent.SelectSingleNode("w:$Local", $Nsm); if (-not $n) { return $null }
    $v = W-Val $n
    if ($null -eq $v -or $v -eq '' -or $v -eq '1' -or $v -eq 'true' -or $v -eq 'on') { return 'true' }
    return 'false'
}
function Format-Twips {
    param([string]$Twips, [string]$Mode = 'points')
    if ($null -eq $Twips -or $Twips -eq '') { return $null }
    $n = 0; if (-not [int]::TryParse($Twips, [ref]$n)) { return "$Twips (raw)" }
    if ($Mode -eq 'inches') { $inch = [math]::Round($n / 1440.0, 3); return "$Twips twips ($inch in)" }
    $pt = [math]::Round($n / 20.0, 2); return "$Twips twips (${pt}pt)"
}
function Format-HalfPt {
    param([string]$Hp)
    if ($null -eq $Hp -or $Hp -eq '') { return $null }
    $n = 0; if (-not [int]::TryParse($Hp, [ref]$n)) { return "$Hp (raw)" }
    $pt = $n / 2.0; return "${pt}pt ($Hp half-pts)"
}
function Format-Border {
    param([System.Xml.XmlElement]$Border)
    if (-not $Border) { return $null }
    $val = $Border.GetAttribute('val', $wNs)
    $sz = $Border.GetAttribute('sz', $wNs)
    $color = $Border.GetAttribute('color', $wNs)
    $themeColor = $Border.GetAttribute('themeColor', $wNs)
    $space = $Border.GetAttribute('space', $wNs)
    $bits = @()
    if ($val) { $bits += "$val" }
    if ($sz) { $bits += "${sz}/8 pt" }
    if ($color -and $color -ne 'auto') { $bits += "#$color" }
    if ($themeColor) { $bits += "theme:$themeColor" }
    if ($space) { $bits += "space:$space" }
    return ($bits -join ', ')
}

function Read-RunProps {
    param([System.Xml.XmlElement]$RPr, [System.Xml.XmlNamespaceManager]$Nsm, [System.Collections.IDictionary]$V)
    if (-not $RPr) { return }
    $rFonts = $RPr.SelectSingleNode('w:rFonts', $Nsm)
    if ($rFonts) {
        $a = $rFonts.GetAttribute('ascii', $wNs); if ($a) { $V['Font (ascii)'] = $a }
        $h = $rFonts.GetAttribute('hAnsi', $wNs); if ($h) { $V['Font (hAnsi)'] = $h }
        $t = $rFonts.GetAttribute('asciiTheme', $wNs); if ($t) { $V['Font (theme)'] = $t }
    }
    $sz = $RPr.SelectSingleNode('w:sz', $Nsm); if ($sz) { $V['Size'] = Format-HalfPt (W-Val $sz) }
    foreach ($pair in @(@{X='b';K='Bold'},@{X='i';K='Italic'},@{X='caps';K='AllCaps'},@{X='smallCaps';K='SmallCaps'},@{X='strike';K='Strikethrough'})) {
        $t = Read-Toggle -Parent $RPr -Local $pair.X -Nsm $Nsm
        if ($null -ne $t) { $V[$pair.K] = $t }
    }
    $u = $RPr.SelectSingleNode('w:u', $Nsm); if ($u) { $V['Underline'] = W-Val $u }
    $color = $RPr.SelectSingleNode('w:color', $Nsm)
    if ($color) {
        $cv = W-Val $color; $tc = $color.GetAttribute('themeColor', $wNs)
        if ($cv -and $cv -ne 'auto') { $V['Color'] = "#$cv" }
        if ($tc) { $V['ThemeColor'] = $tc }
    }
    $vert = $RPr.SelectSingleNode('w:vertAlign', $Nsm); if ($vert) { $V['VerticalAlign'] = W-Val $vert }
}

function Read-ParaProps {
    param([System.Xml.XmlElement]$PPr, [System.Xml.XmlNamespaceManager]$Nsm, [System.Collections.IDictionary]$V)
    if (-not $PPr) { return }
    $jc = $PPr.SelectSingleNode('w:jc', $Nsm); if ($jc) { $V['Alignment'] = W-Val $jc }
    $spacing = $PPr.SelectSingleNode('w:spacing', $Nsm)
    if ($spacing) {
        $b = $spacing.GetAttribute('before', $wNs); if ($b) { $V['SpaceBefore'] = Format-Twips $b }
        $a = $spacing.GetAttribute('after', $wNs); if ($a) { $V['SpaceAfter'] = Format-Twips $a }
        $line = $spacing.GetAttribute('line', $wNs); $rule = $spacing.GetAttribute('lineRule', $wNs)
        if ($line) {
            if ($rule -eq 'auto') { $multi = [math]::Round([int]$line / 240.0, 2); $V['LineSpacing'] = "$line/240ths line ($multi x)" }
            else { $V['LineSpacing'] = "$line twips ($rule)" }
        }
    }
    $ind = $PPr.SelectSingleNode('w:ind', $Nsm)
    if ($ind) {
        $l = $ind.GetAttribute('left', $wNs); if ($l) { $V['IndentLeft'] = Format-Twips $l 'inches' }
        $r = $ind.GetAttribute('right', $wNs); if ($r) { $V['IndentRight'] = Format-Twips $r 'inches' }
    }
    $outline = $PPr.SelectSingleNode('w:outlineLvl', $Nsm); if ($outline) { $V['OutlineLevel'] = W-Val $outline }
    $kn = Read-Toggle -Parent $PPr -Local 'keepNext' -Nsm $Nsm; if ($null -ne $kn) { $V['KeepWithNext'] = $kn }
    $kl = Read-Toggle -Parent $PPr -Local 'keepLines' -Nsm $Nsm; if ($null -ne $kl) { $V['KeepLines'] = $kl }
}

function Read-TableProps {
    param([System.Xml.XmlElement]$TblPr, [System.Xml.XmlNamespaceManager]$Nsm, [System.Collections.IDictionary]$V)
    if (-not $TblPr) { return }
    $tblW = $TblPr.SelectSingleNode('w:tblW', $Nsm)
    if ($tblW) {
        $w = $tblW.GetAttribute('w', $wNs); $t = $tblW.GetAttribute('type', $wNs)
        if ($w) { $V['TableWidth'] = "$w ($t)" }
    }
    $rb = $TblPr.SelectSingleNode('w:tblStyleRowBandSize', $Nsm)
    if ($rb) { $V['RowBandSize'] = W-Val $rb }
    $cb = $TblPr.SelectSingleNode('w:tblStyleColBandSize', $Nsm)
    if ($cb) { $V['ColBandSize'] = W-Val $cb }
    $tblCellMar = $TblPr.SelectSingleNode('w:tblCellMar', $Nsm)
    if ($tblCellMar) {
        foreach ($side in @('top','bottom','left','right')) {
            $b = $tblCellMar.SelectSingleNode("w:$side", $Nsm)
            if ($b) {
                $w = $b.GetAttribute('w', $wNs); $t = $b.GetAttribute('type', $wNs)
                if ($w) {
                    $cap = (Get-Culture).TextInfo.ToTitleCase($side)
                    $V["CellMargin.$cap"] = "$w ($t)"
                }
            }
        }
    }
    $tblBorders = $TblPr.SelectSingleNode('w:tblBorders', $Nsm)
    if ($tblBorders) {
        foreach ($side in @('top','bottom','left','right','insideH','insideV')) {
            $b = $tblBorders.SelectSingleNode("w:$side", $Nsm)
            if ($b) {
                $cap = (Get-Culture).TextInfo.ToTitleCase($side)
                $V["Border.$cap"] = Format-Border -Border $b
            }
        }
    }
    $tblLayout = $TblPr.SelectSingleNode('w:tblLayout', $Nsm)
    if ($tblLayout) { $V['Layout'] = $tblLayout.GetAttribute('type', $wNs) }
    $shd = $TblPr.SelectSingleNode('w:shd', $Nsm)
    if ($shd) {
        $fill = $shd.GetAttribute('fill', $wNs)
        if ($fill -and $fill -ne 'auto') { $V['Shading'] = "#$fill" }
    }
}

function Read-RowProps {
    param([System.Xml.XmlElement]$TrPr, [System.Xml.XmlNamespaceManager]$Nsm, [System.Collections.IDictionary]$V)
    if (-not $TrPr) { return }
    $cantSplit = Read-Toggle -Parent $TrPr -Local 'cantSplit' -Nsm $Nsm; if ($null -ne $cantSplit) { $V['CantSplit'] = $cantSplit }
    $tblHeader = Read-Toggle -Parent $TrPr -Local 'tblHeader' -Nsm $Nsm; if ($null -ne $tblHeader) { $V['TblHeader'] = $tblHeader }
    $trHeight = $TrPr.SelectSingleNode('w:trHeight', $Nsm)
    if ($trHeight) {
        $h = $trHeight.GetAttribute('val', $wNs); $r = $trHeight.GetAttribute('hRule', $wNs)
        if ($h) { $V['RowHeight'] = "$h twips ($r)" }
    }
}

function Read-CellProps {
    param([System.Xml.XmlElement]$TcPr, [System.Xml.XmlNamespaceManager]$Nsm, [System.Collections.IDictionary]$V)
    if (-not $TcPr) { return }
    $shd = $TcPr.SelectSingleNode('w:shd', $Nsm)
    if ($shd) {
        $fill = $shd.GetAttribute('fill', $wNs)
        if ($fill -and $fill -ne 'auto') { $V['CellShading'] = "#$fill" }
    }
    $tcBorders = $TcPr.SelectSingleNode('w:tcBorders', $Nsm)
    if ($tcBorders) {
        foreach ($side in @('top','bottom','left','right')) {
            $b = $tcBorders.SelectSingleNode("w:$side", $Nsm)
            if ($b) {
                $cap = (Get-Culture).TextInfo.ToTitleCase($side)
                $V["CellBorder.$cap"] = Format-Border -Border $b
            }
        }
    }
    $vAlign = $TcPr.SelectSingleNode('w:vAlign', $Nsm); if ($vAlign) { $V['VAlign'] = W-Val $vAlign }
    $noWrap = Read-Toggle -Parent $TcPr -Local 'noWrap' -Nsm $Nsm; if ($null -ne $noWrap) { $V['NoWrap'] = $noWrap }
    $tcW = $TcPr.SelectSingleNode('w:tcW', $Nsm)
    if ($tcW) {
        $w = $tcW.GetAttribute('w', $wNs); $t = $tcW.GetAttribute('type', $wNs)
        if ($w) { $V['TcW'] = "$w ($t)" }
    }
}

# ---------------------------------------------------------------------------
# Open package, read styles.xml
# ---------------------------------------------------------------------------

Write-Host "Reading table styles from $fileName..." -ForegroundColor Cyan

$stream = [System.IO.File]::Open($resolvedPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $entry = $zip.GetEntry('word/styles.xml')
        if (-not $entry) { throw "word/styles.xml not found in package." }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { $stylesXmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
} finally { $stream.Dispose() }

$xml = New-Object System.Xml.XmlDocument
$xml.LoadXml($stylesXmlText)
$nsm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsm.AddNamespace('w', $wNs)

$tableStyles = @()
foreach ($s in $xml.SelectNodes('//w:style[@w:type="table"]', $nsm)) {
    $styleId = $s.GetAttribute('styleId', $wNs)
    $nameNode = $s.SelectSingleNode('w:name', $nsm)
    $displayName = if ($nameNode) { W-Val $nameNode } else { $styleId }

    $identity = [ordered]@{
        StyleId = $styleId
        Type    = 'table'
        BasedOn = $null
        Next    = $null
        Custom  = if ($s.GetAttribute('customStyle', $wNs) -eq '1') { 'true' } else { 'false' }
        Default = if ($s.GetAttribute('default', $wNs) -eq '1') { 'true' } else { 'false' }
    }
    $bo = $s.SelectSingleNode('w:basedOn', $nsm); if ($bo) { $identity['BasedOn'] = W-Val $bo }
    $nx = $s.SelectSingleNode('w:next', $nsm); if ($nx) { $identity['Next'] = W-Val $nx }

    # Whole-table defaults
    $wholeFont = [ordered]@{}; foreach ($a in $FontAttrs) { $wholeFont[$a] = $null }
    $wholePara = [ordered]@{}; foreach ($a in $ParaAttrs) { $wholePara[$a] = $null }
    $wholeTbl  = [ordered]@{}; foreach ($a in $TableAttrs) { $wholeTbl[$a] = $null }

    Read-RunProps -RPr $s.SelectSingleNode('w:rPr', $nsm) -Nsm $nsm -V $wholeFont
    Read-ParaProps -PPr $s.SelectSingleNode('w:pPr', $nsm) -Nsm $nsm -V $wholePara
    Read-TableProps -TblPr $s.SelectSingleNode('w:tblPr', $nsm) -Nsm $nsm -V $wholeTbl

    # Conditional blocks
    $conditionals = [ordered]@{}
    foreach ($cb in $s.SelectNodes('w:tblStylePr', $nsm)) {
        $cType = $cb.GetAttribute('type', $wNs)
        $cFont = [ordered]@{}; foreach ($a in $FontAttrs) { $cFont[$a] = $null }
        $cPara = [ordered]@{}; foreach ($a in $ParaAttrs) { $cPara[$a] = $null }
        $cTbl  = [ordered]@{}; foreach ($a in $TableAttrs) { $cTbl[$a] = $null }
        $cRow  = [ordered]@{}; foreach ($a in $RowAttrs) { $cRow[$a] = $null }
        $cCell = [ordered]@{}; foreach ($a in $CellAttrs) { $cCell[$a] = $null }

        Read-RunProps -RPr $cb.SelectSingleNode('w:rPr', $nsm) -Nsm $nsm -V $cFont
        Read-ParaProps -PPr $cb.SelectSingleNode('w:pPr', $nsm) -Nsm $nsm -V $cPara
        Read-TableProps -TblPr $cb.SelectSingleNode('w:tblPr', $nsm) -Nsm $nsm -V $cTbl
        Read-RowProps -TrPr $cb.SelectSingleNode('w:trPr', $nsm) -Nsm $nsm -V $cRow
        Read-CellProps -TcPr $cb.SelectSingleNode('w:tcPr', $nsm) -Nsm $nsm -V $cCell

        $conditionals[$cType] = [pscustomobject]@{
            Font = $cFont; Para = $cPara; Tbl = $cTbl; Row = $cRow; Cell = $cCell
        }
    }

    $tableStyles += [pscustomobject]@{
        Name         = $displayName
        StyleId      = $styleId
        Identity     = $identity
        WholeFont    = $wholeFont
        WholePara    = $wholePara
        WholeTbl     = $wholeTbl
        Conditionals = $conditionals
    }
}

$tableStyles = $tableStyles | Sort-Object Name

# ---------------------------------------------------------------------------
# Render Markdown
# ---------------------------------------------------------------------------

function Render-Block {
    param([string]$Heading, [System.Collections.IDictionary]$Vals, [System.Text.StringBuilder]$Sb)
    [void]$Sb.AppendLine("**$Heading**")
    [void]$Sb.AppendLine('')
    [void]$Sb.AppendLine('```')
    foreach ($k in $Vals.Keys) {
        $v = $Vals[$k]
        if ($null -eq $v -or $v -eq '') { $v = '(not set)' }
        [void]$Sb.AppendLine(("{0,-22} {1}" -f ($k + ':'), $v))
    }
    [void]$Sb.AppendLine('```')
    [void]$Sb.AppendLine('')
}

$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine("# Table styles in $fileName")
[void]$md.AppendLine('')
[void]$md.AppendLine("Generated $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))  ")
[void]$md.AppendLine("$($tableStyles.Count) table style(s)")
[void]$md.AppendLine('')

foreach ($t in $tableStyles) {
    [void]$md.AppendLine("## $($t.Name)")
    [void]$md.AppendLine('')
    Render-Block -Heading 'Identity' -Vals $t.Identity -Sb $md
    [void]$md.AppendLine('### Whole-table defaults')
    [void]$md.AppendLine('')
    Render-Block -Heading 'Font'      -Vals $t.WholeFont -Sb $md
    Render-Block -Heading 'Paragraph' -Vals $t.WholePara -Sb $md
    Render-Block -Heading 'Table'     -Vals $t.WholeTbl  -Sb $md

    if ($t.Conditionals.Count -gt 0) {
        [void]$md.AppendLine('### Conditional formatting')
        [void]$md.AppendLine('')
        # Render in canonical order using $ConditionLabels, plus any unknown ones at the end
        $emitted = @()
        foreach ($cType in $ConditionLabels.Keys) {
            if ($t.Conditionals.Contains($cType)) {
                $label = $ConditionLabels[$cType]
                $cond = $t.Conditionals[$cType]
                [void]$md.AppendLine("#### $label")
                [void]$md.AppendLine('')
                Render-Block -Heading 'Font'      -Vals $cond.Font -Sb $md
                Render-Block -Heading 'Paragraph' -Vals $cond.Para -Sb $md
                Render-Block -Heading 'Table'     -Vals $cond.Tbl  -Sb $md
                Render-Block -Heading 'Row'       -Vals $cond.Row  -Sb $md
                Render-Block -Heading 'Cell'      -Vals $cond.Cell -Sb $md
                $emitted += $cType
            }
        }
        foreach ($cType in $t.Conditionals.Keys) {
            if ($emitted -notcontains $cType) {
                [void]$md.AppendLine("#### $cType")
                [void]$md.AppendLine('')
                $cond = $t.Conditionals[$cType]
                Render-Block -Heading 'Font'      -Vals $cond.Font -Sb $md
                Render-Block -Heading 'Paragraph' -Vals $cond.Para -Sb $md
                Render-Block -Heading 'Table'     -Vals $cond.Tbl  -Sb $md
                Render-Block -Heading 'Row'       -Vals $cond.Row  -Sb $md
                Render-Block -Heading 'Cell'      -Vals $cond.Cell -Sb $md
            }
        }
    }
}

$mdPath = Join-Path $OutputDir ($baseName + '.table-styles.md')
$md.ToString() | Set-Content -LiteralPath $mdPath -Encoding UTF8
Write-Host ""
Write-Host "Wrote $($tableStyles.Count) table styles to: $mdPath" -ForegroundColor Green

if ($Html) {
    $htmlPath = Join-Path $OutputDir ($baseName + '.table-styles.html')
    Write-Host "(HTML: $htmlPath -- using Markdown-to-HTML conversion is up to you. The .md file is the canonical output.)" -ForegroundColor DarkGray
}
