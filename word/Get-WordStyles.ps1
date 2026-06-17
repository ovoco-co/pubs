<#
.SYNOPSIS
    Lists every style in a single Word file as a Markdown document. One
    section per style: style name as a heading, attributes listed below.

.DESCRIPTION
    Reads word/styles.xml directly from the package and writes a Markdown
    file where each style gets its own section. Within each section the
    attributes are organized into [Identity], [Font], [Paragraph], and
    [Table] subsections. Every relevant attribute is emitted whether or
    not it is set; unset attributes show as "_not set_" so it is clear
    what is omitted from the file vs explicitly defined.

    Sizes and indents are annotated with their human-readable equivalents
    (points, inches) alongside the raw OOXML values (half-points, twips).

.PARAMETER File
    Path to the .docx, .dotx, .docm, or .dotm file. Required.

.PARAMETER OutputDir
    Where to write the output. Defaults to a sibling "_styles" folder.

.PARAMETER Html
    Also produce an HTML version for browser viewing.

.PARAMETER FilterType
    Restrict output to one style type: paragraph, character, table, or
    numbering. Default is all types.

.EXAMPLE
    .\Get-WordStyles.ps1 -File "master.dotx"

.EXAMPLE
    .\Get-WordStyles.ps1 -File "master.dotx" -FilterType paragraph -Html
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [switch]$Html,

    [Parameter(Mandatory = $false)]
    [ValidateSet('paragraph', 'character', 'table', 'numbering', 'all')]
    [string]$FilterType = 'all'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $File -PathType Leaf)) {
    throw "File not found: $File"
}

$resolvedPath = (Resolve-Path -LiteralPath $File).ProviderPath

if ($resolvedPath -notmatch '\.(docx|dotx|docm|dotm)$') {
    throw "Unsupported extension. Path resolved to: $resolvedPath. Expected .docx, .dotx, .docm, or .dotm."
}

$dirName  = [System.IO.Path]::GetDirectoryName($resolvedPath)
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
$fileName = [System.IO.Path]::GetFileName($resolvedPath)

if (-not $OutputDir) {
    $OutputDir = Join-Path $dirName '_styles'
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ---------------------------------------------------------------------------
# Attribute schema. Each section names the attributes shown in that section,
# in the order they should appear. Sections shown depend on the style type.
# ---------------------------------------------------------------------------

$IdentityAttrs  = @('StyleId','Type','BasedOn','Next','Link','Aliases','Custom','Default')
$FontAttrs      = @('Font (ascii)','Font (hAnsi)','Font (theme)','Size','Bold','Italic','AllCaps','SmallCaps','Strikethrough','Underline','Color','ThemeColor','Highlight','VerticalAlign')
$ParagraphAttrs = @('Alignment','SpaceBefore','SpaceAfter','LineSpacing','IndentLeft','IndentRight','IndentFirstLine','IndentHanging','OutlineLevel','KeepWithNext','KeepLines','PageBreakBefore')
$TableAttrs     = @('TableWidth','RowBandSize')

# Which sections apply per style type
$SectionsByType = @{
    'paragraph' = @('Identity','Font','Paragraph')
    'character' = @('Identity','Font')
    'table'     = @('Identity','Font','Paragraph','Table')
    'numbering' = @('Identity','Font','Paragraph')
}

$NotSetText = '_not set_'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function W-Val { param([System.Xml.XmlElement]$N)
    if (-not $N) { return $null }
    $v = $N.GetAttribute('val', $wNs); if ($v) { return $v }
    return $N.GetAttribute('val')
}

function Read-Toggle {
    param([System.Xml.XmlElement]$Parent, [string]$Local, [System.Xml.XmlNamespaceManager]$Nsm)
    $n = $Parent.SelectSingleNode("w:$Local", $Nsm)
    if (-not $n) { return $null }
    $v = W-Val $n
    if ($null -eq $v -or $v -eq '' -or $v -eq '1' -or $v -eq 'true' -or $v -eq 'on') { return 'true' }
    return 'false'
}

function Format-Twips {
    param([string]$Twips, [string]$Mode = 'points')
    if ($null -eq $Twips -or $Twips -eq '') { return $null }
    $n = 0
    if (-not [int]::TryParse($Twips, [ref]$n)) { return "$Twips (raw)" }
    if ($Mode -eq 'inches') {
        $inch = [math]::Round($n / 1440.0, 3)
        return "$Twips twips ($inch in)"
    } else {
        $pt = [math]::Round($n / 20.0, 2)
        return "$Twips twips (${pt}pt)"
    }
}

function Format-HalfPt {
    param([string]$Hp)
    if ($null -eq $Hp -or $Hp -eq '') { return $null }
    $n = 0
    if (-not [int]::TryParse($Hp, [ref]$n)) { return "$Hp (raw)" }
    $pt = $n / 2.0
    return "${pt}pt ($Hp half-pts)"
}

# ---------------------------------------------------------------------------
# Open package, read styles.xml
# ---------------------------------------------------------------------------

Write-Host "Reading styles from $fileName..." -ForegroundColor Cyan

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

# ---------------------------------------------------------------------------
# Walk every style. For each, build a flat hash of every attribute the
# script knows about, populated where set and $null where not.
# ---------------------------------------------------------------------------

$styleEntries = @()

foreach ($s in $xml.SelectNodes('//w:style', $nsm)) {
    $type = $s.GetAttribute('type', $wNs)
    if ($FilterType -ne 'all' -and $type -ne $FilterType) { continue }

    $styleId = $s.GetAttribute('styleId', $wNs)
    $nameNode = $s.SelectSingleNode('w:name', $nsm)
    $displayName = if ($nameNode) { W-Val $nameNode } else { $styleId }

    # Initialize every possible attribute to $null
    $vals = [ordered]@{}
    foreach ($a in ($IdentityAttrs + $FontAttrs + $ParagraphAttrs + $TableAttrs)) { $vals[$a] = $null }

    # Identity
    $vals['StyleId']  = $styleId
    $vals['Type']     = $type
    $basedOn = $s.SelectSingleNode('w:basedOn', $nsm); if ($basedOn) { $vals['BasedOn'] = W-Val $basedOn }
    $next = $s.SelectSingleNode('w:next', $nsm); if ($next) { $vals['Next'] = W-Val $next }
    $link = $s.SelectSingleNode('w:link', $nsm); if ($link) { $vals['Link'] = W-Val $link }
    $aliases = $s.SelectSingleNode('w:aliases', $nsm); if ($aliases) { $vals['Aliases'] = W-Val $aliases }
    $vals['Custom']   = if ($s.GetAttribute('customStyle', $wNs) -eq '1') { 'true' } else { 'false' }
    $vals['Default']  = if ($s.GetAttribute('default', $wNs) -eq '1') { 'true' } else { 'false' }

    # Font / character
    $rPr = $s.SelectSingleNode('w:rPr', $nsm)
    if ($rPr) {
        $rFonts = $rPr.SelectSingleNode('w:rFonts', $nsm)
        if ($rFonts) {
            $a = $rFonts.GetAttribute('ascii', $wNs); if ($a) { $vals['Font (ascii)'] = $a }
            $h = $rFonts.GetAttribute('hAnsi', $wNs); if ($h) { $vals['Font (hAnsi)'] = $h }
            $t = $rFonts.GetAttribute('asciiTheme', $wNs); if ($t) { $vals['Font (theme)'] = $t }
        }
        $sz = $rPr.SelectSingleNode('w:sz', $nsm); if ($sz) { $vals['Size'] = Format-HalfPt (W-Val $sz) }

        $b = Read-Toggle -Parent $rPr -Local 'b' -Nsm $nsm; if ($null -ne $b) { $vals['Bold'] = $b }
        $i = Read-Toggle -Parent $rPr -Local 'i' -Nsm $nsm; if ($null -ne $i) { $vals['Italic'] = $i }
        $caps = Read-Toggle -Parent $rPr -Local 'caps' -Nsm $nsm; if ($null -ne $caps) { $vals['AllCaps'] = $caps }
        $sc = Read-Toggle -Parent $rPr -Local 'smallCaps' -Nsm $nsm; if ($null -ne $sc) { $vals['SmallCaps'] = $sc }
        $st = Read-Toggle -Parent $rPr -Local 'strike' -Nsm $nsm; if ($null -ne $st) { $vals['Strikethrough'] = $st }

        $u = $rPr.SelectSingleNode('w:u', $nsm); if ($u) { $vals['Underline'] = W-Val $u }

        $color = $rPr.SelectSingleNode('w:color', $nsm)
        if ($color) {
            $cv = W-Val $color
            $tc = $color.GetAttribute('themeColor', $wNs)
            if ($cv -and $cv -ne 'auto') { $vals['Color'] = "#$cv" }
            if ($tc) { $vals['ThemeColor'] = $tc }
        }

        $shd = $rPr.SelectSingleNode('w:shd', $nsm)
        if ($shd) {
            $fill = $shd.GetAttribute('fill', $wNs)
            if ($fill -and $fill -ne 'auto') { $vals['Highlight'] = "#$fill" }
        }

        $vert = $rPr.SelectSingleNode('w:vertAlign', $nsm); if ($vert) { $vals['VerticalAlign'] = W-Val $vert }
    }

    # Paragraph
    $pPr = $s.SelectSingleNode('w:pPr', $nsm)
    if ($pPr) {
        $jc = $pPr.SelectSingleNode('w:jc', $nsm); if ($jc) { $vals['Alignment'] = W-Val $jc }

        $spacing = $pPr.SelectSingleNode('w:spacing', $nsm)
        if ($spacing) {
            $b = $spacing.GetAttribute('before', $wNs); if ($b) { $vals['SpaceBefore'] = Format-Twips $b }
            $a = $spacing.GetAttribute('after', $wNs); if ($a) { $vals['SpaceAfter'] = Format-Twips $a }
            $line = $spacing.GetAttribute('line', $wNs)
            $rule = $spacing.GetAttribute('lineRule', $wNs)
            if ($line) {
                if ($rule -eq 'auto') {
                    $multi = [math]::Round([int]$line / 240.0, 2)
                    $vals['LineSpacing'] = "$line/240ths line ($multi x)"
                } else {
                    $vals['LineSpacing'] = "$line twips ($rule)"
                }
            }
        }

        $ind = $pPr.SelectSingleNode('w:ind', $nsm)
        if ($ind) {
            $l = $ind.GetAttribute('left', $wNs); if ($l) { $vals['IndentLeft'] = Format-Twips $l 'inches' }
            $r = $ind.GetAttribute('right', $wNs); if ($r) { $vals['IndentRight'] = Format-Twips $r 'inches' }
            $fl = $ind.GetAttribute('firstLine', $wNs); if ($fl) { $vals['IndentFirstLine'] = Format-Twips $fl 'inches' }
            $hg = $ind.GetAttribute('hanging', $wNs); if ($hg) { $vals['IndentHanging'] = Format-Twips $hg 'inches' }
        }

        $outline = $pPr.SelectSingleNode('w:outlineLvl', $nsm); if ($outline) { $vals['OutlineLevel'] = W-Val $outline }

        $kn = Read-Toggle -Parent $pPr -Local 'keepNext' -Nsm $nsm; if ($null -ne $kn) { $vals['KeepWithNext'] = $kn }
        $kl = Read-Toggle -Parent $pPr -Local 'keepLines' -Nsm $nsm; if ($null -ne $kl) { $vals['KeepLines'] = $kl }
        $pb = Read-Toggle -Parent $pPr -Local 'pageBreakBefore' -Nsm $nsm; if ($null -ne $pb) { $vals['PageBreakBefore'] = $pb }
    }

    # Table-only
    if ($type -eq 'table') {
        $tblPr = $s.SelectSingleNode('w:tblPr', $nsm)
        if ($tblPr) {
            $tblW = $tblPr.SelectSingleNode('w:tblW', $nsm)
            if ($tblW) {
                $w = $tblW.GetAttribute('w', $wNs)
                $t = $tblW.GetAttribute('type', $wNs)
                if ($w) { $vals['TableWidth'] = "$w ($t)" }
            }
            $rb = $tblPr.SelectSingleNode('w:tblStyleRowBandSize', $nsm)
            if ($rb) { $vals['RowBandSize'] = W-Val $rb }
        }
    }

    # Conditional formatting blocks for table styles
    $conditionals = @()
    if ($type -eq 'table') {
        foreach ($cb in $s.SelectNodes('w:tblStylePr', $nsm)) {
            $cType = $cb.GetAttribute('type', $wNs)
            $cVals = [ordered]@{}
            foreach ($a in $FontAttrs + @('Alignment')) { $cVals[$a] = $null }

            $cRPr = $cb.SelectSingleNode('w:rPr', $nsm)
            if ($cRPr) {
                $cSz = $cRPr.SelectSingleNode('w:sz', $nsm); if ($cSz) { $cVals['Size'] = Format-HalfPt (W-Val $cSz) }
                $cb1 = Read-Toggle -Parent $cRPr -Local 'b' -Nsm $nsm; if ($null -ne $cb1) { $cVals['Bold'] = $cb1 }
                $ci  = Read-Toggle -Parent $cRPr -Local 'i' -Nsm $nsm; if ($null -ne $ci) { $cVals['Italic'] = $ci }
                $cClr = $cRPr.SelectSingleNode('w:color', $nsm)
                if ($cClr) {
                    $cv = W-Val $cClr; $tc = $cClr.GetAttribute('themeColor', $wNs)
                    if ($cv -and $cv -ne 'auto') { $cVals['Color'] = "#$cv" }
                    if ($tc) { $cVals['ThemeColor'] = $tc }
                }
            }
            $cPPr = $cb.SelectSingleNode('w:pPr', $nsm)
            if ($cPPr) {
                $cJc = $cPPr.SelectSingleNode('w:jc', $nsm); if ($cJc) { $cVals['Alignment'] = W-Val $cJc }
            }

            $conditionals += [pscustomobject]@{
                Condition = $cType
                Vals      = $cVals
            }
        }
    }

    $styleEntries += [pscustomobject]@{
        Type         = $type
        Name         = $displayName
        StyleId      = $styleId
        Vals         = $vals
        Conditionals = $conditionals
    }
}

$styleEntries = $styleEntries | Sort-Object Type, Name

# ---------------------------------------------------------------------------
# Render Markdown
# ---------------------------------------------------------------------------

function Render-Section {
    param([string]$Heading, [array]$Attrs, [System.Collections.IDictionary]$Vals, [System.Text.StringBuilder]$Sb)

    [void]$Sb.AppendLine("**$Heading**")
    [void]$Sb.AppendLine('')
    [void]$Sb.AppendLine('```')
    foreach ($a in $Attrs) {
        $v = $Vals[$a]
        if ($null -eq $v -or $v -eq '') { $v = '(not set)' }
        [void]$Sb.AppendLine(("{0,-22} {1}" -f ($a + ':'), $v))
    }
    [void]$Sb.AppendLine('```')
    [void]$Sb.AppendLine('')
}

$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine("# Styles in $fileName")
[void]$md.AppendLine('')
[void]$md.AppendLine("Generated $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))  ")
[void]$md.AppendLine("$($styleEntries.Count) style(s)")
[void]$md.AppendLine('')

$lastType = $null
foreach ($entry in $styleEntries) {
    if ($entry.Type -ne $lastType) {
        [void]$md.AppendLine('')
        [void]$md.AppendLine("## $($entry.Type) styles")
        [void]$md.AppendLine('')
        $lastType = $entry.Type
    }

    [void]$md.AppendLine("### $($entry.Name)")
    [void]$md.AppendLine('')

    $sections = $SectionsByType[$entry.Type]
    if (-not $sections) { $sections = @('Identity','Font','Paragraph') }

    foreach ($sec in $sections) {
        $attrs = switch ($sec) {
            'Identity'  { $IdentityAttrs }
            'Font'      { $FontAttrs }
            'Paragraph' { $ParagraphAttrs }
            'Table'     { $TableAttrs }
        }
        Render-Section -Heading $sec -Attrs $attrs -Vals $entry.Vals -Sb $md
    }

    if ($entry.Conditionals.Count -gt 0) {
        [void]$md.AppendLine('**Conditional formatting**')
        [void]$md.AppendLine('')
        foreach ($c in $entry.Conditionals) {
            [void]$md.AppendLine("_$($c.Condition)_")
            [void]$md.AppendLine('')
            [void]$md.AppendLine('```')
            foreach ($a in ($FontAttrs + @('Alignment'))) {
                $v = $c.Vals[$a]
                if ($null -eq $v -or $v -eq '') { $v = '(not set)' }
                [void]$md.AppendLine(("{0,-22} {1}" -f ($a + ':'), $v))
            }
            [void]$md.AppendLine('```')
            [void]$md.AppendLine('')
        }
    }
}

$mdPath = Join-Path $OutputDir ($baseName + '.styles.md')
$md.ToString() | Set-Content -LiteralPath $mdPath -Encoding UTF8
Write-Host ""
Write-Host "Wrote $($styleEntries.Count) styles to: $mdPath" -ForegroundColor Green

# Console summary by type
Write-Host ""
Write-Host "By type:" -ForegroundColor Cyan
$styleEntries | Group-Object Type | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-12} {1}" -f $_.Name, $_.Count)
}

# ---------------------------------------------------------------------------
# Optional HTML
# ---------------------------------------------------------------------------

if ($Html) {
    $htmlPath = Join-Path $OutputDir ($baseName + '.styles.html')
    $h = New-Object System.Text.StringBuilder
    [void]$h.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8'><title>Styles in $fileName</title>")
    [void]$h.AppendLine(@"
<style>
  body { font-family: -apple-system, Segoe UI, Roboto, sans-serif; padding: 18px; color: #222; max-width: 900px; }
  h1 { font-size: 20px; }
  h2 { font-size: 16px; margin-top: 28px; padding-bottom: 4px; border-bottom: 2px solid #4a90e2; }
  h3 { font-size: 14px; margin-top: 18px; margin-bottom: 4px; color: #1a3d6c; }
  h4 { font-size: 12px; margin-top: 10px; margin-bottom: 4px; color: #444; text-transform: uppercase; letter-spacing: 0.04em; }
  pre { background: #f6f8fa; border: 1px solid #e1e4e8; padding: 8px 12px; font-size: 12px; line-height: 1.45; margin: 4px 0 12px 0; }
  .style-block { padding: 6px 0 12px 0; border-bottom: 1px dashed #ddd; }
  .meta { color: #666; font-size: 12px; }
  .cond { margin-left: 20px; border-left: 3px solid #ffeaa7; padding-left: 12px; margin-top: 6px; }
  .not-set { color: #aaa; font-style: italic; }
</style>
"@)
    [void]$h.AppendLine("</head><body>")
    [void]$h.AppendLine("<h1>Styles in $fileName</h1>")
    [void]$h.AppendLine("<div class='meta'>Generated $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) | $($styleEntries.Count) styles</div>")

    function Render-HtmlSection {
        param([string]$Heading, [array]$Attrs, [System.Collections.IDictionary]$Vals, [System.Text.StringBuilder]$Sb)
        [void]$Sb.AppendLine("<h4>$Heading</h4><pre>")
        foreach ($a in $Attrs) {
            $v = $Vals[$a]
            $isNotSet = ($null -eq $v -or $v -eq '')
            $display = if ($isNotSet) { '(not set)' } else { $v }
            $line = ("{0,-22} {1}" -f ($a + ':'), $display)
            $encoded = [System.Web.HttpUtility]::HtmlEncode($line)
            if ($isNotSet) {
                [void]$Sb.AppendLine("<span class='not-set'>$encoded</span>")
            } else {
                [void]$Sb.AppendLine($encoded)
            }
        }
        [void]$Sb.AppendLine('</pre>')
    }

    $lastType = $null
    foreach ($entry in $styleEntries) {
        if ($entry.Type -ne $lastType) {
            [void]$h.AppendLine("<h2>$($entry.Type) styles</h2>")
            $lastType = $entry.Type
        }
        [void]$h.AppendLine("<div class='style-block'>")
        [void]$h.AppendLine("<h3>$([System.Web.HttpUtility]::HtmlEncode($entry.Name))</h3>")

        $sections = $SectionsByType[$entry.Type]
        if (-not $sections) { $sections = @('Identity','Font','Paragraph') }
        foreach ($sec in $sections) {
            $attrs = switch ($sec) {
                'Identity'  { $IdentityAttrs }
                'Font'      { $FontAttrs }
                'Paragraph' { $ParagraphAttrs }
                'Table'     { $TableAttrs }
            }
            Render-HtmlSection -Heading $sec -Attrs $attrs -Vals $entry.Vals -Sb $h
        }

        if ($entry.Conditionals.Count -gt 0) {
            [void]$h.AppendLine("<h4>Conditional formatting</h4>")
            foreach ($c in $entry.Conditionals) {
                [void]$h.AppendLine("<div class='cond'><h4>$([System.Web.HttpUtility]::HtmlEncode($c.Condition))</h4><pre>")
                foreach ($a in ($FontAttrs + @('Alignment'))) {
                    $v = $c.Vals[$a]
                    $isNotSet = ($null -eq $v -or $v -eq '')
                    $display = if ($isNotSet) { '(not set)' } else { $v }
                    $line = ("{0,-22} {1}" -f ($a + ':'), $display)
                    $encoded = [System.Web.HttpUtility]::HtmlEncode($line)
                    if ($isNotSet) {
                        [void]$h.AppendLine("<span class='not-set'>$encoded</span>")
                    } else {
                        [void]$h.AppendLine($encoded)
                    }
                }
                [void]$h.AppendLine('</pre></div>')
            }
        }

        [void]$h.AppendLine("</div>")
    }
    [void]$h.AppendLine("</body></html>")
    $h.ToString() | Set-Content -LiteralPath $htmlPath -Encoding UTF8
    Write-Host "HTML view: $htmlPath" -ForegroundColor Green
}
