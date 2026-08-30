<#
.SYNOPSIS
    Builds NIST-reference.docx: a pandoc reference document that carries the
    NIST template's theme, page setup, numbering, and styles, plus definitions
    for every style name pandoc looks up by name.

.DESCRIPTION
    Pandoc's docx writer resolves styles from the reference document BY NAME,
    not by styleId. It looks for names like "Source Code", "Block Text",
    "Image Caption", and "First Paragraph". If a name is missing from the
    reference doc, pandoc invents its own style with its own look, which then
    has to be cleaned out of every converted file.

    This script starts from the NIST template, copies it to an output .docx,
    and adds a definition for each pandoc style name that is not already
    present, based on the closest NIST style so it inherits the right font,
    color, and spacing. Styles that already exist by name are left untouched,
    so re-running is safe and the template stays the source of truth.

    What comes along automatically by starting from the template: the theme,
    the section properties (page size and mirror margins), numbering.xml, and
    the header and footer parts. Pandoc reuses all of them.

    Deliberately NOT added: a table style named "Table". Pandoc would bind to
    it, but Set-NISTStyleArchitecture's table-style allowlist deletes every
    table style except TableNormal, TableGrid, HM-Cover, HM-Data, NIST-Data,
    and SidebarTable, so it would not survive the build. Convert-ToNISTWord.ps1
    assigns NIST-Data to converted tables instead.

    Operates on the ZIP package with System.IO.Compression. Word is not needed.

.PARAMETER Template
    Path to the NIST template (.docx or .dotx) to build the reference doc from.

.PARAMETER Output
    Path for the reference document to write. Defaults to NIST-reference.docx
    beside this script, which is also where Convert-ToNISTWord.ps1 looks for it.

.PARAMETER ApplyStyleArchitecture
    Run Set-NISTTemplateStyles.ps1 on the copy before adding the pandoc styles.
    Use when the template has not been through the style architecture yet. The
    pandoc styles are added afterwards so the Recommended-pane curation does not
    promote them into the authors' style list.

.PARAMETER Force
    Skip the confirmation prompt and overwrite an existing output file.

.EXAMPLE
    .\Build-NISTReferenceDoc.ps1 -Template "$env:USERPROFILE\Documents\HM-NCSTAR-Template.docx"

.EXAMPLE
    .\Build-NISTReferenceDoc.ps1 -Template "NCSTAR.dotx" -Output "C:\refs\NIST-reference.docx" -ApplyStyleArchitecture -Force
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Template,

    [string]$Output,

    [switch]$ApplyStyleArchitecture,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if (-not $Output) { $Output = Join-Path $PSScriptRoot 'NIST-reference.docx' }

# ---------------------------------------------------------------------------
# The pandoc style table
# ---------------------------------------------------------------------------
# One entry per style name the pandoc docx writer looks up. Fields:
#   Name    - the name pandoc searches for. Must match exactly.
#   Id      - styleId to create if the name is absent.
#   Type    - paragraph or character.
#   Builtin - $true for Word built-in style names (no w:customStyle="1"), so
#             Word keeps its own localized name and behavior for them.
#   Base    - ordered fallback list of NIST styleIds to base on. The first one
#             present in the template wins; Normal is the last resort.
#   PPr/RPr - extra paragraph / run properties for styles with no NIST twin.
#   Note    - what the style is for, printed in the summary.
#
# Anything with a true NIST equivalent is based on it and gets folded back into
# that NIST style by Convert-ToNISTWord.ps1. Only Source Code and Verbatim Char
# have no NIST counterpart and are meant to survive into finished documents.

$pandocStyles = @(
    @{ Name='First Paragraph';  Id='FirstParagraph'; Type='paragraph'; Builtin=$true;  Base=@('BodyText');
       Note='first paragraph after a heading' }

    @{ Name='Compact';          Id='Compact';        Type='paragraph'; Builtin=$false; Base=@('BodyText');
       PPr='<w:contextualSpacing/>'; Note='tight list item text' }

    @{ Name='Block Text';       Id='BlockText';      Type='paragraph'; Builtin=$true;  Base=@('Quote','BodyText');
       Note='block quotes' }

    @{ Name='Source Code';      Id='SourceCode';     Type='paragraph'; Builtin=$false; Base=@('BodyText');
       PPr='<w:keepLines/><w:spacing w:before="60" w:after="60"/><w:ind w:firstLine="0"/>';
       RPr='<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="18"/><w:szCs w:val="18"/>';
       Ligatures=$true; Note='code blocks (no NIST equivalent, kept)' }

    @{ Name='Verbatim Char';    Id='VerbatimChar';   Type='character'; Builtin=$false; Base=@('DefaultParagraphFont');
       RPr='<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="18"/><w:szCs w:val="18"/>';
       Ligatures=$true; Note='inline code (no NIST equivalent, kept)' }

    @{ Name='Image Caption';    Id='ImageCaption';   Type='paragraph'; Builtin=$false; Base=@('Caption');
       Note='figure captions' }

    @{ Name='Table Caption';    Id='TableCaption';   Type='paragraph'; Builtin=$false; Base=@('Caption');
       Note='table captions' }

    @{ Name='Figure';           Id='Figure';         Type='paragraph'; Builtin=$false; Base=@('graphic','BodyText');
       Note='paragraph holding an image' }

    @{ Name='Captioned Figure'; Id='CaptionedFigure';Type='paragraph'; Builtin=$false; Base=@('graphic','BodyText');
       Note='image plus its caption' }

    @{ Name='Author';           Id='Author';         Type='paragraph'; Builtin=$true;  Base=@('BodyText');
       PPr='<w:jc w:val="center"/>'; Note='title-block author line' }

    @{ Name='Date';             Id='Date';           Type='paragraph'; Builtin=$true;  Base=@('BodyText');
       PPr='<w:jc w:val="center"/>'; Note='title-block date line' }

    @{ Name='Abstract';         Id='Abstract';       Type='paragraph'; Builtin=$true;  Base=@('BodyText');
       Note='abstract body' }

    @{ Name='Abstract Title';   Id='AbstractTitle';  Type='paragraph'; Builtin=$false; Base=@('Frontmatterhead1','Heading1');
       Note='abstract heading' }

    @{ Name='Bibliography';     Id='Bibliography';   Type='paragraph'; Builtin=$true;  Base=@('BodyText');
       PPr='<w:ind w:left="720" w:hanging="720"/>'; Note='citeproc reference list' }

    @{ Name='Definition Term';  Id='DefinitionTerm'; Type='paragraph'; Builtin=$true;  Base=@('BodyText');
       RPr='<w:b/><w:bCs/>'; Note='definition-list term' }

    @{ Name='Definition';       Id='Definition';     Type='paragraph'; Builtin=$true;  Base=@('BodyText');
       PPr='<w:ind w:left="360"/>'; Note='definition-list body' }

    @{ Name='TOC Heading';      Id='TOCHeading';     Type='paragraph'; Builtin=$true;  Base=@('Frontmatterhead1','Heading1');
       Note='heading above a generated TOC' }
)

# ---------------------------------------------------------------------------
# Validate and copy
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $Template -PathType Leaf)) { throw "Template not found: $Template" }
$Template = (Resolve-Path -LiteralPath $Template).ProviderPath
$templateExt = [System.IO.Path]::GetExtension($Template).ToLowerInvariant()
if ($templateExt -notin @('.docx', '.dotx', '.docm', '.dotm')) {
    throw "Template must be a Word file (.docx/.dotx/.docm/.dotm): $Template"
}

$outDir = Split-Path -Parent $Output
if ($outDir -and -not (Test-Path -LiteralPath $outDir -PathType Container)) {
    throw "Output folder not found: $outDir"
}

Write-Host "Build NIST reference document" -ForegroundColor Cyan
Write-Host "  Template: $Template" -ForegroundColor Cyan
Write-Host "  Output:   $Output" -ForegroundColor Cyan
Write-Host "  Style architecture: $(if ($ApplyStyleArchitecture) { 'applied to the copy first' } else { 'taken from the template as-is' })" -ForegroundColor Cyan
Write-Host "  Adds $($pandocStyles.Count) pandoc style name(s) where missing." -ForegroundColor Cyan
Write-Host ""

if ((Test-Path -LiteralPath $Output) -and -not $Force) {
    $reply = Read-Host "Output exists. Overwrite? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
} elseif (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

if ($templateExt -in @('.docm', '.dotm')) {
    Write-Warning "Template is macro-enabled. Macros are carried into the reference doc but pandoc ignores them."
}

Copy-Item -LiteralPath $Template -Destination $Output -Force
$Output = (Resolve-Path -LiteralPath $Output).ProviderPath

if ($ApplyStyleArchitecture) {
    $styler = Join-Path $PSScriptRoot 'Set-NISTTemplateStyles.ps1'
    if (-not (Test-Path -LiteralPath $styler -PathType Leaf)) { throw "Styler not found: $styler" }
    Write-Host ""
    Write-Host "===== Applying NIST style architecture to the copy =====" -ForegroundColor Green
    & $styler -File $Output -Force
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Read the package
# ---------------------------------------------------------------------------

$entries   = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
$srcStream = [System.IO.File]::OpenRead($Output)
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

if (-not $entries.ContainsKey('word/styles.xml')) { throw "No word/styles.xml in $Output - not a Word package?" }

# A .dotx/.dotm carries the template content type. Retype it as a document so
# the reference doc opens as a normal .docx.
if ($templateExt -in @('.dotx', '.dotm') -and $entries.ContainsKey('[Content_Types].xml')) {
    $ct = $utf8NoBom.GetString($entries['[Content_Types].xml'])
    $ctNew = $ct.Replace('wordprocessingml.template.main+xml', 'wordprocessingml.document.main+xml')
    if ($ctNew -ne $ct) {
        $entries['[Content_Types].xml'] = $utf8NoBom.GetBytes($ctNew)
        Write-Host "  Content type retyped from template to document." -ForegroundColor DarkGray
    }
}

$xml = $utf8NoBom.GetString($entries['word/styles.xml'])

# w14:ligatures is only legal if the part declares the w14 namespace.
$hasW14 = $xml -match 'xmlns:w14\s*='

# ---------------------------------------------------------------------------
# Index existing styles by name and by id
# ---------------------------------------------------------------------------

# Case-insensitive lookup that returns the styleId in its ORIGINAL casing.
# Templates disagree on case (Graphic vs graphic) and a basedOn written in the
# wrong case points at nothing, so every id emitted below comes from here.
$existingIds   = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::OrdinalIgnoreCase)
$existingNames = @{}   # "type|name" -> styleId

$styleParse = [regex]'(?s)<w:style\s+([^>]*?)>(.*?)</w:style>'
foreach ($m in $styleParse.Matches($xml)) {
    $attrs = $m.Groups[1].Value
    $body  = $m.Groups[2].Value
    $sid = ''; $stype = ''; $sname = ''
    if ($attrs -match 'w:styleId="([^"]+)"') { $sid = $Matches[1] }
    if ($attrs -match 'w:type="([^"]+)"')    { $stype = $Matches[1] }
    if ($body  -match '<w:name w:val="([^"]+)"') { $sname = $Matches[1] }
    if (-not $sid) { continue }
    if (-not $existingIds.ContainsKey($sid)) { $existingIds[$sid] = $sid }
    # w:name may carry comma-separated aliases; each one is a lookup key.
    foreach ($alias in ($sname -split ',')) {
        $a = $alias.Trim()
        if ($a) { $existingNames["$stype|$($a.ToLowerInvariant())"] = $sid }
    }
}

Write-Host "  Template carries $($existingIds.Count) style(s)." -ForegroundColor DarkGray
Write-Host ""

# ---------------------------------------------------------------------------
# Add the missing pandoc styles
# ---------------------------------------------------------------------------

$added = @(); $present = @()

foreach ($ps in $pandocStyles) {
    $key = "$($ps.Type)|$($ps.Name.ToLowerInvariant())"
    if ($existingNames.ContainsKey($key)) {
        $present += [pscustomobject]@{ Name = $ps.Name; Id = $existingNames[$key]; Base = ''; Note = $ps.Note }
        continue
    }

    # First base that actually exists in this template; Normal as the backstop.
    $base = 'Normal'
    foreach ($candidate in $ps.Base) {
        if ($existingIds.ContainsKey($candidate)) { $base = $existingIds[$candidate]; break }
    }
    if ($ps.Type -eq 'character' -and $base -eq 'Normal') { $base = 'DefaultParagraphFont' }
    if ($existingIds.ContainsKey($base)) { $base = $existingIds[$base] }

    # Do not collide with an existing styleId that carries a different name.
    $id = $ps.Id
    if ($existingIds.ContainsKey($id)) {
        $n = 1
        while ($existingIds.ContainsKey("$id$n")) { $n++ }
        $id = "$id$n"
    }

    $customAttr = ''
    if (-not $ps.Builtin) { $customAttr = ' w:customStyle="1"' }

    $rpr = ''
    if ($ps.ContainsKey('RPr')) { $rpr = $ps.RPr }
    if ($ps.ContainsKey('Ligatures') -and $ps.Ligatures -and $hasW14) {
        $rpr += '<w14:ligatures w14:val="none"/>'
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<w:style w:type=`"$($ps.Type)`"$customAttr w:styleId=`"$id`">")
    [void]$sb.Append("<w:name w:val=`"$($ps.Name)`"/>")
    [void]$sb.Append("<w:basedOn w:val=`"$base`"/>")
    if ($ps.Type -eq 'paragraph' -and $existingIds.ContainsKey('BodyText')) {
        [void]$sb.Append("<w:next w:val=`"$($existingIds['BodyText'])`"/>")
    }
    [void]$sb.Append('<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>')
    if ($ps.ContainsKey('PPr') -and $ps.PPr) { [void]$sb.Append("<w:pPr>$($ps.PPr)</w:pPr>") }
    if ($rpr) { [void]$sb.Append("<w:rPr>$rpr</w:rPr>") }
    [void]$sb.Append('</w:style>')

    $xml = [regex]::Replace($xml, '</w:styles>', ($sb.ToString() + '</w:styles>'), 1)
    $existingIds[$id] = $id
    $added += [pscustomobject]@{ Name = $ps.Name; Id = $id; Base = $base; Note = $ps.Note }
}

$entries['word/styles.xml'] = $utf8NoBom.GetBytes($xml)

# ---------------------------------------------------------------------------
# Write the package back
# ---------------------------------------------------------------------------

$tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "nistref-$([guid]::NewGuid()).zip")
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

Move-Item -LiteralPath $tempZip -Destination $Output -Force

# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

Write-Host "Already in the template (left untouched):" -ForegroundColor Green
if ($present.Count -eq 0) { Write-Host "  none" -ForegroundColor DarkGray }
foreach ($p in $present) { Write-Host ("  {0,-18} -> {1}" -f $p.Name, $p.Id) -ForegroundColor DarkGray }

Write-Host ""
Write-Host "Added:" -ForegroundColor Green
if ($added.Count -eq 0) { Write-Host "  none" -ForegroundColor DarkGray }
foreach ($a in $added) { Write-Host ("  {0,-18} id {1,-16} basedOn {2,-18} {3}" -f $a.Name, $a.Id, $a.Base, $a.Note) }

Write-Host ""
Write-Host "Reference document written: $Output" -ForegroundColor Green
Write-Host "Use it with: .\Convert-ToNISTWord.ps1 -File input.html" -ForegroundColor Cyan
