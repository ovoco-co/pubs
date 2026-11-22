<#
.SYNOPSIS
    Puts footnote reference marks on the FootnoteReference character style and
    footnote paragraphs on FootnoteText, replacing the direct formatting that
    was standing in for them.

.DESCRIPTION
    A footnote has three styled parts and a converted or hand-built document
    usually has none of them:

      1. the reference mark in the body text (the run holding
         <w:footnoteReference>), which belongs on the FootnoteReference
         character style
      2. the matching mark at the start of the footnote itself (the run holding
         <w:footnoteRef/>), same style
      3. the footnote paragraphs, which belong on FootnoteText

    Without 3 the footnote text has no pStyle at all and falls through to
    Normal, so footnotes render at body size instead of FootnoteText's smaller
    size. NWIRP chapter 3 had 16 of 17 footnotes in exactly that state, with
    the superscript faked by direct <w:vertAlign> on the reference runs.

    Direct formatting is removed ONLY where the style already supplies it:
    w:vertAlign (FootnoteReference is a superscript style by definition) and a
    plain black w:color. Everything else in the run - fonts, bold, italic -
    is left alone, because a reference mark can legitimately carry them and
    guessing is how real formatting gets destroyed. Pass -KeepDirect to add
    the styles and strip nothing.

.PARAMETER KeepDirect
    Add the styles without removing the redundant direct superscript/colour.

.PARAMETER ListOnly
    Report what would change and write nothing.

.EXAMPLE
    .\Set-FootnoteStyles.ps1 -File "C:\...\Chapter-03.docx" -ListOnly

.EXAMPLE
    .\Set-FootnoteStyles.ps1 -File "C:\...\Chapter-03.docx" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$ReferenceStyle = 'FootnoteReference',

    [string]$TextStyle = 'FootnoteText',

    [switch]$KeepDirect,

    [switch]$ListOnly,

    [switch]$Force,

    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $resolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($resolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|docm)$' -and $_.Name -notlike '~$*' } |
        ForEach-Object { $_.FullName } | Sort-Object
}
if ($targets.Count -eq 0) { throw "No Word files found." }

Write-Host "Set Footnote Styles" -ForegroundColor Cyan
Write-Host "  Reference marks -> $ReferenceStyle   Footnote paragraphs -> $TextStyle" -ForegroundColor Cyan
if (-not $KeepDirect) {
    Write-Host "  Redundant direct superscript and plain black colour removed from the marks." -ForegroundColor Cyan
}
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""
if ($ListOnly) { Write-Host "  -ListOnly: reporting only, writing nothing." -ForegroundColor Yellow }

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

function Read-Package {
    param([string]$Path)
    $names   = New-Object System.Collections.ArrayList
    $entries = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $stream  = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $entries[$e.FullName] = $ms.ToArray()
                [void]$names.Add($e.FullName)
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
    return ,[pscustomobject]@{ Names = $names; Data = $entries }
}

function Get-Text {
    param($Pkg, [string]$Name)
    if (-not $Pkg.Data.ContainsKey($Name)) { return $null }
    $t = [System.Text.Encoding]::UTF8.GetString($Pkg.Data[$Name])
    # No String.StartsWith(char) overload on .NET Framework; the coerced
    # comparison treats U+FEFF as ignorable and is always true.
    if ($t.Length -gt 0 -and $t[0] -eq [char]0xFEFF) { return @($t.Substring(1), $true) }
    return @($t, $false)
}

function Add-RunStyle {
    # w:rStyle is the FIRST child of w:rPr in CT_RPr, so it is inserted at the
    # front of an existing rPr, never appended.
    param([string]$Run, [string]$StyleId, [bool]$Strip)
    $rStyle = '<w:rStyle w:val="' + $StyleId + '"/>'
    $rpr = [regex]::Match($Run, '(?s)<w:rPr(?:\s[^>]*)?>.*?</w:rPr>')
    if ($rpr.Success) {
        $inner = $rpr.Value
        if ($inner -match '<w:rStyle\b') {
            $inner = ([regex]'<w:rStyle\b[^>]*/>').Replace($inner, $rStyle, 1)
        } else {
            $inner = ([regex]'(<w:rPr(?:\s[^>]*)?>)').Replace($inner, ('$1' + $rStyle), 1)
        }
        if ($Strip) {
            # Only what the style itself supplies.
            $inner = [regex]::Replace($inner, '<w:vertAlign\b[^>]*/>', '')
            $inner = [regex]::Replace($inner, '<w:color\b[^>]*\bw:val="000000"[^>]*/>', '')
        }
        $inner = [regex]::Replace($inner, '<w:rPr(?:\s[^>]*)?>\s*</w:rPr>', ('<w:rPr>' + $rStyle + '</w:rPr>'))
        return $Run.Remove($rpr.Index, $rpr.Length).Insert($rpr.Index, $inner)
    }
    $selfClosing = [regex]::Match($Run, '<w:rPr\s*/>')
    if ($selfClosing.Success) {
        return $Run.Remove($selfClosing.Index, $selfClosing.Length).Insert($selfClosing.Index, ('<w:rPr>' + $rStyle + '</w:rPr>'))
    }
    # w:rPr must be the first child of w:r.
    return ([regex]'^(<w:r(?:\s[^>]*)?>)').Replace($Run, ('$1<w:rPr>' + $rStyle + '</w:rPr>'), 1)
}

function Set-ParaStyle {
    # w:pStyle is the first child of w:pPr, and w:pPr the first child of w:p.
    param([string]$Para, [string]$StyleId)
    $pStyle = '<w:pStyle w:val="' + $StyleId + '"/>'
    $ppr = [regex]::Match($Para, '(?s)<w:pPr(?:\s[^>]*)?>.*?</w:pPr>')
    if ($ppr.Success) {
        $inner = $ppr.Value
        if ($inner -match '<w:pStyle\b') {
            $inner = ([regex]'<w:pStyle\b[^>]*/>').Replace($inner, $pStyle, 1)
        } else {
            $inner = ([regex]'(<w:pPr(?:\s[^>]*)?>)').Replace($inner, ('$1' + $pStyle), 1)
        }
        return $Para.Remove($ppr.Index, $ppr.Length).Insert($ppr.Index, $inner)
    }
    $selfClosing = [regex]::Match($Para, '<w:pPr\s*/>')
    if ($selfClosing.Success) {
        return $Para.Remove($selfClosing.Index, $selfClosing.Length).Insert($selfClosing.Index, ('<w:pPr>' + $pStyle + '</w:pPr>'))
    }
    return ([regex]'^(<w:p(?:\s[^>]*)?>)').Replace($Para, ('$1<w:pPr>' + $pStyle + '</w:pPr>'), 1)
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$filesChanged = 0

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $pkg = Read-Package -Path $target
    $styles = (Get-Text -Pkg $pkg -Name 'word/styles.xml')[0]
    if (-not $styles) { Write-Warning "  No word/styles.xml. Skipping."; continue }
    foreach ($sid in @($ReferenceStyle, $TextStyle)) {
        if ($styles -notmatch ('<w:style\b[^>]*\bw:styleId="' + [regex]::Escape($sid) + '"')) {
            throw "Style $sid is not defined in ${target}; run Set-NISTStyleArchitecture.ps1 first."
        }
    }

    $changed = $false
    $report = New-Object System.Collections.ArrayList

    # --- 1. reference marks in the body ---------------------------------
    $docPair = Get-Text -Pkg $pkg -Name 'word/document.xml'
    if ($docPair) {
        $doc = $docPair[0]; $docBom = $docPair[1]
        $n = 0
        $sb = New-Object System.Text.StringBuilder; $pos = 0
        foreach ($m in [regex]::Matches($doc, '(?s)<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:footnoteReference\b[^>]*/>.*?</w:r>')) {
            if ($m.Value -match ('<w:rStyle\b[^>]*\bw:val="' + [regex]::Escape($ReferenceStyle) + '"')) { continue }
            $new = Add-RunStyle -Run $m.Value -StyleId $ReferenceStyle -Strip (-not $KeepDirect)
            [void]$sb.Append($doc.Substring($pos, $m.Index - $pos))
            [void]$sb.Append($new)
            $pos = $m.Index + $m.Length
            $n++
        }
        [void]$sb.Append($doc.Substring($pos))
        if ($n -gt 0) {
            $doc = $sb.ToString()
            $pkg.Data['word/document.xml'] = $utf8NoBom.GetBytes($(if ($docBom) { [string][char]0xFEFF + $doc } else { $doc }))
            $changed = $true
        }
        [void]$report.Add("body reference marks -> ${ReferenceStyle}: $n")
    }

    # --- 2 and 3. the footnotes part -------------------------------------
    $fnPair = Get-Text -Pkg $pkg -Name 'word/footnotes.xml'
    if ($fnPair) {
        $fn = $fnPair[0]; $fnBom = $fnPair[1]
        $marks = 0; $paras = 0
        $sb = New-Object System.Text.StringBuilder; $pos = 0
        # Separators and continuation notices carry w:type and are not content.
        foreach ($m in [regex]::Matches($fn, '(?s)<w:footnote\b[^>]*>.*?</w:footnote>')) {
            $block = $m.Value
            if ($block -match '<w:footnote\b[^>]*\bw:type=') {
                continue
            }
            $new = $block
            $rm = [regex]::Match($new, '(?s)<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:footnoteRef\s*/>.*?</w:r>')
            if ($rm.Success -and $rm.Value -notmatch ('<w:rStyle\b[^>]*\bw:val="' + [regex]::Escape($ReferenceStyle) + '"')) {
                $fixed = Add-RunStyle -Run $rm.Value -StyleId $ReferenceStyle -Strip (-not $KeepDirect)
                $new = $new.Remove($rm.Index, $rm.Length).Insert($rm.Index, $fixed)
                $marks++
            }
            $psb = New-Object System.Text.StringBuilder; $ppos = 0; $hit = 0
            foreach ($pm in [regex]::Matches($new, '(?s)<w:p\b[^>]*>.*?</w:p>')) {
                $st = [regex]::Match($pm.Value, '<w:pStyle\b[^>]*\bw:val="([^"]+)"')
                if ($st.Success -and $st.Groups[1].Value -eq $TextStyle) { continue }
                $fixedP = Set-ParaStyle -Para $pm.Value -StyleId $TextStyle
                [void]$psb.Append($new.Substring($ppos, $pm.Index - $ppos))
                [void]$psb.Append($fixedP)
                $ppos = $pm.Index + $pm.Length
                $hit++
            }
            if ($hit -gt 0) {
                [void]$psb.Append($new.Substring($ppos))
                $new = $psb.ToString()
                $paras += $hit
            }
            if ($new -ne $block) {
                [void]$sb.Append($fn.Substring($pos, $m.Index - $pos))
                [void]$sb.Append($new)
                $pos = $m.Index + $m.Length
            }
        }
        [void]$sb.Append($fn.Substring($pos))
        if ($marks -gt 0 -or $paras -gt 0) {
            $fn = $sb.ToString()
            $pkg.Data['word/footnotes.xml'] = $utf8NoBom.GetBytes($(if ($fnBom) { [string][char]0xFEFF + $fn } else { $fn }))
            $changed = $true
        }
        [void]$report.Add("footnote marks -> ${ReferenceStyle}: $marks")
        [void]$report.Add("footnote paragraphs -> ${TextStyle}: $paras")
    }

    foreach ($r in $report) { Write-Host "    $r" -ForegroundColor DarkGray }

    if (-not $changed) { Write-Host "  Already styled. File unchanged." -ForegroundColor Green; continue }

    foreach ($p in @('word/document.xml', 'word/footnotes.xml')) {
        if (-not $pkg.Data.ContainsKey($p)) { continue }
        $t = [System.Text.Encoding]::UTF8.GetString($pkg.Data[$p])
        if ($t.Length -gt 0 -and $t[0] -eq [char]0xFEFF) { $t = $t.Substring(1) }
        try {
            $probe = New-Object System.Xml.XmlDocument
            $probe.XmlResolver = $null
            $probe.LoadXml($t)
        } catch {
            throw "Rewrite produced malformed $p for ${target}: $($_.Exception.Message)"
        }
    }

    if ($ListOnly) { continue }

    if (-not $NoBackup) {
        if (-not $Force -and (Test-Path -LiteralPath "$target.footnotes.bak")) {
            Write-Warning "  Backup $target.footnotes.bak already exists; leaving it as-is."
        } else {
            Copy-Item -LiteralPath $target -Destination "$target.footnotes.bak" -Force
        }
    }

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "footnotes-$([guid]::NewGuid()).zip")
    $outStream = [System.IO.File]::Create($tempZip)
    try {
        $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $pkg.Names) {
                $ne = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try { $os.Write($pkg.Data[$name], 0, $pkg.Data[$name].Length) } finally { $os.Dispose() }
            }
        } finally { $outZip.Dispose() }
    } finally { $outStream.Dispose() }

    Move-Item -LiteralPath $tempZip -Destination $target -Force
    $filesChanged++
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host ("Done. {0} file(s) changed." -f $filesChanged) -ForegroundColor Green
