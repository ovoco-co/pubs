<#
.SYNOPSIS
    Turns the paragraphs that sit above tables into NIST table captions, with
    an auto-numbering SEQ field.

.DESCRIPTION
    Converted documents often carry a heading above each table where a caption
    belongs. This finds paragraphs on -FromStyle that are immediately followed
    by a table, moves them onto -Style (Caption_Table), and prepends the NIST
    caption number, keeping the existing title text and its formatting.

    The result matches the production NCSTAR chapters exactly:

        Table 4-1.  Income Data Used for Weighting

    which is, character by character: the word Table, a non-breaking space, the
    chapter number, an en dash, a SEQ field, a period, and two spaces.

    Only the table number is a field. The chapter number is written as literal
    text from -Chapter, because a converted extract has no Heading 1 for a
    STYLEREF field to read. Pass -UseStyleRef once the content is merged into a
    real chapter, and the chapter number tracks the heading it lands under.

    The SEQ field is written with its computed value cached between the
    separate and end marks, so the numbers read correctly before Word has
    updated fields. They renumber on Ctrl+A then F9, or on print.

    A paragraph counts as a caption when the next body-level element is a
    table, ignoring empty paragraphs in between. A heading that introduces
    prose rather than a table is left alone, so section headings survive.

    Captions go ABOVE the table, which is the NIST convention for tables and
    the opposite of figures.

    Idempotent in practice: once converted, the paragraphs are no longer on
    -FromStyle, so a second run finds nothing.

    Backs up to <name>.captions.bak before writing.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Folder of Word files. All .docx/.dotx/.docm/.dotm in it are processed.

.PARAMETER FromStyle
    StyleId of the paragraphs to convert. Defaults to Heading1.

.PARAMETER Style
    StyleId of the caption style to apply. Defaults to CaptionTable, whose
    style name in the NIST template is Caption_Table.

.PARAMETER Chapter
    Chapter number written before the dash. Defaults to 1. Ignored when
    -UseStyleRef is passed.

.PARAMETER UseStyleRef
    Emit a STYLEREF 1 \s field for the chapter number instead of literal text,
    matching the production chapters. Only useful once the content sits under a
    real Heading 1.

.PARAMETER Label
    The word before the number. Defaults to Table.

.PARAMETER SeqName
    The SEQ sequence name. Defaults to Table. Keep it aligned with the rest of
    the document so numbering is continuous.

.PARAMETER ListOnly
    Report which paragraphs would become captions and write nothing. Worth a
    first pass on any document whose headings are a real section hierarchy: a
    section heading that happens to sit above a table matches the same rule as
    a caption, and only the title text tells them apart.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    # Always look before writing.
    .\Set-NISTTableCaptions.ps1 -File "descriptivetabs.docx" -ListOnly

.EXAMPLE
    .\Set-NISTTableCaptions.ps1 -File "descriptivetabs.docx" -Chapter 1

.EXAMPLE
    # Once merged into a real chapter, let the chapter number track the heading.
    .\Set-NISTTableCaptions.ps1 -File "Chapter-05.docx" -UseStyleRef
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$FromStyle = 'Heading1',

    [string]$Style = 'CaptionTable',

    [string]$Chapter = '1',

    [switch]$UseStyleRef,

    [string]$Label = 'Table',

    [string]$SeqName = 'Table',

    [switch]$ListOnly,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$NBSP   = [char]0x00A0
$ENDASH = [char]0x2013

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
        Sort-Object Name |
        ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }

# -------- Confirm --------
$sample = if ($UseStyleRef) { "$Label$NBSP{STYLEREF 1 \s}$ENDASH{SEQ $SeqName}.  <title>" }
          else              { "$Label$NBSP$Chapter$ENDASH{SEQ $SeqName}.  <title>" }

Write-Host "Set NIST Table Captions" -ForegroundColor Cyan
Write-Host "  Converts $FromStyle paragraphs that sit directly above a table into $Style." -ForegroundColor Cyan
Write-Host "  Caption form: $sample" -ForegroundColor Cyan
Write-Host "  Headings not followed by a table are left alone." -ForegroundColor Cyan
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""

if ($ListOnly) { Write-Host "  -ListOnly: reporting matches, writing nothing." -ForegroundColor Yellow }
Write-Host ""

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# Boundary patterns. The character after the element name must be whitespace or
# the closing bracket, so <w:pPr>, <w:tblPr>, and <w:tblGrid> are never matched.
$elemPattern = '(?s)<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>|</w:p>|<w:tbl(?:\s[^>]*)?>|</w:tbl>'

function Get-CaptionRuns {
    param([int]$Number)

    $sb = New-Object System.Text.StringBuilder

    # "Table" + non-breaking space, so the label never wraps off its number.
    [void]$sb.Append("<w:r><w:t xml:space=`"preserve`">$Label$NBSP</w:t></w:r>")

    if ($UseStyleRef) {
        [void]$sb.Append('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
        [void]$sb.Append('<w:r><w:instrText xml:space="preserve">STYLEREF 1 \s </w:instrText></w:r>')
        [void]$sb.Append('<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
        [void]$sb.Append("<w:r><w:rPr><w:noProof/></w:rPr><w:t>$Chapter</w:t></w:r>")
        [void]$sb.Append('<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    } else {
        [void]$sb.Append("<w:r><w:t>$Chapter</w:t></w:r>")
    }

    [void]$sb.Append("<w:r><w:t>$ENDASH</w:t></w:r>")

    # The table number itself, as a field so it renumbers on F9.
    [void]$sb.Append('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
    [void]$sb.Append("<w:r><w:instrText xml:space=`"preserve`"> SEQ $SeqName \* ARABIC \s 1 </w:instrText></w:r>")
    [void]$sb.Append('<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
    [void]$sb.Append("<w:r><w:rPr><w:noProof/></w:rPr><w:t>$Number</w:t></w:r>")
    [void]$sb.Append('<w:r><w:fldChar w:fldCharType="end"/></w:r>')

    # Period and the two spaces the production chapters use before the title.
    [void]$sb.Append('<w:r><w:t xml:space="preserve">.  </w:t></w:r>')

    return $sb.ToString()
}

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $entries   = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
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

    if (-not $entries.ContainsKey('word/document.xml')) { throw "No word/document.xml in $target" }
    if (-not $entries.ContainsKey('word/styles.xml'))   { throw "No word/styles.xml in $target" }

    # A reference to an undefined style renders as Normal, with no warning.
    $stylesXml = $utf8NoBom.GetString($entries['word/styles.xml'])
    if (-not [regex]::IsMatch($stylesXml, '<w:style\b[^>]*\bw:styleId="' + [regex]::Escape($Style) + '"')) {
        throw "Style '$Style' is not defined in $target. Apply the template's styles first (Set-NISTStyleArchitecture.ps1)."
    }

    $xml = $utf8NoBom.GetString($entries['word/document.xml'])

    # Work inside <w:body> only, so a sectPr or docPart elsewhere cannot confuse
    # the top-level scan.
    $bodyMatch = [regex]::Match($xml, '(?s)<w:body(?:\s[^>]*)?>.*</w:body>')
    if (-not $bodyMatch.Success) { throw "No <w:body> in $target" }
    $bodyStart = $bodyMatch.Index
    $body      = $bodyMatch.Value

    # Walk the body collecting top-level paragraphs and tables in document
    # order. Depth tracking keeps a table's own rows and any nested table out of
    # the list, so adjacency is measured between siblings.
    $elements = New-Object System.Collections.ArrayList
    $depth    = 0
    $openIdx  = -1
    $openKind = ''
    foreach ($m in [regex]::Matches($body, $elemPattern)) {
        $v = $m.Value
        if ($v -match '^<w:p(?:\s[^>]*)?/>$') {
            if ($depth -eq 0) {
                [void]$elements.Add(@{ Kind='p'; Start=$m.Index; End=$m.Index+$m.Length; Empty=$true })
            }
            continue
        }
        if ($v -eq '</w:p>' -or $v -eq '</w:tbl>') {
            $depth--
            if ($depth -eq 0 -and $openIdx -ge 0) {
                $end = $m.Index + $m.Length
                [void]$elements.Add(@{ Kind=$openKind; Start=$openIdx; End=$end; Empty=$false })
                $openIdx = -1
            }
            continue
        }
        if ($depth -eq 0) {
            $openIdx  = $m.Index
            $openKind = if ($v -like '<w:tbl*') { 'tbl' } else { 'p' }
        }
        $depth++
    }
    if ($depth -ne 0) { throw "Unbalanced paragraph or table tags in $target (depth $depth). File may be corrupt." }

    # Identify which top-level paragraphs are captions: on $FromStyle, with a
    # table as the next sibling once empty paragraphs are skipped.
    $captions = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $elements.Count; $i++) {
        $el = $elements[$i]
        if ($el.Kind -ne 'p' -or $el.Empty) { continue }
        $para = $body.Substring($el.Start, $el.End - $el.Start)
        if (-not [regex]::IsMatch($para, '<w:pStyle\s[^>]*w:val="' + [regex]::Escape($FromStyle) + '"')) { continue }

        $j = $i + 1
        while ($j -lt $elements.Count -and $elements[$j].Kind -eq 'p') {
            $nextPara = $body.Substring($elements[$j].Start, $elements[$j].End - $elements[$j].Start)
            $hasText  = [regex]::Match($nextPara, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>')
            if ($hasText.Success -and $hasText.Groups[1].Value.Trim().Length -gt 0) { break }
            $j++
        }
        if ($j -lt $elements.Count -and $elements[$j].Kind -eq 'tbl') { [void]$captions.Add($i) }
    }

    $tableCount = ($elements | Where-Object { $_.Kind -eq 'tbl' }).Count
    Write-Host "  Top-level tables: $tableCount" -ForegroundColor Gray
    Write-Host "  $FromStyle paragraphs directly above a table: $($captions.Count)" -ForegroundColor Gray

    # Print the titles. A section heading that happens to sit above a table
    # matches the same rule as a caption; the text is the only thing that tells
    # them apart, so it goes on screen either way.
    $seq = 0
    foreach ($idx in $captions) {
        $seq++
        $para = $body.Substring($elements[$idx].Start, $elements[$idx].End - $elements[$idx].Start)
        $title = -join ([regex]::Matches($para, '(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value })
        if ($title.Length -gt 60) { $title = $title.Substring(0, 57) + '...' }
        Write-Host ("    $Label$NBSP$(if($UseStyleRef){'#'}else{$Chapter})$ENDASH{0}.  {1}" -f $seq, $title) -ForegroundColor DarkGray
    }

    if ($captions.Count -eq 0 -or $ListOnly) {
        if ($captions.Count -eq 0) { Write-Host "  Nothing to convert. File unchanged." -ForegroundColor DarkGray }
        else { Write-Host "  -ListOnly: file unchanged." -ForegroundColor Yellow }
        continue
    }

    # Rewrite back to front so recorded offsets stay valid. Numbering is assigned
    # front to back first.
    $numbers = @{}
    for ($n = 0; $n -lt $captions.Count; $n++) { $numbers[$captions[$n]] = $n + 1 }

    $sb = New-Object System.Text.StringBuilder $body
    for ($n = $captions.Count - 1; $n -ge 0; $n--) {
        $idx  = $captions[$n]
        $el   = $elements[$idx]
        $para = $body.Substring($el.Start, $el.End - $el.Start)

        # Point the paragraph at the caption style.
        $newPara = [regex]::Replace(
            $para,
            '<w:pStyle\s[^>]*w:val="' + [regex]::Escape($FromStyle) + '"[^>]*/>',
            "<w:pStyle w:val=`"$Style`"/>",
            1)

        # Insert the numbering runs after <w:pPr>, ahead of the title runs, so
        # the existing text and its formatting are untouched.
        $pPrMatch = [regex]::Match($newPara, '(?s)<w:pPr(?:\s[^>]*)?>.*?</w:pPr>')
        if (-not $pPrMatch.Success) { throw "Caption paragraph lost its pPr in $target" }
        $insertAt = $pPrMatch.Index + $pPrMatch.Length
        $newPara  = $newPara.Insert($insertAt, (Get-CaptionRuns -Number $numbers[$idx]))

        [void]$sb.Remove($el.Start, $el.End - $el.Start)
        [void]$sb.Insert($el.Start, $newPara)
    }

    $newBody = $sb.ToString()
    $newXml  = $xml.Remove($bodyStart, $body.Length).Insert($bodyStart, $newBody)

    Write-Host "  Captions written: $($captions.Count) ($Label$NBSP$(if($UseStyleRef){'{STYLEREF}'}else{$Chapter})$ENDASH`1 through $($captions.Count))" -ForegroundColor Gray

    if (-not $Force -and (Test-Path -LiteralPath "$target.captions.bak")) {
        Write-Warning "  Backup $target.captions.bak already exists; leaving it as-is."
    } else {
        Copy-Item -LiteralPath $target -Destination "$target.captions.bak" -Force
    }

    $entries['word/document.xml'] = $utf8NoBom.GetBytes($newXml)

    $tempZip   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "captions-$([guid]::NewGuid()).zip")
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
    Write-Host "  Saved." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Select all and press F9 in Word to update the SEQ fields." -ForegroundColor Cyan
