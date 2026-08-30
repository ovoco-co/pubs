<#
.SYNOPSIS
    Strips direct character formatting from runs so paragraph styles render
    cleanly. Run after authors hand documents back. Preserves superscript /
    subscript and leaves MathType and OLE-object runs untouched.

.DESCRIPTION
    Authors editing in Word often accumulate direct character formatting on
    individual runs (font, size, color, bold, italic via the Font Format
    dialog or ribbon overrides). The result is that paragraph styles cannot
    show through. This script wipes the <w:rPr> block from every run in the
    document body, headers, footers, footnotes, and endnotes. Equivalent of
    selecting all text and pressing Ctrl+Space (Clear character formatting)
    in Word, with two preservations.

    Preserved automatically:
      - Italic (<w:i/> and <w:iCs/>). Book titles, scientific names,
        symbols, and manual emphasis stay italic.
      - Superscript and subscript (<w:vertAlign w:val="superscript"/> or
        <w:vertAlign w:val="subscript"/>). Chemical formulas and exponents
        stay correct.
      - Footnote and endnote reference marks (the <w:rStyle> pointing at the
        FootnoteReference / EndnoteReference character style). These carry the
        superscript for the inline callout numbers and the note numbers, so
        that rStyle is kept even though other character styles are cleared.
      - Runs that contain OLE objects, MathType equations, or Word equations
        (<w:object>, <o:OLEObject>, <m:oMath>, <m:oMathPara>). These runs
        are left entirely alone, including their rPr, since MathType and
        the equation editor depend on specific run formatting to render.

    On a preserved run, only the listed elements survive in rPr; everything
    else (font, size, color, bold, language, highlighting, and character
    styles other than FootnoteReference / EndnoteReference) is cleared.

    Paragraph mark formatting (<w:rPr> inside <w:pPr>) is left alone, so
    paragraph-level mark properties stay intact.

    Other intentional inline formatting (italic on book titles, hyperlink
    styling, manual emphasis, language tags) is cleared. Reapply where
    needed in Word. Hyperlinks remain functional but lose their styling;
    reapply the Hyperlink character style as needed.

    Idempotent: re-running on a cleaned file does nothing.

    Reports per-part count of runs cleaned plus a total per file.

.PARAMETER File
    Path to a single Word file.

.PARAMETER Folder
    Folder of Word files. All .docx/.dotx/.docm/.dotm in folder are processed.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Reapply-Styles.ps1 -File "Chapter-05.docx"

.EXAMPLE
    .\Reapply-Styles.ps1 -Folder "C:\chapters"
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

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

# -------- Confirm --------
Write-Host "Reapply Styles" -ForegroundColor Cyan
Write-Host "  Strips run-level character formatting from document, headers, footers, foot/endnotes." -ForegroundColor Cyan
Write-Host "  Preserves: italic, superscript/subscript, MathType and OLE-object runs." -ForegroundColor Cyan
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# Outer pattern: each run (open through close, handling internal content).
# Self-closing <w:r/> runs are skipped (they have no rPr).
$runPattern = '(?si)<w:r\b[^>]*>.*?</w:r>'

# Inner pattern: rPr block immediately after the run open tag.
$rPrPattern = '(?si)(<w:r\b[^>]*>)\s*<w:rPr\b(?:\s+[^>]*?)?(?:/>|>(.*?)</w:rPr>)'

# Per-run evaluator: skip OLE/math runs, otherwise strip rPr while preserving
# any <w:vertAlign> element (superscript / subscript).
$runEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($runMatch)
    $run = $runMatch.Value

    # Preserve runs that carry OLE objects, MathType equations, or Word equations
    if ($run -match '(?i)<w:object\b'      -or
        $run -match '(?i)<o:OLEObject\b'   -or
        $run -match '(?i)<m:oMath\b'       -or
        $run -match '(?i)<m:oMathPara\b') {
        return $run
    }

    # Strip the rPr block but preserve vertAlign if present
    return [regex]::Replace(
        $run,
        $rPrPattern,
        [System.Text.RegularExpressions.MatchEvaluator]{
            param($rPrMatch)
            $runOpen  = $rPrMatch.Groups[1].Value
            $rPrInner = if ($rPrMatch.Groups[2].Success) { $rPrMatch.Groups[2].Value } else { '' }

            # Collect any preservation-eligible elements from the rPr.
            # Order them as Word writes them: rStyle (footnote/endnote refs
            # only), rFonts (skipped), b (skipped), i, iCs, ..., vertAlign.
            $kept = New-Object System.Text.StringBuilder

            # Footnote and endnote reference marks - the little superscript
            # callout numbers in the body and the note numbers in the notes -
            # get their superscript from the FootnoteReference / EndnoteReference
            # character style, not from direct formatting. Keep that rStyle so
            # the superscripts survive. Other character styles are still cleared.
            $rStyleMatch = [regex]::Match($rPrInner, '(?i)<w:rStyle\b[^>]*\bw:val="(?:Footnote|Endnote)Reference"[^>]*/>')
            if ($rStyleMatch.Success) { [void]$kept.Append($rStyleMatch.Value) }

            $iMatch  = [regex]::Match($rPrInner, '<w:i\b(?:\s+[^/>]*)?/?>')
            if ($iMatch.Success)  { [void]$kept.Append($iMatch.Value) }

            $iCsMatch = [regex]::Match($rPrInner, '<w:iCs\b(?:\s+[^/>]*)?/?>')
            if ($iCsMatch.Success) { [void]$kept.Append($iCsMatch.Value) }

            $vaMatch  = [regex]::Match($rPrInner, '<w:vertAlign\b[^/>]*/?>')
            if ($vaMatch.Success)  { [void]$kept.Append($vaMatch.Value) }

            if ($kept.Length -gt 0) {
                return "$runOpen<w:rPr>$($kept.ToString())</w:rPr>"
            } else {
                return $runOpen
            }
        }
    )
}

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

    # Identify text-bearing parts
    $textParts = @($entries.Keys | Where-Object {
        $_ -match '^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$'
    })

    $totalCleaned   = 0
    $totalPreserved = 0
    foreach ($part in $textParts) {
        $xml = $utf8NoBom.GetString($entries[$part])

        # Count preservation candidates and rPr-bearing runs before
        $beforeRPr = ([regex]::Matches($xml, '(?si)<w:r\b[^>]*>\s*<w:rPr\b')).Count
        if ($beforeRPr -eq 0) {
            Write-Host "  $part : already clean" -ForegroundColor DarkGray
            continue
        }

        $cleanedXml = [regex]::Replace($xml, $runPattern, $runEvaluator)
        $afterRPr   = ([regex]::Matches($cleanedXml, '(?si)<w:r\b[^>]*>\s*<w:rPr\b')).Count
        $cleaned    = $beforeRPr - $afterRPr
        # Surviving rPr blocks are math/OLE runs or vertAlign-only stubs we kept.
        $entries[$part] = $utf8NoBom.GetBytes($cleanedXml)
        Write-Host "  $part : cleaned $cleaned, preserved $afterRPr" -ForegroundColor Gray
        $totalCleaned   += $cleaned
        $totalPreserved += $afterRPr
    }
    Write-Host "  Total: cleaned $totalCleaned run(s), preserved $totalPreserved." -ForegroundColor Green

    # Write zip back
    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "reapply-$([guid]::NewGuid()).zip")
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
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Note: hyperlinks remain functional but lose styling. Reapply Hyperlink character style in Word as needed." -ForegroundColor Cyan
