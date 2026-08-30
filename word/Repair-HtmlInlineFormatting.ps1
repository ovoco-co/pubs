<#
.SYNOPSIS
    Rewrites CSS-only inline formatting in HTML into real semantic tags, so a
    converter that reads tags (pandoc) stops dropping it.

.DESCRIPTION
    R flextable, and any generator like it, emits no semantic markup at all.
    Every run of text is a <span class="cl-xxxxxxxx"> and the bold, italic,
    superscript, subscript, and underline live only in the generated CSS class:

        .cl-1a2b3c4d{font-family:'Arial';font-size:7pt;font-weight:bold;
                     vertical-align:super;...}

    Pandoc's HTML reader looks at tags, not at stylesheets, so all of it is
    silently flattened on conversion. The text survives; the emphasis does not.

    This script reads the <style> blocks, works out which classes carry
    formatting, and rewrites each affected <span> into the equivalent tags:

        <span class="cl-1a2b3c4d">n</span>   ->   <strong><sup>n</sup></strong>

    The rewrite is a SINGLE pass, which matters. A class is very often both
    bold and superscript at once, and running one pass per tag consumes the
    span on the first pass and loses the rest. Every flag a class carries is
    emitted as one nest of tags in one substitution.

    Spans whose classes carry no formatting are left exactly as they are;
    pandoc handles a plain span fine.

    What this does NOT recover is indentation. flextable carries table row
    hierarchy as leading whitespace, which pandoc also drops - use
    Restore-TableIndents.ps1 for that.

    Idempotent: tags already emitted are not inside a class span any more, so
    re-running finds nothing to do. Backs up to <name>.inline.bak.

.PARAMETER File
    Path to a single .html/.htm/.xhtml file. Either this or -Folder is required.

.PARAMETER Folder
    Folder of HTML files. All .html/.htm/.xhtml in it are processed.

.PARAMETER ListOnly
    Report the formatting classes found and how many spans each one covers,
    and write nothing. Run this first on an unfamiliar file.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Repair-HtmlInlineFormatting.ps1 -File "chapter-05.html" -ListOnly

.EXAMPLE
    .\Repair-HtmlInlineFormatting.ps1 -Folder "C:\src\Chapter5" -Force
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [switch]$ListOnly,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Build target list
# ---------------------------------------------------------------------------

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(html|htm|xhtml)$' } |
        ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No HTML files found." }

Write-Host "Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "  $_" }
Write-Host ""

if (-not $ListOnly -and -not $Force) {
    $reply = Read-Host "Rewrite CSS-only formatting into tags? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# CSS declaration -> tag flags
#
# Nesting order is outermost first. sup/sub go innermost so they wrap the bare
# text, which is what Word's docx writer expects to turn into a vertAlign run.
# ---------------------------------------------------------------------------

$FLAGS = @(
    @{ Name = 'bold';      Pattern = 'font-weight\s*:\s*(?:bold(?:er)?|[6-9]00)';  Open = '<strong>'; Close = '</strong>' }
    @{ Name = 'italic';    Pattern = 'font-style\s*:\s*(?:italic|oblique)';        Open = '<em>';     Close = '</em>' }
    @{ Name = 'underline'; Pattern = 'text-decoration[^;}]*\bunderline\b';         Open = '<u>';      Close = '</u>' }
    @{ Name = 'super';     Pattern = 'vertical-align\s*:\s*super';                 Open = '<sup>';    Close = '</sup>' }
    @{ Name = 'sub';       Pattern = 'vertical-align\s*:\s*sub';                   Open = '<sub>';    Close = '</sub>' }
)

# ---------------------------------------------------------------------------
# Process each file
# ---------------------------------------------------------------------------

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $html = [System.IO.File]::ReadAllText($target)

    # 1. Pull the CSS out of every <style> block. Rules outside a stylesheet
    #    are ignored on purpose - a bare "{...}" in body text is not CSS.
    $css = New-Object System.Text.StringBuilder
    foreach ($m in [regex]::Matches($html, '(?is)<style\b[^>]*>(.*?)</style>')) {
        [void]$css.AppendLine($m.Groups[1].Value)
    }
    if ($css.Length -eq 0) {
        Write-Host "  No <style> block; nothing to read formatting from. Skipping." -ForegroundColor DarkGray
        continue
    }

    # 2. Map class name -> the flags its declarations carry. A selector list
    #    (".a,.b{...}") applies the same block to every class in it.
    $classFlags = @{}
    foreach ($rule in [regex]::Matches($css.ToString(), '(?s)([^{}]+)\{([^{}]*)\}')) {
        $selector = $rule.Groups[1].Value
        $decls    = $rule.Groups[2].Value
        $hits = @()
        foreach ($f in $FLAGS) {
            if ($decls -match $f.Pattern) { $hits += $f.Name }
        }
        if ($hits.Count -eq 0) { continue }
        foreach ($sel in [regex]::Matches($selector, '\.([A-Za-z_][A-Za-z0-9_-]*)')) {
            $name = $sel.Groups[1].Value
            if (-not $classFlags.ContainsKey($name)) { $classFlags[$name] = @() }
            $classFlags[$name] = @($classFlags[$name] + $hits | Select-Object -Unique)
        }
    }

    if ($classFlags.Count -eq 0) {
        Write-Host "  No classes carry bold/italic/underline/super/sub. Nothing to do." -ForegroundColor DarkGray
        continue
    }

    # 3. Rewrite spans. ONE pass: a class that is both bold and superscript
    #    emits both tags in this single substitution. Doing a pass per tag
    #    would consume the span on the first pass and lose the rest.
    $spanCount = 0
    $tagCounts = @{}
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $classAttr = $m.Groups[1].Value
        $inner     = $m.Groups[2].Value

        $applied = @()
        foreach ($cls in ($classAttr -split '\s+')) {
            if ($cls -and $classFlags.ContainsKey($cls)) { $applied += $classFlags[$cls] }
        }
        if ($applied.Count -eq 0) { return $m.Value }   # plain span, leave alone
        $applied = @($applied | Select-Object -Unique)

        $open = ''; $close = ''
        foreach ($f in $FLAGS) {
            if ($applied -contains $f.Name) {
                $open  = $open + $f.Open
                $close = $f.Close + $close
                if (-not $tagCounts.ContainsKey($f.Name)) { $tagCounts[$f.Name] = 0 }
                $tagCounts[$f.Name]++
            }
        }
        $script:spanCount++
        return $open + $inner + $close
    }

    # flextable spans do not nest, so a non-greedy body is safe. The guard
    # against a nested <span> keeps this honest on other generators.
    $spanPattern = '(?s)<span\s+[^>]*class="([^"]*)"[^>]*>((?:(?!<span\b|</span>).)*)</span>'

    if ($ListOnly) {
        Write-Host "  Formatting classes found: $($classFlags.Count)" -ForegroundColor Cyan
        $preview = @{}
        foreach ($m in [regex]::Matches($html, $spanPattern)) {
            foreach ($cls in ($m.Groups[1].Value -split '\s+')) {
                if ($cls -and $classFlags.ContainsKey($cls)) {
                    if (-not $preview.ContainsKey($cls)) { $preview[$cls] = 0 }
                    $preview[$cls]++
                }
            }
        }
        foreach ($k in ($preview.Keys | Sort-Object)) {
            Write-Host ("    {0,-24} {1,5} span(s)   {2}" -f $k, $preview[$k], ($classFlags[$k] -join ', '))
        }
        $unused = $classFlags.Keys | Where-Object { -not $preview.ContainsKey($_) }
        if ($unused) { Write-Host "    ($($unused.Count) formatting class(es) defined but never used)" -ForegroundColor DarkGray }
        continue
    }

    $out = [System.Text.RegularExpressions.Regex]::Replace($html, $spanPattern, $evaluator)

    if ($spanCount -eq 0) {
        Write-Host "  No formatted spans matched. File unchanged." -ForegroundColor DarkGray
        continue
    }

    foreach ($k in ($tagCounts.Keys | Sort-Object)) {
        Write-Host ("  {0,-10} {1,5}" -f $k, $tagCounts[$k]) -ForegroundColor DarkGray
    }
    Write-Host "  Rewrote $spanCount span(s)" -ForegroundColor DarkGray

    $backup = "$target.inline.bak"
    Copy-Item -LiteralPath $target -Destination $backup -Force
    [System.IO.File]::WriteAllText($target, $out, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "  Saved (backup: $(Split-Path -Leaf $backup))" -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
