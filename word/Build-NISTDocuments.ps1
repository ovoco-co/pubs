<#
.SYNOPSIS
    Runs the full NIST formatting pipeline on every Word file in a folder, in
    order: theme -> styles -> page layout -> table look -> 3-column widths ->
    typography -> de-embed fonts. One command to take a folder of documents to
    publication-ready formatting.

.DESCRIPTION
    Orchestrates the individual publishing/word scripts so they run in the right
    dependency order (theme must precede styles; layout/tables/content-cleanup
    follow styles; font de-embedding is last so nothing re-embeds). Each step is
    an existing script called with -Folder and -Force; this wrapper prompts once,
    then runs them all.

    The step list is defined in $steps at the top of the script - reorder,
    comment out, or add steps there. Three scripts are intentionally NOT in the
    default pipeline:
      - Set-NISTHeaderFooterContent.ps1 : needs a per-file -FrontMatter decision
        (chapter first pages vs. front matter), so run it separately.
      - Reapply-Styles.ps1 : strips direct character formatting, which can remove
        intentional manual emphasis; run it deliberately, not in a blanket pass.
      - Remove-Shading.ps1 : shading is often intentional. Several NIST styles
        define fills (NIST-Data, HM-Data, SidebarTable, Quote, BlockText,
        CaptionFigure2), and table cells carry deliberate fills that
        -IncludeTables would strip. Run it per file when you know a document has
        unwanted direct shading.

    Fails fast: if any step throws, the pipeline stops and reports which step and
    file failed, so a folder is never left half-processed silently. Every
    underlying script makes its own backup before writing.

.PARAMETER Folder
    Path to the folder of Word files to process.

.PARAMETER Clean
    Pass -Clean to the style step (Set-NISTStyleArchitecture), applying the
    architecture without the legacy-fix passes. Use for clean template-based
    documents; omit for messy legacy documents that need the full cleanup.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Build-NISTDocuments.ps1 -Folder "C:\Users\Artist\Documents\260806"

.EXAMPLE
    # Clean template-based documents (skip legacy fixes in the style step)
    .\Build-NISTDocuments.ps1 -Folder "C:\...\Templates" -Clean -Force
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Folder,

    [switch]$Clean,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
$Folder = (Resolve-Path -LiteralPath $Folder).ProviderPath

# Ordered pipeline. Each step: a script name + any step-specific args. All are
# called with -Folder <folder> -Force in addition to Args.
$steps = @(
    @{ Name = 'Theme';          Script = 'Set-NISTTheme.ps1';             Args = @{} }
    @{ Name = 'Styles';         Script = 'Set-NISTStyleArchitecture.ps1'; Args = @{ Clean = [bool]$Clean } }
    @{ Name = 'Page layout';    Script = 'Set-PageLayout.ps1';            Args = @{} }
    @{ Name = 'Table look';     Script = 'Set-TableLook.ps1';             Args = @{} }
    @{ Name = '3-column widths'; Script = 'Set-NISTThreeColumnTableWidths.ps1'; Args = @{} }
    @{ Name = 'Typography';     Script = 'Repair-Typography.ps1';         Args = @{} }
    @{ Name = 'De-embed fonts'; Script = 'Remove-EmbeddedFonts.ps1';      Args = @{} }
)

# Verify every step script exists before starting.
foreach ($s in $steps) {
    $p = Join-Path $PSScriptRoot $s.Script
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { throw "Pipeline script not found: $p" }
}

Write-Host "NIST document pipeline" -ForegroundColor Cyan
Write-Host "  Folder: $Folder" -ForegroundColor Cyan
Write-Host "  Style step: $(if ($Clean) { 'Clean (no legacy fixes)' } else { 'Full (legacy fixes on)' })" -ForegroundColor Cyan
Write-Host "  Steps in order:" -ForegroundColor Cyan
$i = 1
foreach ($s in $steps) { Write-Host ("    {0}. {1}  ({2})" -f $i, $s.Name, $s.Script); $i++ }
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Run all $($steps.Count) steps on this folder? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

$stepNum = 0
foreach ($s in $steps) {
    $stepNum++
    Write-Host ""
    Write-Host ("===== Step {0}/{1}: {2} =====" -f $stepNum, $steps.Count, $s.Name) -ForegroundColor Green
    $scriptPath = Join-Path $PSScriptRoot $s.Script
    $callArgs = @{ Folder = $Folder; Force = $true } + $s.Args
    try {
        & $scriptPath @callArgs
    } catch {
        Write-Host ""
        Write-Host ("PIPELINE STOPPED at step {0}/{1} ({2}, {3}):" -f $stepNum, $steps.Count, $s.Name, $s.Script) -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

Write-Host ""
Write-Host "Pipeline complete: all $($steps.Count) steps ran on $Folder." -ForegroundColor Green
