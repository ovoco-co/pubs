<#
.SYNOPSIS
    Runs the full Word-to-DITA conversion pipeline.

.DESCRIPTION
    Orchestrates all four conversion scripts in sequence: content conversion,
    numbering extraction, header/footer extraction, and table/image extraction.
    Produces a complete DITA output directory with topic files, ditamap,
    numbering metadata, header/footer data, and extracted assets.

.PARAMETER InputDocx
    Path to the source .docx file.

.PARAMETER WorkingDir
    Root output directory. Subdirectories (dita/, assets/) are created
    automatically. Created if it does not exist.

.EXAMPLE
    .\Run-All.ps1 -InputDocx "chapter.docx" -WorkingDir "output"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$WorkingDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$styleMap = Join-Path $scriptRoot "stylemap.json"
$contentScript = Join-Path $scriptRoot "Convert-WordToDita.ps1"
$numberingScript = Join-Path $scriptRoot "Extract-Numbering.ps1"
$hfScript = Join-Path $scriptRoot "Extract-HeadersFooters.ps1"
$tablesScript = Join-Path $scriptRoot "Extract-TablesImages.ps1"

if (-not (Test-Path $WorkingDir)) {
    New-Item -ItemType Directory -Path $WorkingDir | Out-Null
}

& $contentScript -InputDocx $InputDocx -OutputDir (Join-Path $WorkingDir "dita") -StyleMapJson $styleMap
& $numberingScript -InputDocx $InputDocx -OutputJson (Join-Path $WorkingDir "numbering.json")
& $hfScript -InputDocx $InputDocx -OutputJson (Join-Path $WorkingDir "headers-footers.json")
& $tablesScript -InputDocx $InputDocx -OutputDir (Join-Path $WorkingDir "assets")

Write-Host "All processing complete."