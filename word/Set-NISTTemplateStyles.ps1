<#
.SYNOPSIS
    Applies the NIST style architecture to a CLEAN document - the template file
    or a new document built from it - without the legacy-fix passes.

.DESCRIPTION
    This is the go-forward styler for the clean-template workflow. It is a thin
    wrapper over Set-NISTStyleArchitecture.ps1 -Clean, so there is ONE source of
    truth for the architecture. It applies the canonical style definitions,
    colors, indents, numbering, ligatures-off, and recommended-style curation,
    but skips the rescue passes that only matter for messy inherited documents
    (the Heading 1/3 cs="Times New Roman" strip and the junk-style deletions).

    Use Set-NISTStyleArchitecture.ps1 (without -Clean) for legacy documents that
    still carry old cruft to clean up.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. Applies to every Word file in it. Either this or -File.

.PARAMETER Force
    Skip prompts.

.EXAMPLE
    .\Set-NISTTemplateStyles.ps1 -File "HM-NCSTAR-Template.docx"

.EXAMPLE
    .\Set-NISTTemplateStyles.ps1 -Folder "C:\Users\Artist\Documents\260806" -Force
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

$engine = Join-Path $PSScriptRoot 'Set-NISTStyleArchitecture.ps1'
if (-not (Test-Path -LiteralPath $engine)) { throw "Engine not found: $engine" }

$fwd = @{ Clean = $true }
if ($PSCmdlet.ParameterSetName -eq 'File') { $fwd['File'] = $File } else { $fwd['Folder'] = $Folder }
if ($Force) { $fwd['Force'] = $true }

& $engine @fwd
