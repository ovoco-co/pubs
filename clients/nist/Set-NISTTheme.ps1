<#
.SYNOPSIS
    Replaces the color scheme in a Word file's theme1.xml with the
    canonical NIST NCSTAR / NWIRP entity palette.

.DESCRIPTION
    Thin wrapper around publishing/word/Set-ThemeColors.ps1 that supplies the
    canonical NIST NCSTAR palette (Resolution Blue, Tawny, Shamrock Green,
    Claret, Picton Blue, Mulberry).

    The canonical palette comes from NWIRP_2-3_Graphics_Recommendations.md.
    Once applied, figures and styles can reference theme slots by name
    (accent1 = Study/NIST, accent2 = Schools, etc.) instead of hardcoding
    hex values, so the colors cannot drift between chapter documents.

    Resolution Blue is set as #00257A (the master-palette value from
    the graphics doc, canonical). The variant #02257A that appears
    elsewhere in NCSTAR styles is treated as a typo and is NOT preserved.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Path to a folder. The script applies to every supported Word file
    in the folder. Either this or -File is required.

.PARAMETER Force
    Overwrite the file without prompting and skip the existing-backup check.

.EXAMPLE
    .\Set-NISTTheme.ps1 -File "C:\templates\HM-NWIRP-Template.dotx"

.EXAMPLE
    .\Set-NISTTheme.ps1 -Folder "\\blue\E\00-clients\nist-HM\Templates" -Force
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

# Canonical NIST NCSTAR color scheme
$nistClrScheme = @'
<a:clrScheme name="NIST NCSTAR">
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="355195"/></a:dk2>
    <a:lt2><a:srgbClr val="CCD3E4"/></a:lt2>
    <a:accent1><a:srgbClr val="00257A"/></a:accent1>
    <a:accent2><a:srgbClr val="D55E00"/></a:accent2>
    <a:accent3><a:srgbClr val="009E73"/></a:accent3>
    <a:accent4><a:srgbClr val="990B38"/></a:accent4>
    <a:accent5><a:srgbClr val="56B4E9"/></a:accent5>
    <a:accent6><a:srgbClr val="C15C94"/></a:accent6>
    <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
    <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
'@

Write-Host @"
NIST NCSTAR theme palette:
  dk1       sys windowText (#000000)
  lt1       sys window (#FFFFFF)
  dk2       #355195  NCSTAR Secondary (secondary text)
  lt2       #CCD3E4  NCSTAR header background
  accent1   #00257A  Resolution Blue   - Study / NIST elements
  accent2   #D55E00  Tawny             - Schools
  accent3   #009E73  Shamrock Green    - Hospitals
  accent4   #990B38  Claret            - Disruptions
  accent5   #56B4E9  Picton Blue       - Secondary emphasis, water/utilities
  accent6   #C15C94  Mulberry          - Additional category
  hlink     #0563C1  Hyperlink
  folHlink  #954F72  Followed Hyperlink
"@ -ForegroundColor Cyan

$genericScript = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'word\Set-ThemeColors.ps1'

$args_ = @{
    ColorSchemeXml = $nistClrScheme
    Force          = $Force
}

if ($PSCmdlet.ParameterSetName -eq 'File') {
    $args_.File = $File
} else {
    $args_.Folder = $Folder
}

& $genericScript @args_
