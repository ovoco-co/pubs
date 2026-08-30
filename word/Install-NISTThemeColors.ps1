<#
.SYNOPSIS
    Installs the NIST NCSTAR theme color file into the user's Word theme
    colors folder so the Design > Colors gallery displays "NIST NCSTAR"
    instead of "Custom 1" / "Custom N" for documents using the palette.

.DESCRIPTION
    Thin wrapper around publishing/word/Install-ThemeColors.ps1 that supplies
    the canonical NIST NCSTAR palette and scheme name.

    The file is per-user, not per-document. It does not affect the .docx
    contents, only how Word renders the picker.

.PARAMETER Force
    Overwrite an existing NIST NCSTAR.xml without prompting.

.EXAMPLE
    .\Install-NISTThemeColors.ps1

.EXAMPLE
    .\Install-NISTThemeColors.ps1 -Force
#>

[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# Canonical NIST NCSTAR clrScheme as a standalone Theme Colors file
$themeColorsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:clrScheme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="NIST NCSTAR">
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

$genericScript = Join-Path $PSScriptRoot 'Install-ThemeColors.ps1'

& $genericScript -SchemeName 'NIST NCSTAR' -ColorSchemeXml $themeColorsXml -Force:$Force
