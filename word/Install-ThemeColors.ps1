<#
.SYNOPSIS
    Installs a theme color file into the user's Word theme colors folder so
    the Design > Colors gallery displays the scheme by name.

.DESCRIPTION
    Word's Design > Colors gallery enumerates *.xml files in
    %APPDATA%\Microsoft\Templates\Document Themes\Theme Colors\. Each file
    is a top-level <a:clrScheme> with a name and twelve colors. When a
    document's embedded clrScheme matches one of these files, Word labels
    the current selection with that file's name; otherwise it falls back
    to a generic "Custom N" label.

    This script writes "<SchemeName>.xml" into that folder. After one
    install, every document with a matching palette will display the scheme
    name in the gallery on this machine.

    The file is per-user, not per-document. It does not affect the .docx
    contents, only how Word renders the picker.

.PARAMETER SchemeName
    Name for the theme color scheme. Used as both the XML file name and
    the display name in Word's Design > Colors gallery.

.PARAMETER ColorSchemeXml
    Complete <a:clrScheme> XML block with all 12 theme color elements.
    Must include the xmlns:a namespace declaration.

.PARAMETER Force
    Overwrite an existing file without prompting.

.EXAMPLE
    .\Install-ThemeColors.ps1 -SchemeName "Corporate Blue" -ColorSchemeXml $myScheme

.EXAMPLE
    .\Install-ThemeColors.ps1 -SchemeName "Corporate Blue" -ColorSchemeXml $myScheme -Force
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SchemeName,

    [Parameter(Mandatory=$true)]
    [string]$ColorSchemeXml,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$appData = [Environment]::GetFolderPath('ApplicationData')
$themeColorsDir = Join-Path $appData 'Microsoft\Templates\Document Themes\Theme Colors'
$targetPath = Join-Path $themeColorsDir "$SchemeName.xml"

Write-Host "Target: $targetPath" -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $themeColorsDir)) {
    Write-Host "Creating $themeColorsDir..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $themeColorsDir -Force | Out-Null
}

if ((Test-Path -LiteralPath $targetPath) -and -not $Force) {
    Write-Host ""
    $reply = Read-Host "File exists. Overwrite? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($targetPath, $ColorSchemeXml, $utf8NoBom)

Write-Host ""
Write-Host "Installed $SchemeName.xml." -ForegroundColor Green
Write-Host "Restart Word for the picker to pick it up." -ForegroundColor Cyan
Write-Host "After that, Design > Colors will show '$SchemeName' as a selectable scheme" -ForegroundColor Cyan
Write-Host "and any document with matching accents will display '$SchemeName' instead of 'Custom N'." -ForegroundColor Cyan
