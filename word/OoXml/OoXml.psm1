# OoXml Module - Generic OOXML style and numbering manipulation
# Dot-source all function files from Styles/ and Numbering/ subdirectories

$modulePath = $PSScriptRoot

foreach ($subdir in @('Styles', 'Numbering')) {
    $dirPath = Join-Path $modulePath $subdir
    if (Test-Path $dirPath) {
        Get-ChildItem -Path $dirPath -Filter '*.ps1' | ForEach-Object {
            . $_.FullName
        }
    }
}

# Exported scriptblock variables for common style mutations

# Strip explicit spacing override and add contextualSpacing for tight list stacking
$ListSingleSpacingEdit = {
    param($block)
    $b = [regex]::Replace($block, '<w:spacing\b[^/]*?/>', '')
    if ($b -notmatch '<w:contextualSpacing\s*/>') {
        if ($b -match '</w:pPr>') {
            $b = [regex]::Replace($b, '</w:pPr>', '<w:contextualSpacing/></w:pPr>', 1)
        } else {
            $b = [regex]::Replace($b, '</w:style>', '<w:pPr><w:contextualSpacing/></w:pPr></w:style>', 1)
        }
    }
    return $b
}

# Strip any explicit spacing override so style inherits cleanly from its base
$BodyTextSpacingEdit = {
    param($block)
    return [regex]::Replace($block, '<w:spacing\b[^/]*?/>', '')
}

# Insert numId=0 to suppress inherited numbering
$SuppressNumbering = {
    param($block)
    $noNumPr = '<w:numPr><w:numId w:val="0"/></w:numPr>'
    if ($block -match '(?s)<w:numPr>.*?</w:numPr>') {
        return [regex]::Replace($block, '(?s)<w:numPr>.*?</w:numPr>', $noNumPr, 1)
    }
    if ($block -match '<w:pPr>') {
        return [regex]::Replace($block, '<w:pPr>', "<w:pPr>$noNumPr", 1)
    }
    if ($block -match '<w:basedOn\s+w:val="[^"]*"\s*/>') {
        return [regex]::Replace($block, '(<w:basedOn\s+w:val="[^"]*"\s*/>)', "`$1<w:pPr>$noNumPr</w:pPr>", 1)
    }
    return [regex]::Replace($block, '</w:style>', "<w:pPr>$noNumPr</w:pPr></w:style>", 1)
}

Export-ModuleMember -Function @(
    'Edit-StyleBlock',
    'Remove-StyleBlock',
    'Get-AllStyleIdsByType',
    'Set-StyleBasedOn',
    'Set-StyleColor',
    'Edit-RunRPrInStyle',
    'Set-StyleRecommended',
    'Set-StyleHidden',
    'Get-HeadingAbstractNumIds',
    'Set-NumberingSuffix',
    'Sync-ListIndents',
    'Set-BulletGlyph'
) -Variable @(
    'ListSingleSpacingEdit',
    'BodyTextSpacingEdit',
    'SuppressNumbering'
)
