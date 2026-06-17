function Remove-StyleBlock {
    <#
    .SYNOPSIS
        Deletes a w:style element by styleId from styles.xml.

    .DESCRIPTION
        Removes the entire w:style element matching the given styleId.
        Returns a two-element array: (modified XML string, [bool] changed).

    .PARAMETER Xml
        The full styles.xml content as a string.

    .PARAMETER Id
        The w:styleId value to remove.
    #>
    param(
        [string]$Xml,
        [string]$Id
    )
    $pattern = '(?si)<w:style\b[^>]*?\bw:styleId="' + [regex]::Escape($Id) + '"[^>]*?>.*?</w:style>\s*'
    $rgx = [regex]::new($pattern)
    if (-not $rgx.IsMatch($Xml)) { return ,$Xml, $false }
    return $rgx.Replace($Xml, '', 1), $true
}
