function Edit-StyleBlock {
    <#
    .SYNOPSIS
        Finds a w:style element by styleId and applies a mutator scriptblock.

    .DESCRIPTION
        Searches the styles.xml string for a w:style element matching the given
        styleId, passes the full element text to the Mutator scriptblock, and
        replaces the original with the mutator's return value.

        Returns a two-element array: (modified XML string, [bool] changed).

    .PARAMETER Xml
        The full styles.xml content as a string.

    .PARAMETER Id
        The w:styleId value to find.

    .PARAMETER Mutator
        A scriptblock that receives the matched style block and returns the
        replacement. Called as: & $Mutator $fullBlock $attributes $innerBody
    #>
    param(
        [string]$Xml,
        [string]$Id,
        [scriptblock]$Mutator
    )
    $pattern = '(?si)<w:style\b([^>]*?\bw:styleId="' + [regex]::Escape($Id) + '"[^>]*?)>(.*?)</w:style>'
    $rgx = [regex]::new($pattern)
    $m = $rgx.Match($Xml)
    # No leading comma - see Remove-StyleBlock. This path runs whenever the
    # document does not define the style being edited, which is common.
    if (-not $m.Success) { return $Xml, $false }
    $newBlock = & $Mutator $m.Value $m.Groups[1].Value $m.Groups[2].Value
    return $rgx.Replace($Xml, [System.Text.RegularExpressions.MatchEvaluator]{ param($mm) $newBlock }, 1), $true
}
