function Set-StyleColor {
    <#
    .SYNOPSIS
        Sets w:color on a style's rPr with optional theme color mapping.

    .DESCRIPTION
        Finds or creates a w:rPr element in the style block and sets or replaces
        the w:color element. If HexToThemeMap is provided and the hex value
        matches a key, a w:themeColor attribute is also added so the style
        follows the document theme.

    .PARAMETER Block
        The style block XML string (a single w:style element).

    .PARAMETER HexNoHash
        The color hex value without the # prefix (e.g., '00257A').

    .PARAMETER HexToThemeMap
        Optional hashtable mapping hex values to theme color slot names
        (e.g., @{ '00257A' = 'accent1' }).
    #>
    param(
        [string]$Block,
        [string]$HexNoHash,
        [hashtable]$HexToThemeMap = @{}
    )
    $themeAttr = ''
    if ($HexToThemeMap.ContainsKey($HexNoHash)) {
        $themeAttr = " w:themeColor=`"$($HexToThemeMap[$HexNoHash])`""
    }
    $colorEl = "<w:color w:val=`"$HexNoHash`"$themeAttr/>"
    if ($Block -match '<w:rPr>(.*?)</w:rPr>') {
        $rprInner = $Matches[1]
        if ($rprInner -match '<w:color\s+[^/]*?/>') {
            $newRprInner = [regex]::Replace($rprInner, '<w:color\s+[^/]*?/>', $colorEl)
        } else {
            $newRprInner = $rprInner + $colorEl
        }
        return [regex]::Replace($Block, '<w:rPr>(.*?)</w:rPr>', "<w:rPr>$newRprInner</w:rPr>", 1)
    } else {
        return [regex]::Replace($Block, '</w:style>', "<w:rPr>$colorEl</w:rPr></w:style>", 1)
    }
}
