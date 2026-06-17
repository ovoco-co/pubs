function Edit-RunRPrInStyle {
    <#
    .SYNOPSIS
        Edits font family, size, and bold in the run-level rPr of a style block.

    .DESCRIPTION
        Modifies only the run-level rPr (the one governing actual text), not the
        paragraph-mark rPr nested inside pPr. Sets rFonts, explicitly disables
        bold, and sets sz/szCs to the specified half-point value.

    .PARAMETER Block
        The style block XML string (a single w:style element).

    .PARAMETER NewFont
        The font family name to set (e.g., 'Source Sans Pro Light').

    .PARAMETER SizeHalfPts
        The font size in half-points (e.g., '18' for 9pt, '20' for 10pt).
    #>
    param(
        [string]$Block,
        [string]$NewFont,
        [string]$SizeHalfPts
    )
    $pPrCloseIdx = $Block.IndexOf('</w:pPr>')
    if ($pPrCloseIdx -ge 0) {
        $cut    = $pPrCloseIdx + '</w:pPr>'.Length
        $prefix = $Block.Substring(0, $cut)
        $tail   = $Block.Substring($cut)
    } else {
        $prefix = ''
        $tail   = $Block
    }

    $tail = [regex]::Replace($tail, '<w:b\s*/>', '')
    $tail = [regex]::Replace($tail, '<w:bCs\s*/>', '')
    $tail = [regex]::Replace($tail, '<w:b\s+w:val="[^"]*"\s*/>', '')
    $tail = [regex]::Replace($tail, '<w:bCs\s+w:val="[^"]*"\s*/>', '')

    $rFontsAndBold = "<w:rFonts w:ascii=`"$NewFont`" w:hAnsi=`"$NewFont`" w:cs=`"$NewFont`"/><w:b w:val=`"0`"/><w:bCs w:val=`"0`"/>"
    if ($tail -match '<w:rFonts\s+[^/]*?/>') {
        $tail = [regex]::Replace($tail, '<w:rFonts\s+[^/]*?/>', $rFontsAndBold, 1)
    } elseif ($tail -match '<w:rPr>') {
        $tail = [regex]::Replace($tail, '<w:rPr>', "<w:rPr>$rFontsAndBold", 1)
    } else {
        $tail = [regex]::Replace($tail, '</w:style>', "<w:rPr>$rFontsAndBold</w:rPr></w:style>", 1)
    }

    if ($tail -match '<w:sz\s+w:val="[^"]*"\s*/>') {
        $tail = [regex]::Replace($tail, '<w:sz\s+w:val="[^"]*"\s*/>', "<w:sz w:val=`"$SizeHalfPts`"/>", 1)
    } elseif ($tail -match '<w:rPr>') {
        $tail = [regex]::Replace($tail, '<w:rPr>', "<w:rPr><w:sz w:val=`"$SizeHalfPts`"/>", 1)
    }
    if ($tail -match '<w:szCs\s+w:val="[^"]*"\s*/>') {
        $tail = [regex]::Replace($tail, '<w:szCs\s+w:val="[^"]*"\s*/>', "<w:szCs w:val=`"$SizeHalfPts`"/>", 1)
    } elseif ($tail -match "<w:sz\s+w:val=`"$SizeHalfPts`"\s*/>") {
        $tail = [regex]::Replace($tail, "(<w:sz\s+w:val=`"$SizeHalfPts`"\s*/>)", "`$1<w:szCs w:val=`"$SizeHalfPts`"/>", 1)
    }

    return $prefix + $tail
}
