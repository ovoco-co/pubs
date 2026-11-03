function Sync-ListIndents {
    <#
    .SYNOPSIS
        Propagates style indents into abstractNum ilvl=0 definitions.

    .DESCRIPTION
        For each style in IndentMap that has a numId in styles.xml, resolves the
        linked abstractNumId in numbering.xml and rewrites the w:ind element in
        ilvl=0 to match the style's indent values, and moves the level's "num"
        tab stop to the same text position (Left). This keeps the abstractNum
        indent AND its tab stop in agreement with the style so the first line's
        text lines up with the wrapped lines and Word renders consistently.

        Only processes styles with a Hanging value greater than 0 (bullet/number
        styles with a hanging indent marker).

    .PARAMETER NumberingXml
        The full numbering.xml content as a string.

    .PARAMETER StylesXml
        The full styles.xml content as a string.

    .PARAMETER IndentMap
        Ordered hashtable mapping styleId to @{ Left = <twips>; Hanging = <twips> }.
    #>
    param(
        [string]$NumberingXml,
        [string]$StylesXml,
        [System.Collections.Specialized.OrderedDictionary]$IndentMap
    )
    foreach ($lstId in $IndentMap.Keys) {
        $stylePattern = '(?s)<w:style\b[^>]*?\bw:styleId="' + [regex]::Escape($lstId) + '"[^>]*?>.*?</w:style>'
        $sm = [regex]::Match($StylesXml, $stylePattern)
        if (-not $sm.Success) { continue }
        $numIdMatch = [regex]::Match($sm.Value, '<w:numId\s+w:val="(\d+)"\s*/>')
        if (-not $numIdMatch.Success) { continue }
        $nid = $numIdMatch.Groups[1].Value
        $left = $IndentMap[$lstId].Left
        $hang = $IndentMap[$lstId].Hanging
        if ($hang -le 0) { continue }
        $am = [regex]::Match($NumberingXml, '<w:num w:numId="' + [regex]::Escape($nid) + '"[^>]*>\s*<w:abstractNumId w:val="(\d+)"', 'Singleline')
        if (-not $am.Success) { continue }
        $absid = $am.Groups[1].Value
        $abnPattern = '(?s)<w:abstractNum [^>]*abstractNumId="' + [regex]::Escape($absid) + '"[^>]*>.*?</w:abstractNum>'
        $abMatch = [regex]::Match($NumberingXml, $abnPattern)
        if (-not $abMatch.Success) { continue }
        $abBlock = $abMatch.Value
        $newInd = "<w:ind w:left=`"$left`" w:hanging=`"$hang`"/>"
        # The numbering "num" tab stop is where the first line's text lands after
        # the bullet/number, so it must sit at the text position (Left). Rebuilt
        # to <w:tab w:val="num" w:pos="Left"/>; if it is left at the old pos the
        # first line does not line up with the wrapped lines or the new indent.
        $numTab = '<w:tab w:val="num" w:pos="' + $left + '"/>'
        $lvlEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
            param($m)
            $head = $m.Groups[1].Value; $body = $m.Groups[2].Value; $tail = $m.Groups[3].Value
            if ($body -match '<w:ind\s+[^/]*?/>') {
                $body = [regex]::Replace($body, '<w:ind\s+[^/]*?/>', $newInd, 1)
            } elseif ($body -match '<w:pPr\b[^>]*>') {
                $body = [regex]::Replace($body, '(<w:pPr\b[^>]*>)', "`$1$newInd", 1)
            } elseif ($body -match '<w:pPr\s*/>') {
                $body = [regex]::Replace($body, '<w:pPr\s*/>', "<w:pPr>$newInd</w:pPr>", 1)
            } else {
                $body = $body + "<w:pPr>$newInd</w:pPr>"
            }
            # Move the "num" tab stop to match the text position.
            if ($body -match '<w:tab\b[^>]*?w:val="num"[^>]*?/>') {
                $body = [regex]::Replace($body, '<w:tab\b[^>]*?w:val="num"[^>]*?/>', $numTab)
            }
            return $head + $body + $tail
        }.GetNewClosure()
        $newAb = [System.Text.RegularExpressions.Regex]::Replace(
            $abBlock,
            '(?s)(<w:lvl w:ilvl="0"[^>]*>)(.*?)(</w:lvl>)',
            $lvlEvaluator,
            1)
        $NumberingXml = $NumberingXml.Substring(0, $abMatch.Index) + $newAb + $NumberingXml.Substring($abMatch.Index + $abMatch.Length)
    }
    return $NumberingXml
}
