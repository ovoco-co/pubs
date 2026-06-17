function Set-BulletGlyph {
    <#
    .SYNOPSIS
        Overrides bullet glyph (lvlText and rFonts) in abstractNum ilvl=0.

    .DESCRIPTION
        For each style in OverrideMap that has a numId in styles.xml, resolves
        the linked abstractNumId in numbering.xml and replaces the w:lvlText
        and w:rFonts in ilvl=0 with the specified glyph character and font.

    .PARAMETER NumberingXml
        The full numbering.xml content as a string.

    .PARAMETER StylesXml
        The full styles.xml content as a string.

    .PARAMETER OverrideMap
        Hashtable mapping styleId to @{ LvlText = <char>; Font = <font name> }.
    #>
    param(
        [string]$NumberingXml,
        [string]$StylesXml,
        [hashtable]$OverrideMap
    )
    foreach ($lstId in $OverrideMap.Keys) {
        $stylePattern = '(?s)<w:style\b[^>]*?\bw:styleId="' + [regex]::Escape($lstId) + '"[^>]*?>.*?</w:style>'
        $sm = [regex]::Match($StylesXml, $stylePattern)
        if (-not $sm.Success) { continue }
        $numIdMatch = [regex]::Match($sm.Value, '<w:numId\s+w:val="(\d+)"\s*/>')
        if (-not $numIdMatch.Success) { continue }
        $nid = $numIdMatch.Groups[1].Value
        $am = [regex]::Match($NumberingXml, '<w:num w:numId="' + [regex]::Escape($nid) + '"[^>]*>\s*<w:abstractNumId w:val="(\d+)"', 'Singleline')
        if (-not $am.Success) { continue }
        $absid = $am.Groups[1].Value
        $abnPattern = '(?s)<w:abstractNum [^>]*abstractNumId="' + [regex]::Escape($absid) + '"[^>]*>.*?</w:abstractNum>'
        $abMatch = [regex]::Match($NumberingXml, $abnPattern)
        if (-not $abMatch.Success) { continue }
        $abBlock = $abMatch.Value
        $bulletOverride = $OverrideMap[$lstId]
        $lvlEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
            param($m)
            $head = $m.Groups[1].Value; $body = $m.Groups[2].Value; $tail = $m.Groups[3].Value
            $newLvlText = "<w:lvlText w:val=`"$($bulletOverride.LvlText)`"/>"
            if ($body -match '<w:lvlText\s+[^/]*?/>') {
                $body = [regex]::Replace($body, '<w:lvlText\s+[^/]*?/>', $newLvlText, 1)
            }
            $newRFonts = "<w:rFonts w:ascii=`"$($bulletOverride.Font)`" w:hAnsi=`"$($bulletOverride.Font)`" w:cs=`"$($bulletOverride.Font)`" w:hint=`"default`"/>"
            if ($body -match '<w:rPr\b[^>]*>.*?<w:rFonts\s+[^/]*?/>.*?</w:rPr>') {
                $body = [regex]::Replace($body, '<w:rFonts\s+[^/]*?/>', $newRFonts, 1)
            } elseif ($body -match '<w:rPr\b[^>]*>') {
                $body = [regex]::Replace($body, '(<w:rPr\b[^>]*>)', "`$1$newRFonts", 1)
            } else {
                $body = $body + "<w:rPr>$newRFonts</w:rPr>"
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
