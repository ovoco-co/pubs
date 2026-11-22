function Set-BulletGlyph {
    <#
    .SYNOPSIS
        Overrides bullet glyph (lvlText and rFonts) in abstractNum ilvl=0.

    .DESCRIPTION
        For each style in OverrideMap that has a numId in styles.xml, resolves
        the linked abstractNumId in numbering.xml and replaces the w:lvlText
        and w:rFonts in ilvl=0 with the specified glyph character and font.

        Three guards decide whether a resolved abstractNum may be written to.
        Without them this function will overwrite ANY ilvl=0 it is pointed at,
        and a list style whose numId points somewhere unexpected takes an
        unrelated list definition down with it. In NWIRP chapter 5 ListBullet3
        pointed at HMList's num, so the Wingdings "n" glyph was written into
        the chapter-numbering definition: level 1 of HMList became a bullet,
        lost its <w:pStyle w:val="Heading1"/>, and Heading 1 stopped
        numbering. The guards skip an abstractNum that

          - carries w:styleLink or w:numStyleLink: it is a SHARED list
            definition owned by a named list style, not this style's private
            bullet list
          - declares w:multiLevelType multilevel: a bullet override belongs to
            a singleLevel (or hybridMultilevel) bullet list
          - has an ilvl=0 whose w:numFmt is not "bullet": there is no bullet
            glyph to override on a numbered level

        A skipped style is left exactly as found rather than guessed at.

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
        # Never assume an attribute's position. Word writes
        # <w:num w16cid:durableId="620185856" w:numId="1"> and
        # <w:abstractNum w15:restartNumberingAfterBreak="0" w:abstractNumId="0">,
        # so a pattern anchored on the first attribute matches almost nothing.
        $numIdMatch = [regex]::Match($sm.Value, '<w:numId\b[^>]*\bw:val="(\d+)"')
        if (-not $numIdMatch.Success) { continue }
        $nid = $numIdMatch.Groups[1].Value
        if ($nid -eq '0') { continue }   # numId 0 means "no numbering"
        $am = [regex]::Match($NumberingXml, '(?s)<w:num\b[^>]*\bw:numId="' + [regex]::Escape($nid) + '"[^>]*>.*?<w:abstractNumId\b[^>]*\bw:val="(\d+)"')
        if (-not $am.Success) { continue }
        $absid = $am.Groups[1].Value
        $abnPattern = '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="' + [regex]::Escape($absid) + '"[^>]*>.*?</w:abstractNum>'
        $abMatch = [regex]::Match($NumberingXml, $abnPattern)
        if (-not $abMatch.Success) { continue }
        $abBlock = $abMatch.Value

        # Guard 1: a shared list definition belongs to a named list style.
        if ($abBlock -match '<w:(?:styleLink|numStyleLink)\b') { continue }
        # Guard 2: a bullet override is not meant for a multilevel list.
        $mlt = [regex]::Match($abBlock, '<w:multiLevelType\b[^>]*\bw:val="([^"]+)"')
        if ($mlt.Success -and $mlt.Groups[1].Value -eq 'multilevel') { continue }
        # Guard 3: only override a level that actually carries a bullet.
        $lvl0 = [regex]::Match($abBlock, '(?s)<w:lvl\b[^>]*\bw:ilvl="0"[^>]*>.*?</w:lvl>')
        if (-not $lvl0.Success) { continue }
        $fmt0 = [regex]::Match($lvl0.Value, '<w:numFmt\b[^>]*\bw:val="([^"]+)"')
        if (-not $fmt0.Success -or $fmt0.Groups[1].Value -ne 'bullet') { continue }

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
        # Instance method, because the static Regex.Replace has no overload
        # taking a replacement count: a 4th argument of 1 binds to
        # RegexOptions.IgnoreCase and replaces every match.
        $newAb = ([regex]'(?s)(<w:lvl\b[^>]*\bw:ilvl="0"[^>]*>)(.*?)(</w:lvl>)').Replace(
            $abBlock, $lvlEvaluator, 1)
        $NumberingXml = $NumberingXml.Substring(0, $abMatch.Index) + $newAb + $NumberingXml.Substring($abMatch.Index + $abMatch.Length)
    }
    return $NumberingXml
}
