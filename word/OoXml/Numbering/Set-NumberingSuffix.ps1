function Set-NumberingSuffix {
    <#
    .SYNOPSIS
        Normalizes the w:suff element on every level in every abstractNum.

    .DESCRIPTION
        Sets w:suff to 'space' for heading-linked abstractNums (so chapter
        numbers sit tight with titles) and 'tab' for all others (so bullet
        glyphs tab out to the paragraph indent).

    .PARAMETER NumberingXml
        The full numbering.xml content as a string.

    .PARAMETER HeadingAbstractNumIds
        A HashSet of abstractNumId strings that are heading-linked
        (from Get-HeadingAbstractNumIds).

    .PARAMETER HeadingSuff
        The suff value for heading abstractNums. Default: 'space'.

    .PARAMETER ListSuff
        The suff value for non-heading abstractNums. Default: 'tab'.
    #>
    param(
        [string]$NumberingXml,
        [System.Collections.Generic.HashSet[string]]$HeadingAbstractNumIds,
        [string]$HeadingSuff = 'space',
        [string]$ListSuff = 'tab'
    )
    # A null set means no heading is numbered, which is a normal document, not
    # an error. Without this the evaluator below calls .Contains on null, and
    # because it runs inside Regex.Replace the failure surfaces as the useless
    # 'Exception calling "Replace" with "3" argument(s)' with no styleId in it.
    if ($null -eq $HeadingAbstractNumIds) {
        $HeadingAbstractNumIds = [System.Collections.Generic.HashSet[string]]::new()
    }

    $abstractNumEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($am)
        $whole = $am.Value
        $absid = $am.Groups[1].Value
        $useSuff = if ($HeadingAbstractNumIds.Contains($absid)) { $HeadingSuff } else { $ListSuff }
        $newSuff = "<w:suff w:val=`"$useSuff`"/>"
        $lvlEval = [System.Text.RegularExpressions.MatchEvaluator]{
            param($lm)
            $lvlBlock = $lm.Value
            $inner    = $lm.Groups[1].Value
            if ($inner -notmatch '<w:lvlText\b') { return $lvlBlock }
            # CT_Lvl is a SEQUENCE:
            #   start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText,
            #   lvlPicBulletId, legacy, lvlJc, pPr, rPr
            # so w:suff belongs immediately BEFORE w:lvlText - after pStyle,
            # not after numFmt. Anchoring the insertion on numFmt or start put
            # suff ahead of lvlRestart and pStyle and made the file
            # schema-invalid. Word does not prompt to repair that: it drops the
            # out-of-order child, and losing <w:pStyle w:val="Heading1"/> is
            # what stops Heading 1 numbering. It reached every heading-linked
            # abstractNum in the NWIRP chapters, HMList ilvl 0 included.
            #
            # Any existing suff is REMOVED first rather than replaced in place,
            # so a file already written with it in the wrong position is
            # repaired by re-running instead of keeping the bad order.
            $lvlBlock = [regex]::Replace($lvlBlock, '<w:suff\b[^>]*/>', '')
            return ([regex]'(<w:lvlText\b)').Replace($lvlBlock, ($newSuff + '$1'), 1)
        }.GetNewClosure()
        return [System.Text.RegularExpressions.Regex]::Replace($whole, '(?s)<w:lvl\b[^>]*>(.*?)</w:lvl>', $lvlEval)
    }.GetNewClosure()
    return [System.Text.RegularExpressions.Regex]::Replace(
        $NumberingXml,
        '(?s)<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"[^>]*>.*?</w:abstractNum>',
        $abstractNumEvaluator)
}
