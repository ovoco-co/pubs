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
            if ($lvlBlock -match '<w:suff\s+w:val="[^"]*"\s*/>') {
                return [regex]::Replace($lvlBlock, '<w:suff\s+w:val="[^"]*"\s*/>', $newSuff, 1)
            }
            if ($lvlBlock -match '<w:numFmt\s+[^/]*/>') {
                return [regex]::Replace($lvlBlock, '(<w:numFmt\s+[^/]*/>)', "`$1$newSuff", 1)
            }
            if ($lvlBlock -match '<w:start\s+[^/]*/>') {
                return [regex]::Replace($lvlBlock, '(<w:start\s+[^/]*/>)', "`$1$newSuff", 1)
            }
            return [regex]::Replace($lvlBlock, '(<w:lvl\b[^>]*>)', "`$1$newSuff", 1)
        }.GetNewClosure()
        return [System.Text.RegularExpressions.Regex]::Replace($whole, '(?s)<w:lvl\b[^>]*>(.*?)</w:lvl>', $lvlEval)
    }.GetNewClosure()
    return [System.Text.RegularExpressions.Regex]::Replace(
        $NumberingXml,
        '(?s)<w:abstractNum [^>]*abstractNumId="(\d+)"[^>]*>.*?</w:abstractNum>',
        $abstractNumEvaluator)
}
