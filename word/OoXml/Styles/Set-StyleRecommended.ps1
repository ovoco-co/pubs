function Set-StyleRecommended {
    <#
    .SYNOPSIS
        Marks a style as recommended in Word's style pane.

    .DESCRIPTION
        Sets uiPriority to the given value, removes semiHidden and
        unhideWhenUsed, and adds qFormat so the style appears in the
        Recommended view of the Styles pane.

    .PARAMETER Xml
        The full styles.xml content as a string.

    .PARAMETER Id
        The w:styleId to mark as recommended.

    .PARAMETER Priority
        The uiPriority value (controls sort order in Recommended view).
    #>
    param(
        [string]$Xml,
        [string]$Id,
        [int]$Priority
    )
    $pat = '(?s)<w:style\b[^>]*?\bw:styleId="' + [regex]::Escape($Id) + '"[^>]*?>(.*?)</w:style>'
    $rg = [regex]::new($pat)
    $mm = $rg.Match($Xml)
    if (-not $mm.Success) { return $Xml }
    $blk = $mm.Value
    $blk = [regex]::Replace($blk, '<w:semiHidden\s*/>', '')
    $blk = [regex]::Replace($blk, '<w:unhideWhenUsed\s*/>', '')
    if ($blk -match '<w:uiPriority\s+w:val="[^"]*"\s*/>') {
        $blk = [regex]::Replace($blk, '<w:uiPriority\s+w:val="[^"]*"\s*/>', "<w:uiPriority w:val=`"$Priority`"/>", 1)
    } elseif ($blk -match '<w:basedOn\s+w:val="[^"]*"\s*/>') {
        $blk = [regex]::Replace($blk, '(<w:basedOn\s+w:val="[^"]*"\s*/>)', "`$1<w:uiPriority w:val=`"$Priority`"/>", 1)
    } elseif ($blk -match '<w:name\s+w:val="[^"]*"\s*/>') {
        $blk = [regex]::Replace($blk, '(<w:name\s+w:val="[^"]*"\s*/>)', "`$1<w:uiPriority w:val=`"$Priority`"/>", 1)
    }
    if ($blk -notmatch '<w:qFormat\s*/>') {
        if ($blk -match '<w:uiPriority\s+w:val="[^"]*"\s*/>') {
            $blk = [regex]::Replace($blk, '(<w:uiPriority\s+w:val="[^"]*"\s*/>)', '$1<w:qFormat/>', 1)
        }
    }
    return $rg.Replace($Xml, [System.Text.RegularExpressions.MatchEvaluator]{ param($x) $blk }, 1)
}
