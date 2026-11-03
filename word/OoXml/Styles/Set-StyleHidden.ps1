function Set-StyleHidden {
    <#
    .SYNOPSIS
        Marks a style as hidden in Word's style pane.

    .DESCRIPTION
        Sets semiHidden and unhideWhenUsed, removes qFormat, and sets
        uiPriority to 99 so hidden styles sort after recommended ones.

    .PARAMETER Xml
        The full styles.xml content as a string.

    .PARAMETER Id
        The w:styleId to hide.
    #>
    param(
        [string]$Xml,
        [string]$Id
    )
    $pat = '(?s)<w:style\b[^>]*?\bw:styleId="' + [regex]::Escape($Id) + '"[^>]*?>(.*?)</w:style>'
    $rg = [regex]::new($pat)
    $mm = $rg.Match($Xml)
    if (-not $mm.Success) { return $Xml }
    $blk = $mm.Value
    $blk = [regex]::Replace($blk, '<w:qFormat\s*/>', '')
    if ($blk -match '<w:uiPriority\s+w:val="[^"]*"\s*/>') {
        $blk = [regex]::Replace($blk, '<w:uiPriority\s+w:val="[^"]*"\s*/>', '<w:uiPriority w:val="99"/>', 1)
    } elseif ($blk -match '<w:basedOn\s+w:val="[^"]*"\s*/>') {
        $blk = [regex]::Replace($blk, '(<w:basedOn\s+w:val="[^"]*"\s*/>)', '$1<w:uiPriority w:val="99"/>', 1)
    } elseif ($blk -match '<w:name\s+w:val="[^"]*"\s*/>') {
        $blk = [regex]::Replace($blk, '(<w:name\s+w:val="[^"]*"\s*/>)', '$1<w:uiPriority w:val="99"/>', 1)
    }
    if ($blk -notmatch '<w:semiHidden\s*/>') {
        if ($blk -match '<w:uiPriority\s+w:val="[^"]*"\s*/>') {
            $blk = [regex]::Replace($blk, '(<w:uiPriority\s+w:val="[^"]*"\s*/>)', '$1<w:semiHidden/>', 1)
        }
    }
    if ($blk -notmatch '<w:unhideWhenUsed\s*/>') {
        if ($blk -match '<w:semiHidden\s*/>') {
            $blk = [regex]::Replace($blk, '(<w:semiHidden\s*/>)', '$1<w:unhideWhenUsed/>', 1)
        }
    }
    return $rg.Replace($Xml, [System.Text.RegularExpressions.MatchEvaluator]{ param($x) $blk }, 1)
}
