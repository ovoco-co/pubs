function Get-HeadingAbstractNumIds {
    <#
    .SYNOPSIS
        Detects abstractNumIds linked to Heading1-9 styles.

    .DESCRIPTION
        Walks each Heading style (Heading1 through Heading9) in styles.xml to
        find its numId, resolves the numId to an abstractNumId in numbering.xml,
        then follows any numStyleLink chains to find the actual abstractNum
        containing the level definitions.

        Returns a HashSet of abstractNumId strings.

    .PARAMETER StylesXml
        The full styles.xml content as a string.

    .PARAMETER NumberingXml
        The full numbering.xml content as a string.
    #>
    param(
        [string]$StylesXml,
        [string]$NumberingXml
    )
    $headingAbstractNumIds = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($hsid in @('Heading1','Heading2','Heading3','Heading4','Heading5','Heading6','Heading7','Heading8','Heading9')) {
        $sp = '(?s)<w:style\b[^>]*?\bw:styleId="' + [regex]::Escape($hsid) + '"[^>]*?>.*?</w:style>'
        $sm = [regex]::Match($StylesXml, $sp)
        if (-not $sm.Success) { continue }
        $nidm = [regex]::Match($sm.Value, '<w:numId\s+w:val="(\d+)"\s*/>')
        if (-not $nidm.Success) { continue }
        $nid = $nidm.Groups[1].Value
        $am = [regex]::Match($NumberingXml, '<w:num w:numId="' + [regex]::Escape($nid) + '"[^>]*>\s*<w:abstractNumId w:val="(\d+)"', 'Singleline')
        if ($am.Success) { [void]$headingAbstractNumIds.Add($am.Groups[1].Value) }
    }

    # Follow numStyleLink chains
    $followLinks = [System.Collections.Generic.HashSet[string]]::new($headingAbstractNumIds)
    foreach ($absid in $headingAbstractNumIds) {
        $ap = '(?s)<w:abstractNum [^>]*abstractNumId="' + [regex]::Escape($absid) + '"[^>]*>.*?</w:abstractNum>'
        $abm = [regex]::Match($NumberingXml, $ap)
        if (-not $abm.Success) { continue }
        $nslMatch = [regex]::Match($abm.Value, '<w:numStyleLink w:val="([^"]+)"')
        if (-not $nslMatch.Success) { continue }
        $linkedStyleId = $nslMatch.Groups[1].Value
        $linkedStyle = [regex]::Match($StylesXml, '(?s)<w:style\b[^>]*?\bw:styleId="' + [regex]::Escape($linkedStyleId) + '"[^>]*?>.*?</w:style>')
        if (-not $linkedStyle.Success) { continue }
        $linkedNumIdMatch = [regex]::Match($linkedStyle.Value, '<w:numId\s+w:val="(\d+)"\s*/>')
        if (-not $linkedNumIdMatch.Success) { continue }
        $linkedNumId = $linkedNumIdMatch.Groups[1].Value
        $linkedAbsMatch = [regex]::Match($NumberingXml, '<w:num w:numId="' + [regex]::Escape($linkedNumId) + '"[^>]*>\s*<w:abstractNumId w:val="(\d+)"', 'Singleline')
        if ($linkedAbsMatch.Success) { [void]$followLinks.Add($linkedAbsMatch.Groups[1].Value) }
    }
    return $followLinks
}
