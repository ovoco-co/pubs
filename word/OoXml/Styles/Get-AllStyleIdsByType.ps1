function Get-AllStyleIdsByType {
    <#
    .SYNOPSIS
        Lists all styleIds of a given type from styles.xml.

    .DESCRIPTION
        Scans the styles.xml string for all w:style elements with the specified
        w:type attribute and returns their styleId values as a list.

    .PARAMETER Xml
        The full styles.xml content as a string.

    .PARAMETER Type
        The style type to match (e.g., 'paragraph', 'character', 'table', 'numbering').
    #>
    param(
        [string]$Xml,
        [string]$Type
    )
    $pattern = '<w:style\b[^>]*?\bw:type="' + [regex]::Escape($Type) + '"[^>]*?\bw:styleId="([^"]+)"'
    $rgx = [regex]::new($pattern)
    $ids = New-Object System.Collections.Generic.List[string]
    foreach ($m in $rgx.Matches($Xml)) { [void]$ids.Add($m.Groups[1].Value) }
    # Comma: without it an empty list returns $null and a one-element list
    # returns a bare string. See Get-HeadingAbstractNumIds.
    return ,$ids
}
