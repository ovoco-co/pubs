function Set-StyleBasedOn {
    <#
    .SYNOPSIS
        Changes or inserts a w:basedOn reference on a style block.

    .DESCRIPTION
        Replaces an existing w:basedOn element or inserts one after the w:name
        element (following OOXML element order: name, basedOn, next, link).

    .PARAMETER Block
        The style block XML string (a single w:style element).

    .PARAMETER NewId
        The new basedOn styleId value.
    #>
    param(
        [string]$Block,
        [string]$NewId
    )
    if ($Block -match '<w:basedOn\s+w:val="[^"]*"\s*/>') {
        return [regex]::Replace($Block, '<w:basedOn\s+w:val="[^"]*"\s*/>', "<w:basedOn w:val=`"$NewId`"/>")
    }
    if ($Block -match '<w:name\s+w:val="[^"]*"\s*/>') {
        return [regex]::Replace($Block, '(<w:name\s+w:val="[^"]*"\s*/>)', "`$1<w:basedOn w:val=`"$NewId`"/>", 1)
    }
    return [regex]::Replace($Block, '(<w:style\b[^>]*?>)', "`$1<w:basedOn w:val=`"$NewId`"/>", 1)
}
