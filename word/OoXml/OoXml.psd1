@{
    ModuleVersion     = '1.0.0'
    RootModule        = 'OoXml.psm1'
    GUID              = 'a3f7e2d1-8b4c-4f6e-9a1d-5c3b7e8f2a4d'
    Author            = 'Ovoco'
    Description       = 'Generic OOXML style and numbering manipulation functions for Word documents (.docx/.dotx). Operates on XML strings extracted from ZIP packages.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Edit-StyleBlock',
        'Remove-StyleBlock',
        'Get-AllStyleIdsByType',
        'Set-StyleBasedOn',
        'Set-StyleColor',
        'Edit-RunRPrInStyle',
        'Set-StyleRecommended',
        'Set-StyleHidden',
        'Get-HeadingAbstractNumIds',
        'Set-NumberingSuffix',
        'Sync-ListIndents',
        'Set-BulletGlyph'
    )
    VariablesToExport = @(
        'ListSingleSpacingEdit',
        'BodyTextSpacingEdit',
        'SuppressNumbering'
    )
    CmdletsToExport   = @()
    AliasesToExport   = @()
}
