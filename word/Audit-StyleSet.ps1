<#
.SYNOPSIS
    Audits a Word file (or every Word file in a folder) for the presence
    of required style ids. Catches the "renamed style breaks the script
    silently" trap.

.DESCRIPTION
    Parses styles.xml directly from a Word package, lists every defined
    styleId, and compares against caller-supplied required and optional
    style sets.

    Each required entry may include a UsedBy list naming which scripts
    depend on it, so a missing-style report points straight at what will
    break next.

    Exit codes:
        0 = clean (every required style present in every file)
        1 = drift (one or more required styles missing somewhere)
        2 = input error

.PARAMETER File
    Single Word file to audit. Either -File or -Folder is required.

.PARAMETER Folder
    Folder whose Word files to audit. Recurses one level only.

.PARAMETER Filter
    Optional filename filter for -Folder mode (default: *.docx).

.PARAMETER RequiredStyles
    Ordered hashtable mapping styleId to @{ Category; Type; UsedBy }.
    Every style in this set must be present for the audit to pass.

.PARAMETER OptionalStyles
    Ordered hashtable mapping styleId to @{ Category; Type }.
    Reported only when -ShowOptional is set.

.PARAMETER ReportTitle
    Title shown in the console report header (default: "Style set audit").

.PARAMETER ShowOptional
    Also report missing optional styles (default: only required misses).

.PARAMETER Json
    Emit a JSON report on stdout instead of the colored table.

.EXAMPLE
    .\Audit-StyleSet.ps1 -File "report.docx" -RequiredStyles $myStyles

.EXAMPLE
    .\Audit-StyleSet.ps1 -Folder "C:\reports" -RequiredStyles $req -OptionalStyles $opt -ShowOptional

.EXAMPLE
    .\Audit-StyleSet.ps1 -Folder "C:\reports" -RequiredStyles $req -Json | ConvertFrom-Json
#>

[CmdletBinding(DefaultParameterSetName='Folder')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [string]$Filter = '*.docx',

    [Parameter(Mandatory=$true)]
    [System.Collections.Specialized.OrderedDictionary]$RequiredStyles,

    [System.Collections.Specialized.OrderedDictionary]$OptionalStyles,

    [string]$ReportTitle = 'Style set audit',

    [switch]$ShowOptional,

    [switch]$Json
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# File list
# ---------------------------------------------------------------------------

$files = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) {
        Write-Error "File not found: $File"; exit 2
    }
    $files += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
        Write-Error "Folder not found: $Folder"; exit 2
    }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $files = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles($Filter) |
        Where-Object { $_.Extension -imatch '^\.(docx|dotx|docm|dotm)$' } |
        ForEach-Object { $_.FullName }
}

if ($files.Count -eq 0) { Write-Error "No files to audit."; exit 2 }

# ---------------------------------------------------------------------------
# Audit each file
# ---------------------------------------------------------------------------

function Get-StyleIds {
    param([string]$Path)
    $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, 'Read')
        try {
            $entry = $zip.GetEntry('word/styles.xml')
            if (-not $entry) { return @{} }
            $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            try { $stylesXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }

    $present = @{}
    $matches = [regex]::Matches($stylesXml, '<w:style\s+w:type="(\w+)"[^>]*w:styleId="([^"]+)"')
    foreach ($m in $matches) {
        $present[$m.Groups[2].Value] = $m.Groups[1].Value
    }
    return $present
}

$reportFiles = New-Object System.Collections.ArrayList
$anyRequiredMiss = $false

foreach ($f in $files) {
    $present = Get-StyleIds -Path $f
    $missingRequired = @()
    $typeMismatch    = @()
    $missingOptional = @()

    foreach ($k in $RequiredStyles.Keys) {
        if (-not $present.ContainsKey($k)) {
            $usedBy = if ($RequiredStyles[$k].UsedBy) { $RequiredStyles[$k].UsedBy -join ', ' } else { '' }
            $missingRequired += [pscustomobject]@{
                StyleId  = $k
                Category = $RequiredStyles[$k].Category
                Type     = $RequiredStyles[$k].Type
                UsedBy   = $usedBy
            }
        } elseif ($present[$k] -ne $RequiredStyles[$k].Type) {
            $usedBy = if ($RequiredStyles[$k].UsedBy) { $RequiredStyles[$k].UsedBy -join ', ' } else { '' }
            $typeMismatch += [pscustomobject]@{
                StyleId  = $k
                Expected = $RequiredStyles[$k].Type
                Actual   = $present[$k]
                Category = $RequiredStyles[$k].Category
                UsedBy   = $usedBy
            }
        }
    }
    if ($OptionalStyles) {
        foreach ($k in $OptionalStyles.Keys) {
            if (-not $present.ContainsKey($k)) {
                $missingOptional += [pscustomobject]@{
                    StyleId  = $k
                    Category = $OptionalStyles[$k].Category
                    Type     = $OptionalStyles[$k].Type
                }
            }
        }
    }

    if ($missingRequired.Count -gt 0 -or $typeMismatch.Count -gt 0) { $anyRequiredMiss = $true }

    [void]$reportFiles.Add([pscustomobject]@{
        File             = $f
        TotalStyles      = $present.Count
        MissingRequired  = $missingRequired
        TypeMismatch     = $typeMismatch
        MissingOptional  = $missingOptional
    })
}

# ---------------------------------------------------------------------------
# Render
# ---------------------------------------------------------------------------

if ($Json) {
    [pscustomobject]@{
        Clean = -not $anyRequiredMiss
        Files = $reportFiles
    } | ConvertTo-Json -Depth 6
} else {
    Write-Host ""
    Write-Host $ReportTitle -ForegroundColor Cyan
    Write-Host ""
    foreach ($r in $reportFiles) {
        $name = Split-Path -Leaf $r.File
        Write-Host "  $name" -ForegroundColor White
        Write-Host ("    Total styles defined: {0}" -f $r.TotalStyles) -ForegroundColor DarkGray
        if ($r.MissingRequired.Count -eq 0 -and $r.TypeMismatch.Count -eq 0) {
            Write-Host "    Required styles: all present" -ForegroundColor Green
        } else {
            if ($r.MissingRequired.Count -gt 0) {
                Write-Host ("    MISSING required ({0}):" -f $r.MissingRequired.Count) -ForegroundColor Red
                foreach ($m in $r.MissingRequired) {
                    $usedByText = if ($m.UsedBy) { "needed by: $($m.UsedBy)" } else { '' }
                    Write-Host ("      - {0,-26} [{1,-12}] {2}" -f $m.StyleId, $m.Category, $usedByText) -ForegroundColor Red
                }
            }
            if ($r.TypeMismatch.Count -gt 0) {
                Write-Host ("    TYPE MISMATCH ({0}):" -f $r.TypeMismatch.Count) -ForegroundColor Yellow
                foreach ($m in $r.TypeMismatch) {
                    Write-Host ("      - {0,-26} expected {1}, found {2}" -f $m.StyleId, $m.Expected, $m.Actual) -ForegroundColor Yellow
                }
            }
        }
        if ($ShowOptional -and $r.MissingOptional.Count -gt 0) {
            Write-Host ("    Missing optional ({0}):" -f $r.MissingOptional.Count) -ForegroundColor DarkGray
            foreach ($m in $r.MissingOptional) {
                Write-Host ("      - {0,-26} [{1}]" -f $m.StyleId, $m.Category) -ForegroundColor DarkGray
            }
        }
        Write-Host ""
    }
    if ($anyRequiredMiss) {
        Write-Host "Drift: at least one file is missing a required style id." -ForegroundColor Red
    } else {
        Write-Host "Clean. Every audited file defines every required style id." -ForegroundColor Green
    }
    Write-Host ""
}

if ($anyRequiredMiss) { exit 1 } else { exit 0 }
