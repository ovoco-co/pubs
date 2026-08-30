<#
.SYNOPSIS
    Normalizes 3-column body tables to a consistent fixed layout: 2" / 0.5" / 0.5".

.DESCRIPTION
    This is the table-width pass for the NIST pipeline. It walks
    word/document.xml, finds tables whose tblGrid has exactly three columns,
    and rewrites the table grid plus every cell width so the columns are
    consistent:

        column 1 = 2.0"  (2880 twips)
        column 2 = 0.5"  ( 720 twips)
        column 3 = 0.5"  ( 720 twips)

    The table is switched to fixed layout and its preferred width is set to
    the summed width (4320 twips = 3.0"). Tables styled SidebarTable are left
    alone; they have their own 3" float treatment.

    This script only edits table geometry. It does not touch runs, so
    superscripts, subscripts, and other inline formatting survive unchanged.
    Idempotent: re-running makes the same widths again.

    Backs up to <name>.tblwidths.bak before writing.

.PARAMETER File
    Path to a single Word file. Either this or -Folder is required.

.PARAMETER Folder
    Folder of Word files. All .docx/.dotx/.docm/.dotm files in it are processed.

.PARAMETER FirstColumnInches
    Width for the first column. Defaults to 2.0.

.PARAMETER SecondColumnInches
    Width for the second column. Defaults to 0.5.

.PARAMETER ThirdColumnInches
    Width for the third column. Defaults to 0.5.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-NISTThreeColumnTableWidths.ps1 -Folder "C:\docs"
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$File,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$Folder,

    [double]$FirstColumnInches  = 2.0,
    [double]$SecondColumnInches = 0.5,
    [double]$ThirdColumnInches  = 0.5,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ---------------------------------------------------------------------------
# Target list
# ---------------------------------------------------------------------------

$targets = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw "File not found: $File" }
    $targets += (Resolve-Path -LiteralPath $File).ProviderPath
} else {
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Folder not found: $Folder" }
    $folderResolved = (Resolve-Path -LiteralPath $Folder).ProviderPath
    $targets = [System.IO.DirectoryInfo]::new($folderResolved).GetFiles() |
        Where-Object { $_.Extension -match '^\.(docx|dotx|docm|dotm)$' } |
        Sort-Object Name |
        ForEach-Object { $_.FullName }
}
if ($targets.Count -eq 0) { throw "No supported Word files found." }

$widths = @(
    [int][Math]::Round($FirstColumnInches  * 1440),
    [int][Math]::Round($SecondColumnInches * 1440),
    [int][Math]::Round($ThirdColumnInches  * 1440)
)
$totalWidth = ($widths | Measure-Object -Sum).Sum

Write-Host "Set NIST Three-Column Table Widths" -ForegroundColor Cyan
Write-Host "  Layout: $FirstColumnInches`" / $SecondColumnInches`" / $ThirdColumnInches`" ($totalWidth twips total)" -ForegroundColor Cyan
Write-Host "  Targets:" -ForegroundColor Cyan
$targets | ForEach-Object { Write-Host "    $_" }
Write-Host ""

if (-not $Force) {
    $reply = Read-Host "Proceed? (y/N)"
    if ($reply -notmatch '^[yY]') { throw "Aborted." }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Read-Entries {
    param([string]$Path)
    $dict = New-Object 'System.Collections.Generic.Dictionary[string,byte[]]'
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                $ms = New-Object System.IO.MemoryStream
                $es = $e.Open()
                try { $es.CopyTo($ms) } finally { $es.Dispose() }
                $dict[$e.FullName] = $ms.ToArray()
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
    return ,$dict
}

function Get-XmlText {
    param([byte[]]$Bytes)
    return $utf8NoBom.GetString($Bytes).TrimStart([char]0xFEFF)
}

function Get-IntAttr {
    param([System.Xml.XmlElement]$Node, [string]$Name)
    if (-not $Node) { return $null }
    $raw = $Node.GetAttribute($Name, $W)
    if (-not $raw) { $raw = $Node.GetAttribute($Name) }
    if (-not $raw) { return $null }
    $n = 0
    if ([int]::TryParse($raw, [ref]$n)) { return $n }
    return $null
}

function Ensure-TablePrNode {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$Table,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $tblPr = $Table.SelectSingleNode('./w:tblPr', $Ns)
    if ($tblPr) { return $tblPr }

    $tblPr = $Doc.CreateElement('w', 'tblPr', $W)
    $firstChild = $Table.FirstChild
    if ($firstChild) { [void]$Table.InsertBefore($tblPr, $firstChild) } else { [void]$Table.AppendChild($tblPr) }
    return $tblPr
}

function Ensure-ChildAfter {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$Parent,
        [string]$ChildName,
        [string]$AfterName,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $node = $Parent.SelectSingleNode("./w:$ChildName", $Ns)
    if ($node) { return $node }

    $node = $Doc.CreateElement('w', $ChildName, $W)
    $after = if ($AfterName) { $Parent.SelectSingleNode("./w:$AfterName", $Ns) } else { $null }
    if ($after) { [void]$Parent.InsertAfter($node, $after) } else { [void]$Parent.PrependChild($node) }
    return $node
}

function Ensure-ChildBefore {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$Parent,
        [string]$ChildName,
        [string]$BeforeName,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $node = $Parent.SelectSingleNode("./w:$ChildName", $Ns)
    if ($node) { return $node }

    $node = $Doc.CreateElement('w', $ChildName, $W)
    $before = if ($BeforeName) { $Parent.SelectSingleNode("./w:$BeforeName", $Ns) } else { $null }
    if ($before) { [void]$Parent.InsertBefore($node, $before) } else { [void]$Parent.AppendChild($node) }
    return $node
}

function Set-TcWidth {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$Tc,
        [int]$WidthTwips,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $tcPr = $Tc.SelectSingleNode('./w:tcPr', $Ns)
    if (-not $tcPr) {
        $tcPr = $Doc.CreateElement('w', 'tcPr', $W)
        $firstChild = $Tc.FirstChild
        if ($firstChild) { [void]$Tc.InsertBefore($tcPr, $firstChild) } else { [void]$Tc.AppendChild($tcPr) }
    }

    $tcW = $tcPr.SelectSingleNode('./w:tcW', $Ns)
    if (-not $tcW) {
        $tcW = $Doc.CreateElement('w', 'tcW', $W)
        $gridSpan = $tcPr.SelectSingleNode('./w:gridSpan', $Ns)
        if ($gridSpan) { [void]$tcPr.InsertBefore($tcW, $gridSpan) } else { [void]$tcPr.PrependChild($tcW) }
    }
    [void]$tcW.SetAttribute('w', $W, $WidthTwips.ToString())
    [void]$tcW.SetAttribute('type', $W, 'dxa')
}

function Get-RowSpanPlan {
    param(
        [System.Xml.XmlElement]$Row,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $spans = New-Object System.Collections.Generic.List[int]
    $trPr = $Row.SelectSingleNode('./w:trPr', $Ns)
    $gridBefore = 0
    $gridAfter  = 0
    if ($trPr) {
        $gb = $trPr.SelectSingleNode('./w:gridBefore', $Ns)
        if ($gb) {
            $v = Get-IntAttr -Node $gb -Name 'val'
            if ($null -ne $v) { $gridBefore = $v }
        }
        $ga = $trPr.SelectSingleNode('./w:gridAfter', $Ns)
        if ($ga) {
            $v = Get-IntAttr -Node $ga -Name 'val'
            if ($null -ne $v) { $gridAfter = $v }
        }
    }

    foreach ($tc in $Row.SelectNodes('./w:tc', $Ns)) {
        $span = 1
        $tcPr = $tc.SelectSingleNode('./w:tcPr', $Ns)
        if ($tcPr) {
            $gridSpan = $tcPr.SelectSingleNode('./w:gridSpan', $Ns)
            if ($gridSpan) {
                $v = Get-IntAttr -Node $gridSpan -Name 'val'
                if ($null -ne $v -and $v -gt 0) { $span = $v }
            }
        }
        $spans.Add($span)
    }
    $total = $gridBefore + $gridAfter
    foreach ($span in $spans) { $total += $span }
    return [pscustomobject]@{
        GridBefore   = $gridBefore
        GridAfter    = $gridAfter
        Spans        = $spans
        TotalColumns  = $total
    }
}

# ---------------------------------------------------------------------------
# Process each file
# ---------------------------------------------------------------------------

foreach ($target in $targets) {
    Write-Host ""
    Write-Host "Processing $target..." -ForegroundColor Yellow

    $entries   = Read-Entries -Path $target
    $srcXmlRaw = $null
    if (-not $entries.ContainsKey('word/document.xml')) { Write-Warning "  No word/document.xml. Skipping."; continue }
    $srcXmlRaw = Get-XmlText -Bytes $entries['word/document.xml']

    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($srcXmlRaw)
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('w', $W)

    $tables = @($doc.SelectNodes('//w:tbl', $ns))
    $normalized = 0
    $skipped = 0

    foreach ($tbl in $tables) {
        $styleNode = $tbl.SelectSingleNode('./w:tblPr/w:tblStyle', $ns)
        if ($styleNode) {
            $styleId = $styleNode.GetAttribute('val', $W)
            if (-not $styleId) { $styleId = $styleNode.GetAttribute('val') }
            if ($styleId -eq 'SidebarTable') { $skipped++; continue }
        }

        $tblGrid = $tbl.SelectSingleNode('./w:tblGrid', $ns)
        if (-not $tblGrid) { $skipped++; continue }
        $gridCols = @($tblGrid.SelectNodes('./w:gridCol', $ns))
        if ($gridCols.Count -ne 3) { continue }

        $tblPr = Ensure-TablePrNode -Doc $doc -Table $tbl -Ns $ns

        $rows = @($tbl.SelectNodes('./w:tr', $ns))
        if ($rows.Count -eq 0) { $skipped++; continue }

        $rowPlans = @()
        $simple = $true
        foreach ($row in $rows) {
            $plan = Get-RowSpanPlan -Row $row -Ns $ns
            if ($plan.TotalColumns -ne 3) { $simple = $false; break }
            $rowPlans += ,$plan
        }
        if (-not $simple) { $skipped++; continue }

        # tblW: fixed 3" table width.
        $tblW = Ensure-ChildAfter -Doc $doc -Parent $tblPr -ChildName 'tblW' -AfterName 'tblStyle' -Ns $ns
        [void]$tblW.SetAttribute('w', $W, $totalWidth.ToString())
        [void]$tblW.SetAttribute('type', $W, 'dxa')

        # tblLayout: fixed.
        $tblLayout = Ensure-ChildBefore -Doc $doc -Parent $tblPr -ChildName 'tblLayout' -BeforeName 'tblLook' -Ns $ns
        [void]$tblLayout.SetAttribute('type', $W, 'fixed')

        # tblGrid: 2" / 0.5" / 0.5"
        for ($i = 0; $i -lt 3; $i++) {
            $gridCols[$i].SetAttribute('w', $W, $widths[$i].ToString())
        }

        # tcW: apply the width implied by each row's spans.
        for ($ri = 0; $ri -lt $rows.Count; $ri++) {
            $row = $rows[$ri]
            $plan = $rowPlans[$ri]
            $cells = @($row.SelectNodes('./w:tc', $ns))
            $col = $plan.GridBefore
            $cellIndex = 0
            foreach ($span in $plan.Spans) {
                if ($col + $span -gt 3) { $simple = $false; break }
                $w = 0
                for ($i = 0; $i -lt $span; $i++) { $w += $widths[$col + $i] }
                $tc = $cells[$cellIndex]
                Set-TcWidth -Doc $doc -Tc $tc -WidthTwips $w -Ns $ns
                $col += $span
                $cellIndex++
            }
            if (-not $simple) { break }
        }

        if (-not $simple) { $skipped++; continue }
        $normalized++
    }

    if ($normalized -eq 0) {
        Write-Host "  No simple 3-column tables found. File unchanged." -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Normalized $normalized table(s) to 2`"/0.5`"/0.5`" fixed layout" -ForegroundColor Gray
    if ($skipped -gt 0) {
        Write-Host "  Skipped $skipped SidebarTable or unsupported table(s)." -ForegroundColor DarkGray
    }

    $xmlMs = New-Object System.IO.MemoryStream
    $xmlSettings = New-Object System.Xml.XmlWriterSettings
    $xmlSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $xmlSettings.OmitXmlDeclaration = $false
    $xmlSettings.NewLineHandling = [System.Xml.NewLineHandling]::None
    $xmlWriter = [System.Xml.XmlWriter]::Create($xmlMs, $xmlSettings)
    try {
        $doc.Save($xmlWriter)
    } finally {
        $xmlWriter.Dispose()
    }
    $entries['word/document.xml'] = $xmlMs.ToArray()

    # Back up, then write the package back.
    if (-not $Force -and (Test-Path -LiteralPath "$target.tblwidths.bak")) {
        Write-Warning "  Backup $target.tblwidths.bak already exists; leaving it as-is."
    } elseif (-not (Test-Path -LiteralPath "$target.tblwidths.bak")) {
        Copy-Item -LiteralPath $target -Destination "$target.tblwidths.bak" -Force
    }

    $tempZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "tblwidths-$([guid]::NewGuid()).zip")
    $outStream = [System.IO.File]::Create($tempZip)
    try {
        $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $entries.Keys) {
                $ne = $outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try { $os.Write($entries[$name], 0, $entries[$name].Length) } finally { $os.Dispose() }
            }
        } finally { $outZip.Dispose() }
    } finally { $outStream.Dispose() }

    Move-Item -LiteralPath $tempZip -Destination $target -Force
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
