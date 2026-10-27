<#
.SYNOPSIS
    Builds a PowerPoint figures deck from the charts in an Excel workbook, one
    chart per slide.

.DESCRIPTION
    Each worksheet that carries a chart becomes one slide. The slide is NAMED
    after the worksheet, which is the figure key StyleCharts_PowerPoint.bas
    reads: run the macro afterwards and every figure picks up its own colours
    and number format without anyone having to keep slide numbers in step.

    The deck is authored at the size the figure will occupy in Word (6.5 in by
    default, the NIST text column), so what is on the slide is what lands on
    the page. The slide is a workspace, not a layout: the chart fills it edge
    to edge and the figure caption is written in Word, not here.

    If the workbook is already open in Excel the running instance is reused, so
    the script does not fight the user for the file lock. PowerPoint has to be
    visible; it has no headless mode.

.PARAMETER Workbook
    Path to the .xlsx holding the charts.

.PARAMETER Output
    Path to write the .pptx to. Defaults to the workbook path with a .pptx
    extension.

.PARAMETER Sheets
    Worksheets to place, in slide order. Defaults to every worksheet that has
    at least one chart, in workbook order.

.PARAMETER SlideWidthInches
    Slide width. Default 6.5, the NIST text column.

.PARAMETER SlideHeightInches
    Slide height. Default 4.5.

.PARAMETER Force
    Overwrite an existing output file.

.EXAMPLE
    .\Build-ChartDeck.ps1 -Workbook C:\...\HM_Ch1_Figures.xlsx

.EXAMPLE
    .\Build-ChartDeck.ps1 -Workbook C:\...\HM_Ch1_Figures.xlsx -SlideWidthInches 13.333 -SlideHeightInches 7.5
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Workbook,

    [string]$Output,

    [string[]]$Sheets,

    [double]$SlideWidthInches = 6.5,

    [double]$SlideHeightInches = 4.5,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$ppLayoutBlank = 12

function Resolve-WindowsPath {
    param([string]$Path)
    return [System.IO.Path]::GetFullPath($Path)
}

$Workbook = Resolve-WindowsPath $Workbook
if (-not (Test-Path -LiteralPath $Workbook)) { throw "Workbook not found: $Workbook" }

if (-not $Output) {
    $Output = [System.IO.Path]::ChangeExtension($Workbook, '.pptx')
}
$Output = Resolve-WindowsPath $Output

if ((Test-Path -LiteralPath $Output) -and -not $Force) {
    throw "Output already exists: $Output  (use -Force to overwrite)"
}

# ------------------------------------------------------------------
# Excel: reuse a running instance if it already has the workbook open
# ------------------------------------------------------------------
$xl = $null
$startedExcel = $false
$openedWorkbook = $false
$wb = $null

try {
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    $xl = $null
}

if ($xl) {
    foreach ($candidate in $xl.Workbooks) {
        if ($candidate.FullName -eq $Workbook) { $wb = $candidate; break }
    }
}

if (-not $wb) {
    if (-not $xl) {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $startedExcel = $true
    }
    $wb = $xl.Workbooks.Open($Workbook)
    $openedWorkbook = $true
}

$ppt = $null
$pres = $null

try {
    # --------------------------------------------------------------
    # Which sheets, in which order
    # --------------------------------------------------------------
    if (-not $Sheets) {
        $Sheets = @()
        foreach ($ws in $wb.Worksheets) {
            if ($ws.ChartObjects().Count -gt 0) { $Sheets += $ws.Name }
        }
    }
    if ($Sheets.Count -eq 0) { throw "No worksheet in $Workbook carries a chart." }

    Write-Host ("Placing {0} chart(s): {1}" -f $Sheets.Count, ($Sheets -join ', '))

    # --------------------------------------------------------------
    # PowerPoint
    # --------------------------------------------------------------
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = 1     # PowerPoint has no headless mode
    $pres = $ppt.Presentations.Add()

    $pres.PageSetup.SlideWidth  = $SlideWidthInches * 72
    $pres.PageSetup.SlideHeight = $SlideHeightInches * 72

    $index = 0
    foreach ($sheetName in $Sheets) {
        $index++
        $ws = $wb.Worksheets.Item($sheetName)
        if ($ws.ChartObjects().Count -eq 0) { throw "Sheet '$sheetName' has no chart." }

        $slide = $pres.Slides.Add($index, $ppLayoutBlank)
        $slide.Name = $sheetName

        # Copy from Excel, paste into PowerPoint. The clipboard is shared with
        # everything else on the desktop and the paste can arrive before the
        # copy has landed, so retry rather than trusting the first attempt.
        $shape = $null
        for ($attempt = 1; $attempt -le 5 -and -not $shape; $attempt++) {
            $ws.ChartObjects(1).Copy() | Out-Null
            Start-Sleep -Milliseconds (250 * $attempt)
            try {
                $shape = $slide.Shapes.Paste()
            } catch {
                $shape = $null
                Write-Host ("  retry {0} on '{1}': {2}" -f $attempt, $sheetName, $_.Exception.Message)
            }
        }
        if (-not $shape) { throw "Could not paste the chart from '$sheetName'." }

        # Fill the slide. The caption belongs in Word, so nothing is reserved
        # at the top for a title.
        $shape.LockAspectRatio = 0
        $shape.Left   = 0
        $shape.Top    = 0
        $shape.Width  = $pres.PageSetup.SlideWidth
        $shape.Height = $pres.PageSetup.SlideHeight

        $hasChart = $false
        try { $hasChart = ($shape.HasChart -ne 0) } catch { $hasChart = $false }

        Write-Host ("  slide {0}: '{1}'  hasChart={2}" -f $index, $slide.Name, $hasChart)
        if (-not $hasChart) {
            Write-Warning "Slide $index pasted as a picture, not a chart. StyleCharts_PowerPoint.bas will skip it."
        }
    }

    if (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Force }
    $pres.SaveAs($Output)
    Write-Host ("Wrote {0}" -f $Output)
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt)  { try { $ppt.Quit()  } catch {} }
    if ($openedWorkbook -and $wb) { try { $wb.Close($false) } catch {} }
    if ($startedExcel -and $xl)   { try { $xl.Quit() } catch {} }
}
