' StyleCharts VBA Macro for Excel
' Applies NWIRP publication styling to all charts in figure tabs (F5-xx).
'
' To use:
' 1. Open HM_Data_Model_v2.xlsx in Excel
' 2. Alt+F11 to open VBA editor
' 3. Insert > Module
' 4. Paste this entire file
' 5. Close VBA editor
' 6. Alt+F8 > StyleAllCharts > Run

Option Explicit

Private Const NAVY_SOFT As Long = &H8F3C1A
Private Const TAWNY_SOFT As Long = &H307AE8
Private Const SHAMROCK_SOFT As Long = &H8EB52A
Private Const PICTON_SOFT As Long = &HEEC57C
Private Const CLARET_SOFT As Long = &H5A33B8
Private Const HARVEST_SOFT As Long = &H40B8F0
Private Const MULBERRY_SOFT As Long = &HAC7ED0

Private Const WHISPER_LINE As Long = &HF2EEEE
Private Const WHITE_COLOR As Long = &HFFFFFF
Private Const BLACK_COLOR As Long = &H0

Sub StyleAllCharts()
    StyleWorkbook
    MsgBox "Styling complete. Manually adjust pie labels for clean positioning.", vbInformation
End Sub

' The same walk without the MsgBox. Call this one from COM automation: a hidden
' application has nowhere to show a dialog and hangs forever on it.
Sub StyleWorkbook()
    Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim sheetName As String

    For Each ws In ActiveWorkbook.Worksheets
        sheetName = ws.Name
        If IsFigureTab(sheetName) Then
            For Each chtObj In ws.ChartObjects
                Set cht = chtObj.Chart

                Dim chartType As Long
                chartType = cht.ChartType

                Select Case chartType
                    Case xlPie, xl3DPie, xlPieExploded, xl3DPieExploded
                        StylePieChart cht, sheetName
                    Case xlLine, xlLineMarkers, xlLineStacked, xlLineMarkersStacked, _
                         xlLineStacked100, xlLineMarkersStacked100
                        StyleLineChart cht, sheetName
                    Case Else
                        StyleBarChart cht, sheetName
                End Select
            Next chtObj
        End If
    Next ws
End Sub

Function IsFigureTab(name As String) As Boolean
    ' A tab named for the figure it holds - "Fig 5-2 Schools Infra",
    ' "Fig 1-1 Counts Line" - which is what the figure workbooks on file
    ' actually use. The job-number list below predates that convention.
    If name Like "Fig #-#*" Then
        IsFigureTab = True
        Exit Function
    End If

    Select Case name
        Case "1239-NWIRP", "1240-NWIRP", "1241-NWIRP", "1242-NWIRP", _
             "1243-NWIRP", "1244-NWIRP", "1245-NWIRP", "1246-NWIRP", _
             "1247-NWIRP", "1248-NWIRP", "1249-NWIRP", "1250-NWIRP"
            IsFigureTab = True
        Case Else
            IsFigureTab = False
    End Select
End Function

Sub StyleBarChart(cht As Chart, sheetName As String)
    Dim s As Series
    Dim i As Long
    Dim isHistogram As Boolean
    Dim isStacked As Boolean

    isHistogram = (sheetName = "1250-NWIRP")

    ' A stacked group is a different animal: Excel forces Overlap to 100, and
    ' the per-series drop shadow reads as grime between the bands.
    Select Case cht.ChartType
        Case xlColumnStacked, xlBarStacked, xlColumnStacked100, xlBarStacked100
            isStacked = True
    End Select

    cht.HasTitle = False

    ' Chart area
    With cht.ChartArea
        .Format.Fill.ForeColor.RGB = WHITE_COLOR
        .Format.Line.Visible = msoFalse
    End With

    ' Plot area
    With cht.PlotArea
        .Format.Fill.ForeColor.RGB = WHITE_COLOR
        .Format.Line.Visible = msoFalse
    End With

    ' Series colors
    ApplyBarSeriesColors cht, sheetName

    ' Series formatting
    For i = 1 To cht.SeriesCollection.Count
        Set s = cht.SeriesCollection(i)
        With s.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = WHITE_COLOR
            .Weight = 1.5
        End With

        If isStacked Then
            s.Format.Shadow.Visible = msoFalse
        Else
            With s.Format.Shadow
                .Visible = msoTrue
                .ForeColor.RGB = BLACK_COLOR
                .Transparency = 0.85
                .Blur = 4
                .OffsetX = 0
                .OffsetY = 2
            End With
        End If
    Next i

    ' Gap and overlap
    With cht.ChartGroups(1)
        If isStacked Then
            .GapWidth = 40
            .Overlap = 100
        ElseIf isHistogram Then
            .GapWidth = 0
            .Overlap = 0
        Else
            .GapWidth = 60
            .Overlap = 10
        End If
    End With

    ' Ensure axes exist
    cht.HasAxis(xlValue) = True
    cht.HasAxis(xlCategory) = True

    ' Y Axis
    On Error Resume Next
    With cht.Axes(xlValue)
        .Format.Line.Visible = msoFalse
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 8
        .TickLabels.Font.Color = BLACK_COLOR
        .TickLabels.NumberFormat = AxisValueFormat(sheetName)

        If .HasMajorGridlines Then
            With .MajorGridlines.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = WHISPER_LINE
                .Weight = 0.75
            End With
        End If

        If .HasTitle Then
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Name = "Arial"
                .Size = 8
                .Fill.ForeColor.RGB = BLACK_COLOR
                .Bold = msoFalse
            End With
        End If
    End With

    ' X Axis
    On Error Resume Next
    With cht.Axes(xlCategory)
        With .Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = WHISPER_LINE
            .Weight = 0.75
        End With
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 8
        .TickLabels.Font.Color = BLACK_COLOR
        .HasMajorGridlines = True
        With .MajorGridlines.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = WHISPER_LINE
            .Weight = 0.75
        End With
    End With
    On Error GoTo 0

    ' Y axis title
    On Error Resume Next
    With cht.Axes(xlValue)
        .HasTitle = True
        Select Case sheetName
            Case "1241-NWIRP", "1242-NWIRP", "1243-NWIRP", "1244-NWIRP", "1245-NWIRP"
                .AxisTitle.Text = "% of Respondents"
            Case "1246-NWIRP", "1247-NWIRP"
                .AxisTitle.Text = "Mean Recovery (%)"
            Case "1250-NWIRP"
                .AxisTitle.Text = "Number of Schools"
        End Select
        .AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Arial"
        .AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
        .AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BLACK_COLOR
    End With
    On Error GoTo 0

    ' Legend
    If cht.SeriesCollection.Count > 1 Then
        cht.HasLegend = True
        With cht.Legend
            .Position = xlLegendPositionBottom
            .IncludeInLayout = True
            .Format.TextFrame2.TextRange.Font.Name = "Arial"
            .Format.TextFrame2.TextRange.Font.Size = 9   ' one step up from the axis: on a stacked chart the legend is the only key to which band is which
            .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BLACK_COLOR
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoFalse
        End With
    Else
        cht.HasLegend = False
    End If

    ' Data labels on histogram only
    If isHistogram Then
        cht.SeriesCollection(1).HasDataLabels = True
        With cht.SeriesCollection(1).DataLabels
            .Font.Name = "Arial"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = BLACK_COLOR
            .NumberFormat = "#,##0"
            .Position = xlLabelPositionOutsideEnd
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoFalse
        End With
    End If
End Sub

' ============================================================
' LINE CHART STYLING
' ============================================================
' A time series is a line, and neither of the other two stylers fits it.
' StyleBarChart in particular sets ChartGroups(1).GapWidth, which a line
' group does not have, so routing a line chart there raises an error rather
' than styling it badly.
Sub StyleLineChart(cht As Chart, sheetName As String)
    Dim s As Series
    Dim i As Long

    cht.HasTitle = False

    ' Chart area
    With cht.ChartArea
        .Format.Fill.ForeColor.RGB = WHITE_COLOR
        .Format.Line.Visible = msoFalse
    End With

    ' Plot area
    With cht.PlotArea
        .Format.Fill.ForeColor.RGB = WHITE_COLOR
        .Format.Line.Visible = msoFalse
    End With

    ' Stroke first, colour second. On a line the weight and the colour both
    ' live on Format.Line, so the colour has to be applied last to survive.
    For i = 1 To cht.SeriesCollection.Count
        Set s = cht.SeriesCollection(i)
        With s.Format.Line
            .Visible = msoTrue
            .Weight = 2.25
            .DashStyle = msoLineSolid
        End With

        ' No shadow. The drop shadow that gives a bar its lift reads as a
        ' blur on a 2pt line.
        s.Format.Shadow.Visible = msoFalse

        ' Straight segments between real data points. A smoothed line invents
        ' values between them, which is a misrepresentation in a report.
        s.Smooth = False
        s.MarkerStyle = xlMarkerStyleNone
    Next i

    ApplyLineSeriesColors cht, sheetName

    ' Y Axis
    On Error Resume Next
    With cht.Axes(xlValue)
        .Format.Line.Visible = msoFalse
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 8
        .TickLabels.Font.Color = BLACK_COLOR
        .TickLabels.NumberFormat = AxisValueFormat(sheetName)

        If .HasMajorGridlines Then
            With .MajorGridlines.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = WHISPER_LINE
                .Weight = 0.75
            End With
        End If

        If .HasTitle Then
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Name = "Arial"
                .Size = 8
                .Fill.ForeColor.RGB = BLACK_COLOR
                .Bold = msoFalse
            End With
        End If
    End With
    On Error GoTo 0

    ' X Axis. A time axis keeps its baseline and loses its gridlines: the
    ' horizontal ones already carry the reading.
    On Error Resume Next
    With cht.Axes(xlCategory)
        With .Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = WHISPER_LINE
            .Weight = 0.75
        End With
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 8
        .TickLabels.Font.Color = BLACK_COLOR
        If .HasMajorGridlines Then .HasMajorGridlines = False
    End With
    On Error GoTo 0

    ' Legend
    If cht.SeriesCollection.Count > 1 Then
        cht.HasLegend = True
        With cht.Legend
            .Position = xlLegendPositionBottom
            .IncludeInLayout = True
            .Format.TextFrame2.TextRange.Font.Name = "Arial"
            .Format.TextFrame2.TextRange.Font.Size = 9   ' one step up from the axis: on a stacked chart the legend is the only key to which band is which
            .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BLACK_COLOR
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoFalse
        End With
    Else
        cht.HasLegend = False
    End If
End Sub

Function AxisValueFormat(sheetName As String) As String
    ' The value axis format for a figure, shared by the bar and line stylers
    ' so a pair of figures showing the same measure cannot drift apart.
    Select Case sheetName
        Case "1241-NWIRP", "1242-NWIRP", "1243-NWIRP", "1244-NWIRP", "1245-NWIRP"
            AxisValueFormat = "0%"
        Case "1246-NWIRP", "1247-NWIRP"
            AxisValueFormat = "0""%"""
        Case "Fig 1-2 Dollars Line", "Fig 1-2 Dollars Stacked"
            ' Billions, so the axis reads $1.5B rather than 1500000000
            AxisValueFormat = "$#,##0.0,,,""B"""
        Case Else
            AxisValueFormat = "#,##0"
    End Select
End Function

Sub StylePieChart(cht As Chart, sheetName As String)
    Dim pt As Point
    Dim i As Long

    cht.HasTitle = False

    With cht.ChartArea
        .Format.Fill.ForeColor.RGB = WHITE_COLOR
        .Format.Line.Visible = msoFalse
    End With
    With cht.PlotArea
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With

    ' Slice colors
    ApplyPieSliceColors cht, sheetName

    ' Slice formatting
    With cht.SeriesCollection(1)
        For i = 1 To .Points.Count
            Set pt = .Points(i)

            ' Reset any prior explosion
            pt.Explosion = 0

            ' White border
            With pt.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = WHITE_COLOR
                .Weight = 1.5
            End With

            ' Shadow
            With pt.Format.Shadow
                .Visible = msoTrue
                .ForeColor.RGB = BLACK_COLOR
                .Transparency = 0.85
                .Blur = 4
                .OffsetX = 0
                .OffsetY = 2
            End With
        Next i
    End With

    ' No legend
    cht.HasLegend = False

    ' Data labels: series-level defaults
    cht.SeriesCollection(1).HasDataLabels = True
    With cht.SeriesCollection(1).DataLabels
        .ShowCategoryName = True
        .ShowPercentage = True
        .ShowValue = False
        .ShowSeriesName = False
        .Separator = " "
        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Bold = True
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
        .Position = xlLabelPositionCenter
    End With

    ' Per-slice text color based on wedge luminance
    With cht.SeriesCollection(1)
        For i = 1 To .Points.Count
            Dim fillClr As Long
            Dim r As Long, g As Long, b As Long
            Dim lum As Double
            fillClr = .Points(i).Format.Fill.ForeColor.RGB
            r = fillClr And &HFF
            g = (fillClr \ &H100) And &HFF
            b = (fillClr \ &H10000) And &HFF
            lum = 0.299 * r + 0.587 * g + 0.114 * b
            If lum > 150 Then
                .Points(i).DataLabel.Font.Color = BLACK_COLOR
            Else
                .Points(i).DataLabel.Font.Color = WHITE_COLOR
            End If
            .Points(i).DataLabel.Font.Name = "Arial"
            .Points(i).DataLabel.Font.Size = 8
            .Points(i).DataLabel.Font.Bold = True
        Next i
    End With
End Sub

Sub ApplyBarSeriesColors(cht As Chart, sheetName As String)
    Select Case sheetName
        Case "1241-NWIRP", "1242-NWIRP"
            SetSeriesColor cht, 1, CLARET_SOFT
            SetSeriesColor cht, 2, PICTON_SOFT
            SetSeriesColor cht, 3, NAVY_SOFT
            SetSeriesColor cht, 4, HARVEST_SOFT
        Case "1243-NWIRP"
            SetSeriesColor cht, 1, TAWNY_SOFT
            SetSeriesColor cht, 2, SHAMROCK_SOFT
        Case "1244-NWIRP", "1245-NWIRP"
            SetSeriesColor cht, 1, NAVY_SOFT
            SetSeriesColor cht, 2, PICTON_SOFT
            SetSeriesColor cht, 3, HARVEST_SOFT
            SetSeriesColor cht, 4, CLARET_SOFT
        Case "1246-NWIRP"
            SetSeriesColor cht, 1, NAVY_SOFT
            SetSeriesColor cht, 2, TAWNY_SOFT
        Case "1247-NWIRP"
            SetSeriesColor cht, 1, NAVY_SOFT
            SetSeriesColor cht, 2, SHAMROCK_SOFT
        Case "1250-NWIRP"
            SetSeriesColor cht, 1, NAVY_SOFT

        ' Stacked, bottom band to top: Completed, past planning but not
        ' complete, still planning. The stack height is the portfolio total.
        Case "Fig 1-1 Counts Stacked", "Fig 1-2 Dollars Stacked"
            SetSeriesColor cht, 1, SHAMROCK_SOFT
            SetSeriesColor cht, 2, PICTON_SOFT
            SetSeriesColor cht, 3, TAWNY_SOFT
    End Select
End Sub

Sub ApplyLineSeriesColors(cht As Chart, sheetName As String)
    Dim colors(1 To 7) As Long
    Dim i As Long

    ' The palette in order, so an unmapped figure tab still comes out styled
    ' instead of keeping Excel's defaults. That is the deliberate difference
    ' from ApplyBarSeriesColors, which leaves an unmapped chart alone.
    colors(1) = NAVY_SOFT
    colors(2) = TAWNY_SOFT
    colors(3) = SHAMROCK_SOFT
    colors(4) = PICTON_SOFT
    colors(5) = CLARET_SOFT
    colors(6) = HARVEST_SOFT
    colors(7) = MULBERRY_SOFT

    ' Per-figure overrides. These two carry the same three categories in a
    ' different series order, so without the override the same category would
    ' change colour between the figures.
    Select Case sheetName
        ' Completed and Planning only. Total is the envelope, not a peer:
        ' it is flat from 2023 and on the dollar axis it is 12x Completed,
        ' which flattens Completed out of view.
        Case "Fig 1-1 Counts Line", "Fig 1-2 Dollars Line"
            colors(1) = SHAMROCK_SOFT
            colors(2) = TAWNY_SOFT
    End Select

    ' A line carries its colour on the stroke, not on a fill.
    For i = 1 To cht.SeriesCollection.Count
        If i <= 7 Then
            cht.SeriesCollection(i).Format.Line.ForeColor.RGB = colors(i)
        End If
    Next i
End Sub

Sub ApplyPieSliceColors(cht As Chart, sheetName As String)
    Select Case sheetName
        Case "1239-NWIRP"
            SetPieColor cht, 1, SHAMROCK_SOFT
            SetPieColor cht, 2, PICTON_SOFT
            SetPieColor cht, 3, HARVEST_SOFT
            SetPieColor cht, 4, TAWNY_SOFT
            SetPieColor cht, 5, CLARET_SOFT
        Case "1240-NWIRP"
            SetPieColor cht, 1, SHAMROCK_SOFT
            SetPieColor cht, 2, HARVEST_SOFT
            SetPieColor cht, 3, TAWNY_SOFT
        Case "1248-NWIRP"
            SetPieColor cht, 1, MULBERRY_SOFT
            SetPieColor cht, 2, CLARET_SOFT
            SetPieColor cht, 3, HARVEST_SOFT
        Case "1249-NWIRP"
            SetPieColor cht, 1, CLARET_SOFT
            SetPieColor cht, 2, MULBERRY_SOFT
            SetPieColor cht, 3, HARVEST_SOFT
    End Select
End Sub

Sub SetSeriesColor(cht As Chart, idx As Long, clr As Long)
    If idx <= cht.SeriesCollection.Count Then
        cht.SeriesCollection(idx).Format.Fill.ForeColor.RGB = clr
    End If
End Sub

Sub SetPieColor(cht As Chart, idx As Long, clr As Long)
    If idx <= cht.SeriesCollection(1).Points.Count Then
        cht.SeriesCollection(1).Points(idx).Format.Fill.ForeColor.RGB = clr
    End If
End Sub
