' StyleCharts VBA Macro for PowerPoint
' Applies NWIRP publication styling to all charts in the active presentation.
'
' To use:
' 1. Open Ch5_Figures.pptx in PowerPoint
' 2. Alt+F11 to open VBA editor
' 3. Insert > Module
' 4. Paste this entire file
' 5. Close VBA editor
' 6. Alt+F8 > StyleAllCharts > Run
'
' The macro styles all 12 slides automatically. Run once, then
' manually drag pie labels into clean positions and adjust any
' overlapping text.

Option Explicit

' ============================================================
' COLOR CONSTANTS
' ============================================================
Private Const NAVY_SOFT As Long = &H8F3C1A&
Private Const TAWNY_SOFT As Long = &H307AE8&
Private Const SHAMROCK_SOFT As Long = &H8EB52A&
Private Const PICTON_SOFT As Long = &HEEC57C&
Private Const CLARET_SOFT As Long = &H5A33B8&
Private Const HARVEST_SOFT As Long = &H40B8F0&
Private Const MULBERRY_SOFT As Long = &HAC7ED0&

Private Const CHARCOAL As Long = &H3B3333&
Private Const SLATE As Long = &H766B6B&
Private Const WHISPER_LINE As Long = &HF2EEEE&
Private Const SILVER_LINE As Long = &HDED8D8&
Private Const WHITE_COLOR As Long = &HFFFFFF&
Private Const BLACK_COLOR As Long = &H0&

' Note: VBA uses BGR format, not RGB. The constants above are already BGR.
' To convert RGB hex to VBA: #1A3C8F -> &H8F3C1A

' ============================================================
' MAIN ENTRY POINT
' ============================================================
Sub StyleAllCharts()
    Dim sld As Slide
    Dim shp As Shape
    Dim cht As Chart
    Dim slideNum As Long

    For Each sld In ActivePresentation.Slides
        slideNum = sld.SlideIndex

        ' Style the chart (leave size and position as-is)
        For Each shp In sld.Shapes
            If shp.HasChart Then
                Set cht = shp.Chart

                ' Determine chart type and apply styling
                Select Case cht.ChartType
                    Case xlPie, xl3DPie
                        StylePieChart cht, slideNum
                    Case xlColumnClustered, xlBarClustered
                        StyleBarChart cht, slideNum
                End Select
            End If
        Next shp
    Next sld

    MsgBox "Styling complete. Manually adjust pie labels for clean positioning.", vbInformation
End Sub

' ============================================================
' BAR CHART STYLING
' ============================================================
Sub StyleBarChart(cht As Chart, slideNum As Long)
    Dim s As Series
    Dim i As Long
    Dim isHistogram As Boolean

    isHistogram = (slideNum = 12)

    ' Remove chart title
    cht.HasTitle = False

    ' Chart area: white fill, no border
    With cht.ChartArea
        .Format.Fill.ForeColor.RGB = WHITE_COLOR
        .Format.Line.Visible = msoFalse
    End With

    ' Plot area: white fill, no border
    With cht.PlotArea
        .Format.Fill.ForeColor.RGB = WHITE_COLOR
        .Format.Line.Visible = msoFalse
    End With

    ' Apply series colors
    ApplyBarSeriesColors cht, slideNum

    ' Series formatting: shadow, gap, overlap
    For i = 1 To cht.SeriesCollection.Count
        Set s = cht.SeriesCollection(i)

        ' No border on bars
        s.Format.Line.Visible = msoFalse

        ' Outer shadow: black 85% transparent, blur 4, distance 2, angle 90
        With s.Format.Shadow
            .Visible = msoTrue
            .Type = msoShadow21 ' Outer offset bottom
            .ForeColor.RGB = BLACK_COLOR
            .Transparency = 0.85
            .Blur = 4
            .OffsetX = 0
            .OffsetY = 2
        End With

        ' Inner shadow for soft top
        With s.Format.SoftEdge
            .Type = msoSoftEdgeType1
        End With
    Next i

    ' Gap and overlap
    With cht.ChartGroups(1)
        If isHistogram Then
            .GapWidth = 0
            .Overlap = 0
        Else
            .GapWidth = 60
            .Overlap = 10
        End If
    End With

    ' Y Axis (Value Axis)
    With cht.Axes(xlValue)
        .Format.Line.Visible = msoFalse
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 8
        .TickLabels.Font.Color = SLATE
        .TickLabels.NumberFormat = "#,##0"

        ' Gridlines
        If .HasMajorGridlines Then
            With .MajorGridlines.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = WHISPER_LINE
                .Weight = 0.75
            End With
        End If

        ' Axis title
        If .HasTitle Then
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Name = "Arial"
                .Size = 8
                .Fill.ForeColor.RGB = SLATE
                .Bold = msoFalse
            End With
        End If
    End With

    ' X Axis (Category Axis)
    With cht.Axes(xlCategory)
        With .Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = SILVER_LINE
            .Weight = 0.5
        End With
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 8
        .TickLabels.Font.Color = SLATE

        ' Delete vertical gridlines
        If .HasMajorGridlines Then .HasMajorGridlines = False
    End With

    ' Legend
    If cht.SeriesCollection.Count > 1 Then
        cht.HasLegend = True
        With cht.Legend
            .Position = xlLegendPositionBottom
            .IncludeInLayout = False
            .Format.TextFrame2.TextRange.Font.Name = "Arial"
            .Format.TextFrame2.TextRange.Font.Size = 8
            .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = SLATE
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoFalse
        End With
    Else
        cht.HasLegend = False
    End If

    ' Data labels (histogram only)
    If isHistogram Then
        cht.SeriesCollection(1).HasDataLabels = True
        With cht.SeriesCollection(1).DataLabels
            .Font.Name = "Arial"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = CHARCOAL
            .NumberFormat = "#,##0"
            .Position = xlLabelPositionOutsideEnd
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoFalse
        End With
    End If
End Sub

' ============================================================
' PIE CHART STYLING
' ============================================================
Sub StylePieChart(cht As Chart, slideNum As Long)
    Dim pt As Point
    Dim i As Long
    Dim maxVal As Double
    Dim maxIdx As Long

    ' Remove chart title
    cht.HasTitle = False

    ' Chart area and plot area: white, no border
    With cht.ChartArea
        .Format.Fill.ForeColor.RGB = WHITE_COLOR
        .Format.Line.Visible = msoFalse
    End With
    With cht.PlotArea
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With

    ' Apply slice colors
    ApplyPieSliceColors cht, slideNum

    ' Style each slice: white border, shadow
    maxVal = 0
    maxIdx = 1
    With cht.SeriesCollection(1)
        For i = 1 To .Points.Count
            Set pt = .Points(i)

            ' White border between slices
            With pt.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = WHITE_COLOR
                .Weight = 2
            End With

            ' Shadow
            With pt.Format.Shadow
                .Visible = msoTrue
                .Type = msoShadow21
                .ForeColor.RGB = BLACK_COLOR
                .Transparency = 0.88
                .Blur = 6
                .OffsetX = 0
                .OffsetY = 3
            End With

            ' Track largest slice
            If pt.DataLabel.Text <> "" Then
                ' Use the value to find largest
            End If
        Next i

        ' Explode the first point slightly (usually the largest)
        .Points(1).Explosion = 4
    End With

    ' Delete legend
    cht.HasLegend = False

    ' Data labels: category + percentage, white text on slices
    cht.SeriesCollection(1).HasDataLabels = True
    With cht.SeriesCollection(1).DataLabels
        .ShowCategoryName = True
        .ShowPercentage = True
        .ShowValue = False
        .ShowSeriesName = False

        .Separator = vbLf
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = WHITE_COLOR
        .Position = xlLabelPositionCenter
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With



    ' Leader lines: not available via VBA in PowerPoint.
    ' After running the macro, manually enable them:
    ' Click a pie label > Format Data Labels > Label Options > check "Show Leader Lines"
End Sub

' ============================================================
' SERIES COLOR APPLICATION
' ============================================================
Sub ApplyBarSeriesColors(cht As Chart, slideNum As Long)
    Dim colors() As Long
    Dim i As Long

    Select Case slideNum
        Case 3, 4 ' Infra disruption: Power, Water, IT, Phone
            ReDim colors(1 To 4)
            colors(1) = CLARET_SOFT
            colors(2) = PICTON_SOFT
            colors(3) = NAVY_SOFT
            colors(4) = HARVEST_SOFT

        Case 5 ' Power comparison: Schools, Hospitals
            ReDim colors(1 To 2)
            colors(1) = TAWNY_SOFT
            colors(2) = SHAMROCK_SOFT

        Case 6, 7 ' Employee problems: Not a Problem, Minor, Moderate, Severe
            ReDim colors(1 To 4)
            colors(1) = NAVY_SOFT
            colors(2) = PICTON_SOFT
            colors(3) = HARVEST_SOFT
            colors(4) = CLARET_SOFT

        Case 8 ' Recovery Schools: T1, T2
            ReDim colors(1 To 2)
            colors(1) = NAVY_SOFT
            colors(2) = TAWNY_SOFT

        Case 9 ' Recovery Hospitals: T1, T2
            ReDim colors(1 To 2)
            colors(1) = NAVY_SOFT
            colors(2) = SHAMROCK_SOFT

        Case 12 ' Shelter histogram
            ReDim colors(1 To 1)
            colors(1) = NAVY_SOFT

        Case Else
            Exit Sub
    End Select

    For i = 1 To UBound(colors)
        If i <= cht.SeriesCollection.Count Then
            cht.SeriesCollection(i).Format.Fill.ForeColor.RGB = colors(i)
        End If
    Next i
End Sub

Sub ApplyPieSliceColors(cht As Chart, slideNum As Long)
    Dim i As Long

    Select Case slideNum
        Case 1 ' Damage State Schools: None, Minor, Moderate, Severe, Destroyed
            SetPieColor cht, 1, SHAMROCK_SOFT
            SetPieColor cht, 2, PICTON_SOFT
            SetPieColor cht, 3, HARVEST_SOFT
            SetPieColor cht, 4, TAWNY_SOFT
            SetPieColor cht, 5, CLARET_SOFT

        Case 2 ' Damage State Hospitals: None, Moderate, Severe
            SetPieColor cht, 1, SHAMROCK_SOFT
            SetPieColor cht, 2, HARVEST_SOFT
            SetPieColor cht, 3, TAWNY_SOFT

        Case 10 ' Compounding Schools: COVID, Maria, Other
            SetPieColor cht, 1, MULBERRY_SOFT
            SetPieColor cht, 2, CLARET_SOFT
            SetPieColor cht, 3, HARVEST_SOFT

        Case 11 ' Compounding Hospitals: Maria, COVID, Other
            SetPieColor cht, 1, CLARET_SOFT
            SetPieColor cht, 2, MULBERRY_SOFT
            SetPieColor cht, 3, HARVEST_SOFT
    End Select
End Sub

Sub SetPieColor(cht As Chart, idx As Long, clr As Long)
    If idx <= cht.SeriesCollection(1).Points.Count Then
        cht.SeriesCollection(1).Points(idx).Format.Fill.ForeColor.RGB = clr
    End If
End Sub
