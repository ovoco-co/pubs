' StyleCharts VBA Macro for PowerPoint
' Applies NWIRP publication styling to all charts in the active presentation.
'
' To use:
' 1. Open the figures deck in PowerPoint
' 2. Alt+F11 to open VBA editor
' 3. Insert > Module
' 4. Paste this entire file
' 5. Close VBA editor
' 6. Alt+F8 > StyleAllCharts > Run
'
' The macro styles every slide automatically. Run once, then manually drag pie
' labels into clean positions and adjust any overlapping text.
'
' Figures are keyed two ways. A slide NAME that matches a figure key is used
' first, which is how the Excel module works and how a deck built by
' Build-ChartDeck.ps1 identifies its figures. A deck whose slides are unnamed
' falls back to the slide NUMBER, which is how the NWIRP deck is mapped.

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
' TYPE SIZE
' ============================================================
' Chart text does NOT scale when the shape is resized. An 8pt axis authored
' for a 6.5in figure is still 8pt after the chart is stretched across a wide
' slide, which is where a legend stops being readable. Size the type from the
' chart's actual width against the 6.5in Word text column these figures are
' eventually pasted into, so a deck authored at any slide size comes out with
' the same apparent type size.
Private Const REFERENCE_WIDTH_PT As Single = 468#   ' 6.5 in

Private Const AXIS_PT As Single = 8#
Private Const LEGEND_PT As Single = 9#              ' one step up from the axis:
                                                    ' on a stacked chart the
                                                    ' legend is the only key to
                                                    ' which band is which
Private Const DATALABEL_PT As Single = 8#
Private Const PIELABEL_PT As Single = 10#

' ============================================================
' MAIN ENTRY POINT
' ============================================================
Sub StyleAllCharts()
    StyleDeck
    MsgBox "Styling complete. Manually adjust pie labels for clean positioning.", vbInformation
End Sub

' The same walk without the MsgBox. Call this one from COM automation: a
' hidden application has nowhere to show a dialog and hangs forever on it.
Sub StyleDeck()
    Dim sld As Slide
    Dim shp As Shape
    Dim cht As Chart
    Dim slideNum As Long
    Dim figKey As String

    For Each sld In ActivePresentation.Slides
        slideNum = sld.SlideIndex
        figKey = FigureKey(sld)

        ' Style the chart (leave size and position as-is)
        For Each shp In sld.Shapes
            If shp.HasChart Then
                Set cht = shp.Chart

                ' Determine chart type and apply styling
                Select Case cht.ChartType
                    Case xlPie, xl3DPie, xlPieExploded, xl3DPieExploded
                        StylePieChart cht, slideNum, figKey
                    Case xlLine, xlLineMarkers, xlLineStacked, xlLineMarkersStacked, _
                         xlLineStacked100, xlLineMarkersStacked100
                        StyleLineChart cht, slideNum, figKey
                    Case xlColumnClustered, xlBarClustered, _
                         xlColumnStacked, xlBarStacked, _
                         xlColumnStacked100, xlBarStacked100
                        StyleBarChart cht, slideNum, figKey
                End Select
            End If
        Next shp
    Next sld
End Sub

' A slide carries its figure key in its name. PowerPoint names slides
' "Slide1", "Slide2" by default, and those are not figure keys: reject them so
' an unnamed deck falls through to the slide-number mapping.
Function FigureKey(sld As Slide) As String
    Dim nm As String

    nm = ""
    On Error Resume Next
    nm = sld.Name
    On Error GoTo 0

    If Len(nm) >= 5 Then
        If LCase$(Left$(nm, 5)) = "slide" Then nm = ""
    End If

    FigureKey = nm
End Function

Function ScaledPt(cht As Chart, basePt As Single) As Single
    Dim w As Single

    w = 0
    On Error Resume Next
    w = cht.Parent.Width      ' the Shape the chart sits in, in points
    On Error GoTo 0

    If w <= 0 Then
        ScaledPt = basePt
        Exit Function
    End If

    ScaledPt = basePt * (w / REFERENCE_WIDTH_PT)

    ' Never smaller than the authored size, and never more than about double,
    ' past which the type starts crowding the plot instead of labelling it.
    If ScaledPt < basePt Then ScaledPt = basePt
    If ScaledPt > basePt * 2.25 Then ScaledPt = basePt * 2.25
End Function

' ============================================================
' BAR CHART STYLING
' ============================================================
Sub StyleBarChart(cht As Chart, slideNum As Long, figKey As String)
    Dim s As Series
    Dim i As Long
    Dim isHistogram As Boolean
    Dim isStacked As Boolean

    isHistogram = (slideNum = 12 And Len(figKey) = 0)

    Select Case cht.ChartType
        Case xlColumnStacked, xlBarStacked, xlColumnStacked100, xlBarStacked100
            isStacked = True
    End Select

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
    ApplyBarSeriesColors cht, slideNum, figKey

    ' Series formatting: shadow, gap, overlap
    For i = 1 To cht.SeriesCollection.Count
        Set s = cht.SeriesCollection(i)

        ' No border on bars
        s.Format.Line.Visible = msoFalse

        If isStacked Then
            ' The drop shadow that lifts a free-standing bar off the page reads
            ' as grime in the seam between two stacked bands, because the band
            ' above is sitting directly on it.
            s.Format.Shadow.Visible = msoFalse
        Else
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
        End If
    Next i

    ' Gap and overlap
    With cht.ChartGroups(1)
        If isStacked Then
            ' A stacked group has to be fully overlapped or the bands step
            ' sideways instead of stacking. The narrower gap is what turns a
            ' 19-quarter series into a readable band of columns.
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

    ' Y Axis (Value Axis)
    With cht.Axes(xlValue)
        .Format.Line.Visible = msoFalse
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = ScaledPt(cht, AXIS_PT)
        .TickLabels.Font.Color = SLATE
        .TickLabels.NumberFormat = AxisValueFormat(slideNum, figKey)

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
                .Size = ScaledPt(cht, AXIS_PT)
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
        .TickLabels.Font.Size = ScaledPt(cht, AXIS_PT)
        .TickLabels.Font.Color = SLATE

        ' Delete vertical gridlines
        If .HasMajorGridlines Then .HasMajorGridlines = False
    End With

    ' Legend
    StyleLegend cht

    ' Data labels (histogram only)
    If isHistogram Then
        cht.SeriesCollection(1).HasDataLabels = True
        With cht.SeriesCollection(1).DataLabels
            .Font.Name = "Arial"
            .Font.Size = ScaledPt(cht, DATALABEL_PT)
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
Sub StylePieChart(cht As Chart, slideNum As Long, figKey As String)
    Dim pt As Point
    Dim i As Long

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
    ApplyPieSliceColors cht, slideNum, figKey

    ' Style each slice: white border, shadow
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
        .Font.Size = ScaledPt(cht, PIELABEL_PT)
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
' LINE CHART STYLING
' ============================================================
' A time series is a line, and the Select Case above used to have no arm for
' one: a line chart fell straight through and came out unstyled. It cannot be
' routed to StyleBarChart either, because that sets ChartGroups(1).GapWidth,
' which a line group does not have.
Sub StyleLineChart(cht As Chart, slideNum As Long, figKey As String)
    Dim s As Series
    Dim i As Long

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

    ApplyLineSeriesColors cht, slideNum, figKey

    ' Y Axis (Value Axis)
    With cht.Axes(xlValue)
        .Format.Line.Visible = msoFalse
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = ScaledPt(cht, AXIS_PT)
        .TickLabels.Font.Color = SLATE
        .TickLabels.NumberFormat = AxisValueFormat(slideNum, figKey)

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
                .Size = ScaledPt(cht, AXIS_PT)
                .Fill.ForeColor.RGB = SLATE
                .Bold = msoFalse
            End With
        End If
    End With

    ' X Axis (Category Axis). A time axis keeps its baseline and loses its
    ' gridlines: the horizontal ones already carry the reading.
    With cht.Axes(xlCategory)
        With .Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = SILVER_LINE
            .Weight = 0.5
        End With
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = ScaledPt(cht, AXIS_PT)
        .TickLabels.Font.Color = SLATE

        If .HasMajorGridlines Then .HasMajorGridlines = False
    End With

    ' Legend
    StyleLegend cht
End Sub

' ============================================================
' LEGEND
' ============================================================
' One legend treatment for every chart type that keeps one. A single-series
' chart names its quantity on the axis, so the legend is only noise.
Sub StyleLegend(cht As Chart)
    If cht.SeriesCollection.Count <= 1 Then
        cht.HasLegend = False
        Exit Sub
    End If

    cht.HasLegend = True
    With cht.Legend
        .Position = xlLegendPositionBottom

        ' Keep the legend in the layout. Taking it out gives the plot the whole
        ' frame and then floats the legend on top of it, which lands the legend
        ' squarely across the category labels on any chart with a deep axis.
        ' The guide used to call for False plus a manual plot-area resize; this
        ' is the same result without the manual step.
        .IncludeInLayout = True
        .Format.TextFrame2.TextRange.Font.Name = "Arial"
        .Format.TextFrame2.TextRange.Font.Size = ScaledPt(cht, LEGEND_PT)
        .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = SLATE
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With
End Sub

' ============================================================
' AXIS NUMBER FORMAT
' ============================================================
Function AxisValueFormat(slideNum As Long, figKey As String) As String
    Select Case figKey
        Case "Fig 1-2 Dollars Stacked", "Fig 1-2 Dollars Line"
            ' Billions, so the axis reads $1.5B rather than 1500000000
            AxisValueFormat = "$#,##0.0,,,""B"""
            Exit Function
        Case "Fig 1-1 Counts Stacked", "Fig 1-1 Counts Line"
            AxisValueFormat = "#,##0"
            Exit Function
    End Select

    Select Case slideNum
        Case 2 ' COR3 obligations in dollars
            AxisValueFormat = "$#,##0.0,,,""B"""
        Case Else
            AxisValueFormat = "#,##0"
    End Select
End Function

' ============================================================
' SERIES COLOR APPLICATION
' ============================================================
Sub ApplyBarSeriesColors(cht As Chart, slideNum As Long, figKey As String)
    Dim colors() As Long
    Dim i As Long

    ' Figure keys first. The COR3 stacked pair carries the same three
    ' categories in the same order in both figures, so the colour is fixed to
    ' the category: Completed shamrock, the derived middle band picton,
    ' Planning tawny.
    Select Case figKey
        Case "Fig 1-1 Counts Stacked", "Fig 1-2 Dollars Stacked"
            ReDim colors(1 To 3)
            colors(1) = SHAMROCK_SOFT
            colors(2) = PICTON_SOFT
            colors(3) = TAWNY_SOFT
            GoTo ApplyColors
    End Select

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

ApplyColors:
    For i = 1 To UBound(colors)
        If i <= cht.SeriesCollection.Count Then
            cht.SeriesCollection(i).Format.Fill.ForeColor.RGB = colors(i)
        End If
    Next i
End Sub

Sub ApplyLineSeriesColors(cht As Chart, slideNum As Long, figKey As String)
    Dim colors(1 To 7) As Long
    Dim i As Long

    ' The palette in order, so an unmapped figure still comes out styled
    ' instead of keeping PowerPoint's defaults. That is the deliberate
    ' difference from ApplyBarSeriesColors, which Exit Subs on an unmapped
    ' figure and leaves the chart alone.
    colors(1) = NAVY_SOFT
    colors(2) = TAWNY_SOFT
    colors(3) = SHAMROCK_SOFT
    colors(4) = PICTON_SOFT
    colors(5) = CLARET_SOFT
    colors(6) = HARVEST_SOFT
    colors(7) = MULBERRY_SOFT

    ' Per-figure overrides, so a category keeps its colour across the pair of
    ' figures that report it in counts and in dollars.
    Select Case figKey
        Case "Fig 1-1 Counts Line", "Fig 1-2 Dollars Line"
            colors(1) = SHAMROCK_SOFT   ' Completed
            colors(2) = TAWNY_SOFT      ' Planning
        Case ""
            Select Case slideNum
                Case 1 ' COR3 project counts: Total, Completed, Planning
                    colors(1) = NAVY_SOFT
                    colors(2) = SHAMROCK_SOFT
                    colors(3) = TAWNY_SOFT

                Case 2 ' COR3 project dollars: Total, Planning, Completed
                    colors(1) = NAVY_SOFT
                    colors(2) = TAWNY_SOFT
                    colors(3) = SHAMROCK_SOFT
            End Select
    End Select

    ' A line carries its colour on the stroke, not on a fill.
    For i = 1 To cht.SeriesCollection.Count
        If i <= 7 Then
            cht.SeriesCollection(i).Format.Line.ForeColor.RGB = colors(i)
        End If
    Next i
End Sub

Sub ApplyPieSliceColors(cht As Chart, slideNum As Long, figKey As String)
    If Len(figKey) > 0 Then Exit Sub

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
