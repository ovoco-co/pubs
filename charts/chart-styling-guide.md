# Chart Styling Guide

Step-by-step for styling charts in Ch5_Figures.pptx before pasting into the Word document. The Word caption handles the title. The slide is just a workspace. Only the chart gets pasted.

## Setup

1. Open the figures deck in PowerPoint desktop
2. For each slide, delete the title text box, subtitle text box, accent stripe, and bottom rule. Keep only the chart.
3. Resize the chart to fill the entire slide: position 0,0, width 6.5", height 4.5"
4. Name each slide after its figure (Home > Section, or the Selection Pane). The macro reads the slide name as the figure key, the same way the Excel module reads the sheet name, so a named slide gets its colors and number format without anyone tracking slide numbers.

`Build-ChartDeck.ps1` does all four for a deck built from an Excel workbook.

## Bar Charts (Slides 3, 4, 5, 6, 7, 8, 9, 12)

### Bars

1. Double-click the chart to enter edit mode
2. Click any bar to select the series
3. Right-click > Format Data Series
4. Fill: Solid fill, color from the palette at the end of this doc
5. Border: No line
6. Effects > Shadow:
   - Outer, Offset Bottom
   - Color: black, 85% transparent
   - Blur: 4pt, Distance: 2pt, Angle: 90
7. Gap Width: 60% (for grouped bars) or 0% (histogram, slide 12)
8. Overlap: 10% for grouped bars (creates a layered look where bars slightly overlap within each group)
9. Repeat for each series with its color
10. To soften the flat tops, add an inner shadow on each series:
    - Inner, Offset Top
    - Color: white, 50% transparent
    - Blur: 3pt, Distance: 1pt

### Y Axis

1. Click the Y axis numbers
2. Format Axis > Line: No line
3. Tick Marks: Major = None, Minor = None
4. Number format: #,##0
5. Font: Arial 8pt, color #6B6B76
6. Axis title (if present): Arial 8pt, color #6B6B76, not bold

### X Axis

1. Click the X axis labels
2. Format Axis > Line: Solid, color #D8D8DE, 0.5pt
3. Tick Marks: Major = None, Minor = None
4. Font: Arial 8pt, color #6B6B76
5. If labels overlap: Format Axis > Text Options > Custom angle: -30 degrees

### Gridlines

1. Click any horizontal gridline
2. Line: Solid, color #EEEEF2, 0.75pt
3. Delete any vertical gridlines (click and press Delete)

### Chart Area and Plot Area

1. Click the white space outside the bars (chart area)
2. Fill: Solid, white (#FFFFFF)
3. Border: No line
4. Click inside the plot area (the rectangle the bars sit in)
5. Fill: Solid, white (#FFFFFF)
6. Border: No line

### Legend

For charts with 2-4 series, build a custom legend instead of using the default:

1. Delete the default legend (click it, press Delete)
2. Below the chart plot area, draw a text box spanning the chart width
3. For each series, type: [colored circle] Series Name, separated by three spaces
4. To make the colored circles:
   - Insert > Shapes > Oval
   - Hold Shift to make a perfect circle, about 8pt diameter
   - Fill: the series color, no outline
   - Position it inline with the text
5. Space the circle-label pairs evenly across the text box
6. Font: Arial 8pt, color #6B6B76
7. No border or fill on the text box

Alternative (faster but less polished):
1. Keep the default legend
2. Move it to bottom center
3. Font: Arial 9pt, color #6B6B76
4. No border, no fill
5. Right-click legend > Format Legend > Legend Options > leave "Show the legend without overlapping the chart" CHECKED

The legend is one step larger than the axis because on a stacked chart it is
the only key to which band is which. Leave it in the layout: taking it out
gives the plot the whole frame and floats the legend on top of it, which lands
it across the category labels on any chart with a two-line date axis.

Keep the labels short. A legend entry is the category name, not the
measurement: `Completed`, not `Completed - Number of Projects (DIs)`. What is
being measured belongs on the value axis title, where it is said once. Three
long entries will not fit across a 6.5in figure and Excel shrinks them rather
than wrapping.

For single-series charts (slide 12): no legend at all.

### Data Labels (Histogram Only, Slide 12)

1. Click a bar > right-click > Add Data Labels > Above
2. Font: Arial 8pt bold, color #33333B
3. No border, no fill
4. Number format: #,##0

## Stacked Columns (composition over time)

Reach for a stacked column when the series are parts of one whole and the
reading is how the mix changes. The height of the stack is the total, so the
total is present without being plotted as a competing series - which is what
goes wrong if it is added as a fourth bar or a line: it is the envelope, so it
dwarfs its own parts and flattens them into the baseline.

Set up as for a bar chart, with three differences:

1. Gap Width 40%, Overlap 100%. A stacked group has to be fully overlapped or
   the bands step sideways instead of stacking, and the narrower gap is what
   turns a long quarterly series into a readable band of columns
2. No shadow on the series. The drop shadow that lifts a free-standing bar off
   the page reads as grime in the seam between two bands, because the band above
   is sitting directly on it
3. Order the series so the one the reader is tracking sits on the baseline. A
   band that floats between two others has no fixed edge to read against, so
   only the bottom band and the total are actually measurable

Stack order runs bottom to top in series order.

## Line Charts (time series)

A quantity tracked over time is a line, not a bar. Use this for anything with a
date or a quarter on the category axis.

### Lines

1. Double-click the chart to enter edit mode
2. Click a line to select the series
3. Right-click > Format Data Series > Line
4. Solid line, color from the palette, width 2.25pt
5. Marker Options: None
6. Uncheck "Smoothed line". A smoothed line invents values between the real
   data points, which is a misrepresentation in a report
7. No shadow. The drop shadow that gives a bar its lift reads as a blur on a
   2pt line
8. Repeat for each series with its color

### Axes and gridlines

Same as bar charts: Y axis no line, no tick marks, Arial 8pt; horizontal
gridlines #EEEEF2 at 0.75pt; X axis a #D8D8DE 0.5pt baseline with its vertical
gridlines deleted. The Y axis number format follows the quantity - `#,##0` for
counts, `$#,##0.0,,,"B"` to read dollars in billions.

### Legend

A time series usually has more than one line, so use the built-in legend at the
bottom rather than the custom circle legend. Arial 8pt, no fill, no border.

### Assigning colors across a pair of figures

When two figures report the same categories against different measures - counts
in one, dollars in the other - the category has to keep its color in both. The
series order often differs between them, so the color cannot simply follow the
series index. Fix the color to the category and map it per figure:

| Category | Color |
|---|---|
| Completed | Shamrock Soft #2AB58E |
| Other statuses | Picton Soft #7CC5EE |
| Planning | Tawny Soft #E87A30 |
| Total | Navy Soft #1A3C8F |

`ApplyLineSeriesColors` in both .bas files encodes exactly this: a default
palette rotation for any unmapped figure, then a per-figure override where the
order needs correcting.

Note that this is the one place the line styler behaves differently from the
bar and pie stylers. An unmapped bar or pie chart is left alone; an unmapped
line chart still gets the palette in order, so a new figure comes out styled
the first time the macro runs rather than in Excel's defaults.

## Pie Charts (Slides 1, 2, 10, 11)

### Slices

For each slice (click the pie, then click the individual slice):

1. Fill: Solid fill, color from palette
2. Border: Solid, white (#FFFFFF), 2pt (clean separation between slices)
3. Effects > Shadow:
   - Outer, Offset Bottom
   - Color: black, 88% transparent
   - Blur: 6pt, Distance: 3pt, Angle: 90
4. For the largest slice only: drag it outward slightly (3-5% explosion) to draw the eye. Or Format Data Point > Series Options > Point Explosion: 4%

### Chart Area and Plot Area

1. Chart area: Solid fill, white (#FFFFFF), no border
2. Plot area: Solid fill, white (#FFFFFF), no border

### Labels

1. Delete the default legend
2. Right-click pie > Add Data Labels
3. Label Options:
   - Check: Category Name, Percentage
   - Uncheck: Value, Series Name
   - Separator: (New Line)
4. Font: Arial 9pt, color #33333B
5. Check "Show Leader Lines"
6. Click individual labels and drag them outside the slices into clean positions. Arrange them so they don't overlap and leader lines don't cross.
7. Format leader lines: solid, #D8D8DE, 0.5pt

For labels near dark slices (Claret, Navy, Mulberry fills):
- If the label overlaps the dark slice, change that label's font color to white (#FFFFFF)
- If the label is fully outside the slice connected by a leader line, keep it dark (#33333B)

### Custom Legend (alternative to direct labels for simple pies)

If a pie has only 2-3 slices and direct labels feel cluttered:

1. Delete data labels
2. Below the pie, draw a horizontal text box
3. For each slice: colored circle (8pt, filled with slice color, no outline) followed by "Category Name (XX%)"
4. Space evenly
5. Font: Arial 9pt, color #33333B
6. This reads cleaner than labels on tiny slices

## Color Palette

### Bar and Pie Fills

| Name | Hex | Use |
|---|---|---|
| Navy Soft | #1A3C8F | Primary bars, IT/comms, T1 recovery, histogram |
| Tawny Soft | #E87A30 | Schools T2 bars |
| Shamrock Soft | #2AB58E | Hospitals T2 bars, None damage |
| Picton Soft | #7CC5EE | Water bars, Minor damage |
| Claret Soft | #B8335A | Severe bars, power disruption, Hurricane Maria |
| Harvest Soft | #F0B840 | Moderate bars, phone, Moderate damage |
| Mulberry Soft | #D07EAC | COVID-19 |

### Text

| Context | Color |
|---|---|
| Data labels on light fills | #33333B |
| Data labels on dark fills | #FFFFFF |
| Axis labels, legend text | #6B6B76 |

### Lines

| Element | Color | Weight |
|---|---|---|
| X axis | #D8D8DE | 0.5pt |
| Gridlines | #EEEEF2 | 0.75pt |
| Pie slice borders | #FFFFFF | 2pt |
| Leader lines | #D8D8DE | 0.5pt |

## Per-Slide Colors

### Slide 1: Damage State, Schools (pie)
| Slice | Color |
|---|---|
| None | #2AB58E |
| Minor | #7CC5EE |
| Moderate | #F0B840 |
| Severe | #E87A30 |
| Destroyed | #B8335A |

### Slide 2: Damage State, Hospitals (pie)
| Slice | Color |
|---|---|
| None | #2AB58E |
| Moderate | #F0B840 |
| Severe | #E87A30 |

### Slides 3-4: Infrastructure Disruption (grouped bars)
| Series | Color |
|---|---|
| Power | #B8335A |
| Water | #7CC5EE |
| IT/Comms | #1A3C8F |
| Phone | #F0B840 |

### Slide 5: Power Outage Comparison (grouped bars)
| Series | Color |
|---|---|
| Schools | #E87A30 |
| Hospitals | #2AB58E |

### Slides 6-7: Employee Problems (grouped bars)
| Series | Color |
|---|---|
| Not a Problem | #1A3C8F |
| Minor | #7CC5EE |
| Moderate | #F0B840 |
| Severe | #B8335A |

### Slides 8-9: Recovery (grouped bars)
| Series | Color |
|---|---|
| 1 Year Post | #1A3C8F |
| 2 Years Post (Schools, slide 8) | #E87A30 |
| 2 Years Post (Hospitals, slide 9) | #2AB58E |

### Slide 10: Compounding, Schools (pie)
| Slice | Color |
|---|---|
| COVID-19 | #D07EAC |
| Hurricane Maria | #B8335A |
| Other | #F0B840 |

### Slide 11: Compounding, Hospitals (pie)
| Slice | Color |
|---|---|
| Hurricane Maria | #B8335A |
| COVID-19 | #D07EAC |
| Other | #F0B840 |

### Slide 12: Shelter Duration (histogram)
| Series | Color |
|---|---|
| Schools | #1A3C8F |

## Pasting into Word

1. In PowerPoint, click the styled chart (not the slide, the chart object)
2. Ctrl+C
3. In Word, click at the figure placeholder text "[Insert chart from Excel F5-xx]"
4. Delete the placeholder text
5. Paste Special (Ctrl+Alt+V) > Picture (Enhanced Metafile)
   - EMF keeps it vector so it prints crisp at any size
   - Do not paste as a linked chart or embedded object
6. Right-click the pasted image > Size and Position > Width: 6.5"
   - Lock aspect ratio should be checked
7. The figure caption above it (Caption_Figure style) auto-numbers when you press Ctrl+A then F9

## 508 Compliance

Before pasting into Word, for each chart in PowerPoint:

1. Right-click chart > Edit Alt Text
2. Write a description including key data points:
   - "Pie chart. Damage state for schools: Severe 47%, Moderate 19%, None 20%, Minor 11%, Destroyed 3%."
   - "Grouped bar chart. Employee problems for schools by severity. Transportation: 81 not a problem, 84 minor, 75 moderate, 33 severe."
3. After pasting as EMF in Word, right-click the image > Edit Alt Text and paste the same description
4. Verify: no information is conveyed by color alone (every series has a legend label or direct label with text)
