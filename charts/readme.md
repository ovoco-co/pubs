# Chart Styling

VBA that applies the NWIRP publication look to charts, plus the guide it
implements. Two modules, same styling, different host application, and one
PowerShell script that assembles a deck to run the PowerPoint module against.

Everything here works on charts that already exist. It restyles them; it does
not build them from data.

## Scripts

| File | What it does |
|---|---|
| `StyleCharts_PowerPoint.bas` | Styles every chart in the active presentation, keyed by slide name, falling back to slide number. Entry point `StyleAllCharts`, or `StyleDeck` from automation. |
| `StyleCharts_Excel.bas` | Styles every chart on the figure tabs of the active workbook, keyed by sheet name. Entry point `StyleAllCharts`. `IsFigureTab` accepts any tab named `Fig <n>-<n> ...`, plus a legacy list of job-number tabs. |
| `Build-ChartDeck.ps1` | Builds a figures deck from an Excel workbook, one chart per slide, each slide named after its worksheet. |
| `chart-styling-guide.md` | The specification both modules implement: bars, pie slices, axes, gridlines, legends, data labels, the palette, per-figure color assignments, how to paste into Word, and the 508 checklist. |
| `gfx_chart_spec.py` | Writes `chart_spec.json` for `Build-GfxChartDecks.ps1`: one entry per GFX chart (NWIRP 2-1 chapters 4, 5 and 6) naming its workbook, the original image whose proportions set the slide height, the chart kind and layout, and the series colors. Also writes the Figure 4-24 workbook from the native Word chart's data. |
| `Build-GfxChartDecks.ps1` | One deck per GFX ID (`<gfx>a-NWIRP.pptx`) plus a 300 dpi PNG. Styles each chart over COM to the guide - Arial 8 pt throughout, the guide palette and lines - sizes it to its placed size so the 8 pt survives, and pastes it into a 6.5 in slide whose height follows the original figure. Covers the scatter, area, horizontal-bar and multi-panel layouts the `.bas` modules lack (`single`, `side`, `stack`, and `grid` with `grid = [columns, rows]`, whose panels share one plot rectangle so axes line up), and needs no VBA project access. |
| `tighten_source_number.py` | Pulls each deck's source number up to 2 pt under the art and trims the slide height to it, measured from the rendered PNG. |
| `place_figure_pngs.py` | Replaces the picture above each named figure caption in a Word document with its PNG, at natural size, uncropped. Also replaces a native Word chart. |
| `build_grid_figures.py` | GFX 1371 (heatmap) and 1381 (contour) as native PowerPoint grids, since Excel has no chart for either: one rectangle per cell with its value printed, on a sequential Claret ramp, black or white text by contrast so every cell passes 4.5:1. |

Both modules share the same shape: `StyleAllCharts` walks the charts and hands
each to `StyleBarChart`, `StylePieChart`, or `StyleLineChart`, then
`ApplyBarSeriesColors` / `ApplyPieSliceColors` / `ApplyLineSeriesColors` assign
the per-figure colors.

A line chart cannot be routed to `StyleBarChart`: that sets
`ChartGroups(1).GapWidth`, which a line group does not have, so it raises an
error rather than styling the chart badly. `StyleLineChart` is the arm for a
time series - 2.25pt stroke, no markers, no smoothing, no shadow.

## Running one

The `.bas` files are pasted into the host's VBA editor:

1. Open the presentation or workbook
2. Alt+F11, then Insert > Module
3. Paste the whole `.bas` file
4. Close the editor, Alt+F8, run `StyleAllCharts`

`StyleAllCharts` ends in a `MsgBox`. Driving it from COM against a hidden
application therefore hangs forever, because the dialog has nowhere to appear.
Call `StyleDeck` (PowerPoint) instead, which is the same walk without it.

Importing a module through COM also needs *Trust access to the VBA project
object model* in the host's Trust Center. It is off by default in PowerPoint,
and a hand-pasted module does not need it - that is the reason the documented
route is the VBA editor.

## Building a deck

`Build-ChartDeck.ps1` turns a workbook of charts into a deck to style:

```powershell
.\Build-ChartDeck.ps1 -Workbook C:\...\HM_Ch1_Figures.xlsx
```

Every worksheet carrying a chart becomes one slide, at the size the figure will
occupy in Word (6.5 in by default) rather than a presentation slide size, so
what is on the slide is what lands on the page. The chart fills the slide edge
to edge: the caption is written in Word, not here.

Each slide is **named after its worksheet**, and that name is the figure key
`StyleCharts_PowerPoint.bas` reads. Name a slide `Fig 1-1 Counts Stacked` and
it picks up exactly the colors and number format the Excel module gives that
tab, with no slide numbers to keep in step. A deck whose slides are unnamed
(PowerPoint's own `Slide1`, `Slide2`) falls back to the slide-number mapping,
which is how the NWIRP deck is keyed.

If the workbook is already open in Excel the running instance is reused, so the
script does not fight the user for the file lock.

The colors are assigned per figure, so a new deck or workbook needs those
mappings extended before the right chart gets the right color. Line charts are
the exception: an unmapped line chart falls back to the palette in order, so it
comes out styled on the first run and the mapping is only needed when a
specific category has to keep a specific color.

## GFX chart decks

For figures that exist only as images (NWIRP 2-1 chapters 4, 5 and 6), the data is first rebuilt into
one workbook per chart in the chapter's `charts-xlsx\` folder, then:

```powershell
python gfx_chart_spec.py --hm C:\Users\Artist\Documents\HM-261113 --hm5 C:\Users\Artist\Documents\HM-261119
.\Build-GfxChartDecks.ps1 -Spec .\chart_spec.json            # -Only 1317,1329 to rebuild a few
python build_grid_figures.py --hm C:\Users\Artist\Documents\HM-261113
python tighten_source_number.py <chapter>\ppt [<chapter>\ppt ...]   # then re-render the PNGs
python place_figure_pngs.py <chapter>.docx <chapter>\png 4-5=1317 4-6=1318 ...
```

Decks and PNGs land in each chapter folder's `ppt\` and `png\`. The grid script writes only the
deck; render its PNG from PowerPoint at 300 dpi, as the build script does.

`tighten_source_number.py` measures each rendered PNG, moves the source number to 2 pt under the
lowest art, and trims the slide to it (Excel's chart padding otherwise leaves a gap of up to a
quarter inch). `place_figure_pngs.py` swaps the picture above each figure caption for the PNG at its
natural size, with any crop the old picture carried removed.

## Palette

Fills, all 508 compliant:

| Name | Hex | Use |
|---|---|---|
| Navy Soft | `#1A3C8F` | Primary bars, IT/comms, T1 recovery, histogram |
| Tawny Soft | `#E87A30` | Schools T2 bars |
| Shamrock Soft | `#2AB58E` | Hospitals T2 bars, None damage |
| Picton Soft | `#7CC5EE` | Water bars, Minor damage |
| Claret Soft | `#B8335A` | Severe bars, power disruption, Hurricane Maria |
| Harvest Soft | `#F0B840` | Moderate bars, phone, Moderate damage |
| Mulberry Soft | `#D07EAC` | COVID-19 |

Text is `#33333B` on light fills, `#FFFFFF` on dark, and `#6B6B76` for axis and
legend. Lines are `#D8D8DE` for the x axis and leader lines, `#EEEEF2` for
gridlines, and white 2pt between pie slices.

The `.bas` files hold these as `Long` constants in VBA's BGR byte order, which
is why `NAVY_SOFT` reads `&H8F3C1A` and not `1A3C8F`. Reverse the byte pairs
when adding a color.

## Getting a chart into Word

Copy the chart object in PowerPoint, then in Word use Paste Special
(Ctrl+Alt+V) > **Picture (Enhanced Metafile)**, and set the width to 6.5" with
the aspect ratio locked. EMF stays vector, so it prints crisp at any size.

Do not paste as a linked chart or an embedded object.

Figure captions sit **below** the figure on `Caption_Figure`, which is the
opposite of the table convention. They auto-number on Ctrl+A then F9.

## 508

Alt text goes on the chart in PowerPoint *before* pasting, and again on the
pasted image in Word, because the paste does not carry it across. Write the key
data points into the description rather than naming the chart type alone, and
check that nothing is conveyed by color alone: every series needs a legend label
or a direct label.

The guide has the full checklist and worked alt-text examples.

## Where the report content lives

The figure-by-figure recommendations for a given report - which figures are new,
replaced or enhanced, and what each should show - are deliverable content rather
than tooling, so they stay in the internal workspace. For NWIRP 2-3 that is
`NWIRP_2-3_Graphics_Recommendations.md`.
