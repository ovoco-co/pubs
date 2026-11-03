# NIST HM Color Palette

The master palette for the NCSTAR and NWIRP document sets. Every color used in
a chart, a table fill, or a chapter header comes from here. Nothing in a
deliverable should carry a color that is not on this page.

Hex values are written without the hash, because that is how OOXML wants them
(`w:fill="CCECE3"`). VBA wants the same six digits reversed into BGR - see
`charts/chart-styling-guide.md`.

## Primary

| Name | Hex | RGB |
|---|---|---|
| Resolution Blue | 00257A | 0, 37, 122 |
| Tawny | D55E00 | 213, 94, 0 |
| Shamrock Green | 009E73 | 0, 158, 115 |
| Claret | 990B38 | 153, 11, 56 |
| Picton Blue | 56B4E9 | 86, 180, 233 |
| Harvest Gold | E69F00 | 230, 159, 0 |

## Secondary and accent

| Name | Hex | RGB |
|---|---|---|
| Dark Moss Green | 365022 | 54, 80, 34 |
| Mulberry | C15C94 | 193, 92, 148 |
| Lion | CDA572 | 205, 165, 114 |
| Maize | F0E442 | 240, 228, 66 |

## Tint scales

Each primary has a tint ramp, for sequential data, shading, and emphasis. The
percentage is how far the color has been lightened toward white, so 80% is a
pale wash and Base is full strength.

| Tint | Resolution Blue | Shamrock Green | Tawny | Claret | Picton Blue |
|---|---|---|---|---|---|
| Base | 00257A | 009E73 | D55E00 | 990B38 | 56B4E9 |
| 20% | 335195 | 33B18F | DD7E33 | AD3C60 | 78C3ED |
| 40% | 667CAF | 66C5AB | E69E66 | C26D88 | 9AD2F2 |
| 60% | 99A8CA | 99D8C7 | EEBF99 | D69DB0 | BCE1F6 |
| 80% | CCD3E4 | CCECE3 | F6DFCC | EBCED7 | DEF0FB |

The 80% row is the one table shading uses. It is light enough to sit under 9pt
text and still darker than the D9D9D9 banding stripe, so a banner row reads as
a heading rather than as another banded row. The 90% tints (Shamrock E6F5F1,
Tawny FBF0E6) are lighter than the banding and only work in a table with
banding turned off, which is what `Ch5_Tables_v2.docx` did.

## Neutrals

| Name | Hex | Use |
|---|---|---|
| White | FFFFFF | Ocean fill, page backgrounds |
| Alternating Row Gray | D9D9D9 | 15% gray, table row banding |
| Light Gray | D5D5D5 | Minimal hazard areas on maps |
| Background Gray | E5E5E5 | Map backgrounds |
| Mid Gray | 9D9D9D | Secondary boundaries |
| Dark Gray | 4A4A4A | Text, primary boundaries |
| Dark Text | 333333 | Standard label color |

## Chapter header gradients

NWIRP: 37627E, 457B9D, 7DA3BA, C7D7E2

NCSTAR: 02257A, 355195, 677CAF, CCD3E4

NCSTAR's gradient starts at 02257A rather than the palette's 00257A. Both
values are in circulation; `Set-NISTTheme.ps1` treats 00257A as canonical for
theme accent1 and the `NIST-Data` header row, and 02257A is the gradient's own
first stop.

## Data color assignments

These are semantic. Once a color has been assigned to a category it means that
category everywhere in the document set - in a chart, in a table fill, in a map
legend. Do not reuse it for anything else in the same document.

### Schools and hospitals

| Institution | Color | Hex | 80% tint for table fills |
|---|---|---|---|
| Schools | Tawny | D55E00 | F6DFCC |
| Hospitals | Shamrock Green | 009E73 | CCECE3 |

This is the pairing that carries through a whole chapter, so a reader can tell
at a glance which population a table or a figure is about. In a table it shows
up in two places, and both are driven by the caption text:

| Element | Strength | Schools | Hospitals | Neither | Script |
|---|---|---|---|---|---|
| Accent bar, left edge of every row, 4pt | Base | D55E00 | 009E73 | 00257A | `word/Set-TableAccentBar.ps1 -ColorRules` |
| Group-label banner rows | 80% tint | F6DFCC | CCECE3 | CCD3E4 | `word/Set-TableBannerRows.ps1 -FillRules` |

The accent bar is the primary carrier. It runs the full height of the table on
the first cell of every row, header included but footnote rows excluded, and
it does not depend on the table having banner rows, so it still codes the
small summary tables that have none. The footnote row sits outside the data
grid; a spine through it reads as the table continuing past its own last row. Draw it at Base strength: a 4pt rule in an 80% tint reads as a
printing artifact rather than as a deliberate mark.

The banner tint is the secondary carrier, for the tables that do have group
labels. The banner green already in use is CCECE3, which is exactly Shamrock
Green at 80%, so hospital tables need no change; schools take the matching
Tawny tint F6DFCC.

A table that describes neither population takes Resolution Blue, which reads
as structural rather than as a third group.

### Timeline categories

| Category | Color | Hex | Examples |
|---|---|---|---|
| Major hurricanes | Picton Blue | 56B4E9 | Maria, Fiona |
| Contextual disruptions | Claret | 990B38 | Earthquakes, COVID-19, LUMA takeover |
| Project milestones | Resolution Blue | 00257A | Contract award, kickoff, report delivery |
| Data collection | Tawny | D55E00 | Surveys, pilots, interviews |
| Government actions | Shamrock Green | 009E73 | Shelter openings, emergency declarations |
| Communication channels | Harvest Gold | E69F00 | NWR, WEA, press, Twitter |

### NCSTAR 4-3 survey charts

| Element | Color | Hex |
|---|---|---|
| TV | Resolution Blue | 00257A |
| Radio | Shamrock Green | 009E73 |
| Social media | Picton Blue | 56B4E9 |
| Cell phone | Harvest Gold | E69F00 |
| Internet | Mulberry | C15C94 |
| Tumba Coco | Tawny | D55E00 |
| Evacuated | Tawny | D55E00 |
| Did not evacuate | Resolution Blue | 00257A |
| Yes (general) | Shamrock Green | 009E73 |

## A note on the "Soft" chart colors

`charts/chart-styling-guide.md` lists a parallel set - Navy Soft 1A3C8F, Tawny
Soft E87A30, Shamrock Soft 2AB58E, Picton Soft 7CC5EE, Claret Soft B8335A,
Harvest Soft F0B840, Mulberry Soft D07EAC. These are desaturated variants used
for large filled areas in charts, where the primaries read as too heavy. They
are not substitutes for the primaries and must not be used in tables. When a
chart and a table both refer to the same category, the chart may use the Soft
variant and the table the 80% tint; the hue is what carries the meaning.
