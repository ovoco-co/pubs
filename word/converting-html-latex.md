# Converting HTML and LaTeX into the NIST Template

How to turn HTML, LaTeX, and Markdown source into Word documents that carry the NIST template. Two scripts do the work: `Build-NISTReferenceDoc.ps1` (run once per template revision) and `Convert-ToNISTWord.ps1` (run per job).

## Before you start

You need pandoc. It is the only external binary in the toolkit, and only these two scripts use it. Everything else in `publishing/word/` runs without it.

```powershell
winget install JohnMacFarlane.Pandoc
pandoc --version
```

Pandoc is installed on the Windows side, not inside WSL, so run these scripts from Windows PowerShell like the rest of the toolkit:

```powershell
cd \\wsl$\Ubuntu-24.04\home\artist\ovoco\tools\publishing\word
```

Do not copy the scripts into your document folder. They resolve `Build-NISTDocuments.ps1` and the `OoXml/` module by their own location.

You also need a NIST template file (`.docx` or `.dotx`) to build the reference document from.

## Step one: build the reference document

Do this once, and again whenever the template changes.

```powershell
.\Build-NISTReferenceDoc.ps1 -Template "$env:USERPROFILE\Documents\260806\HM-NCSTAR-Template.docx" -ApplyStyleArchitecture
```

This writes `NIST-reference.docx` next to the scripts, which is exactly where `Convert-ToNISTWord.ps1` looks for it by default. Pass `-Output` to put it somewhere else.

The current template is `Documents\260806\HM-NCSTAR-Template.docx`. Older copies of that file name exist under `Documents\zz\framescripts\` and elsewhere; they are one architecture generation behind (140 styles instead of 245, no Sidebar Table style). Check the folder date before you point at one.

`-ApplyStyleArchitecture` runs `Set-NISTTemplateStyles.ps1` on the copy before the pandoc styles are added, so the reference document is current and the Recommended-pane curation does not promote the pandoc styles into the authors' style list. It is safe to pass on a current template and necessary on an older one.

```powershell
.\Build-NISTReferenceDoc.ps1 -Template "NCSTAR.dotx" -ApplyStyleArchitecture -Force
```

The script reports what it found and what it added:

```
Already in the template (left untouched):
  Block Text         -> BlockText
  Figure             -> Figure

Added:
  Source Code        id SourceCode    basedOn BodyText    code blocks (no NIST equivalent, kept)
  Image Caption      id ImageCaption  basedOn Caption     figure captions
  ...
```

### Why this step exists

Pandoc resolves styles from the reference document **by name**, not by styleId. It looks for names like `Source Code`, `Block Text`, `Image Caption`, and `First Paragraph`. Any name it cannot find, it invents with its own look, which then has to be cleaned out of every converted file.

`Build-NISTReferenceDoc.ps1` adds a definition for each of those 17 names, based on the closest NIST style so it inherits the right font, color, and spacing. Names your template already defines are left alone, so the template stays the source of truth and re-running is safe.

The theme, the section properties (page size and margins), `numbering.xml`, and the header and footer parts come along automatically, because the reference document is a copy of your template.

## Step two: convert

One file:

```powershell
.\Convert-ToNISTWord.ps1 -File "C:\src\chapter-03.html"
```

A whole folder, into a dedicated output folder:

```powershell
.\Convert-ToNISTWord.ps1 -Folder "C:\src\report" -OutputFolder "C:\out\report"
```

Accepted source types: `.html`, `.htm`, `.xhtml`, `.tex`, `.md`, `.markdown`, `.docx`.

A `.docx` source is read back against the reference document, which retargets an existing pandoc or Quarto Word file at the NIST template while keeping its tables, merged cells, and images. Pass it with `-File`; `-Folder` skips Word files, because a folder scan would convert its own output on a re-run and would overwrite the sources in place when `-OutputFolder` is left unset. The script refuses outright if the output path resolves back to the source.

Point `-OutputFolder` at a dedicated folder. The last stage runs the NIST build on **every** Word file in the output folder, not just the ones it converted, so a folder that already holds unrelated documents will see them reformatted too. Left unset, output lands beside the source.

### What the five stages do

**Stage zero, inline formatting (HTML sources only).** Rewrites CSS-only inline formatting into real tags before pandoc sees it, using `Repair-HtmlInlineFormatting.ps1`. Generators like R flextable emit no `<strong>`, `<em>`, or `<sup>` at all: every run of text is a `<span class="cl-xxxx">` and the bold, italic, and superscript live only in the generated stylesheet. Pandoc reads tags, not CSS, so without this the emphasis is flattened and nothing warns you. The source file is never touched; the repair runs on a working copy in the source folder, so relative image paths still resolve, and the copy is deleted afterwards. `-SkipInlineRepair` turns the stage off for HTML that already carries real tags.

**Stage one, pandoc.** Converts each source file to `.docx` against the reference document, so the output starts with the NIST theme, page setup, headers and footers, numbering, and styles. Images are embedded, and `--extract-media` also drops the originals in a `media` folder beside the output. LaTeX math becomes native Word equations (OMML).

**Stage two, remap.** Folds pandoc's styles into NIST styles, assigns the `NIST-Data` table style to every table, and deletes the pandoc style definitions that nothing references any more.

**Stage three, build.** Runs `Build-NISTDocuments.ps1 -Clean` on the output folder: theme, styles, page layout (1" margins), table look, 3-column width normalization, typography, and font de-embedding.

**Stage four, tables.** Finishes what the generic pipeline cannot, because the pipeline's settings are tuned for hand-authored NCSTAR documents rather than converted statistical output. Three things happen per file, and a fourth on request:

- **Table Style Options** are set with `Set-TableLook.ps1`: Header Row on, Banded Rows on, everything else off. Each option is a parameter on `Convert-ToNISTWord.ps1`. Leave **First Column** off unless you want it: turning it on is what bolds and shades column 1. Repeat-header-row is separate and stays on regardless. Every table is also pinned to the width its own grid adds up to, with a fixed layout: a converted table arrives as `tblW auto`, so Word sizes it from its contents and two tables with an identical grid still render at different widths. `-TableWidthInches` pins them all to one width instead, rescaling the columns in proportion; `-AutoFitTables` leaves Word to it. On top of that the statistical tables - the ones whose header row reads `Characteristic / N / %` - get the exact 2.7" / 0.6" / 0.6" (3.9" overall), so they are all identical down the page. The two small columns are equal and sized off their own content; column 1 is sized off the **green banner rows**, which span all three columns, so the table has to be wide enough for the longest group label or it wraps to two lines. `-ColumnWidthsInches` and `-MatchHeader` change that pairing, `-ColumnWidthsInches @()` turns it off, and any table that does not match the header keeps its own grid.
- **Cell text** moves onto `TableText` with `Set-TableParagraphStyle.ps1`, which rewrites `w:pStyle` and nothing else, so bold, superscripts, subscripts, and `w:ind` all survive. `-TableTextStyle ''` leaves cell paragraph styles alone.
- **Banner rows** - the full-width merged rows that title a group (`Ownership (n = 297)`) rather than carry data - are shaded `CCECE3` and bolded by `Set-TableBannerRows.ps1`. It runs after the footnote merge, because that is what separates the two kinds of full-width merged row: the notes are on `TableandFigureFootnotes` by then, so whatever is still unclaimed is a group label. `-BannerFill` sets the color, `-SkipTableBanners` turns it off.
- **Bold in named columns**, with `-UnboldColumns`, is stripped from the body rows, header row excluded. Off by default: the bold in column 1 is normally a banner row, which is meant to be bold. Converted statistical tables carry their group labels (`Ownership (n = 297)`) as bold direct run formatting, which no Table Style Option can turn off, because it is not the style doing it. `-UnboldColumns` takes the column numbers, or an empty array to leave all bold alone.
- **Row indentation** is rebuilt with `Restore-TableIndents.ps1`. With a `.docx` source the levels are read from the original's leading tabs; from HTML the target's own leading spaces are read. `-SkipTableIndents` turns it off.
- **Table captions** are made from the header sitting above each table by `Set-NISTTableCaptions.ps1`, unless `-SkipTableCaptions`. The header moves onto `Caption_Table` and picks up the NIST number ahead of its existing title text: the word Table, a non-breaking space, the chapter number, an en dash, a `SEQ Table` field, a period, and two spaces. Only the table number is a field, so the set renumbers on Ctrl+A then F9. `-CaptionFromStyle` says which styles the headers are on and takes a list, which converted HTML usually needs, since the headers are rarely all one level and an HTML `<caption>` element arrives on `Caption`. The chapter number is a `STYLEREF` field by default, so each caption reads the chapter it actually sits under rather than a number baked in at conversion time; passing `-CaptionChapter` is what asks for a literal instead. A standalone extract with no Heading 1 above it shows a field error until it is merged into a real chapter. Any number the source already put in the title is dropped, so the caption reads `Table 1-1.  Sample Characteristics` and not `Table 1-1.  1 Sample Characteristics`; it is only dropped when it matches the number the caption is being given, so a title like `2020 Enrollment Data` keeps its 2020. `-KeepSourceNumbers` turns that off.
- **Table footnotes** are collapsed, unless `-SkipTableFootnotes`, into one full-width unshaded row per table by `Merge-TableFootnotes.ps1`. A converter emits each note as its own row, and because those rows are full-width merges they pick up the banding stripe, or a banner fill if something upstream mistook a merged row for a banner. The merged row carries `w:fill="auto"` and a zeroed `w:cnfStyle`, so no Table Style Option can paint it wherever it lands in the stripe pattern, and the notes go on `TableandFigureFootnotes` with the bold and the indent removed. Superscript markers survive. From an HTML source the notes are identified from the original's `<tfoot>` blocks, which is exact; otherwise by row shape. That structural rule also matches a table whose last row is a banner or a merged total, so on a new document set run the script alone with `-ListOnly` once and read the listing.

The order inside stage four matters: the footnote merge rewrites whole rows, so it runs after the cell style and the indents have settled, and the captions come last because the caption style only exists once stage three has applied the style architecture.

Stage two runs before stage three on purpose. `Set-NISTStyleArchitecture.ps1` deletes styles it does not recognize but never edits `document.xml`, so a paragraph still pointing at a deleted pandoc style would quietly fall back to Normal. The remap rewrites those references first. Stage four runs after stage three for the same kind of reason: `Build-NISTDocuments.ps1` runs `Set-TableLook.ps1` with its own defaults, so the converted-table options have to be written afterwards or they are overwritten.

## Step three: finish the document

The conversion does not do everything. After it runs:

- **Headers and footers.** Run `Set-NISTHeaderFooterContent.ps1` per file, with `-FrontMatter` on the front-matter file and without it on chapters.
- **Document title.** Run `Set-DocTitle.ps1`, which prompts per file.
- **Captions.** Pandoc's captions land on the `Caption` style but carry no Word caption fields, so cross-references and a table of figures need the fields inserted in Word.
- **Display equations.** Math converts correctly but does not pick up the `MTDisplayEquation` style. Apply it where you want NIST equation spacing.

The tables are mostly already done: stage four sets the Table Style Options, moves cell text onto `TableText`, and rebuilds the row indentation. `Set-TableLook.ps1`, `Set-TableParagraphStyle.ps1`, and `Restore-TableIndents.ps1` are still there to re-run by hand on a file that needs different settings.

Two of the stage-four steps decide by shape what a thing means, so on a **new** document set check them once. Both run by default; these are the escape hatches:

- **Table captions.** `Set-NISTTableCaptions.ps1 -ListOnly` lists what it would take. Watch for a section heading that happens to sit above a table. Turn it off with `-SkipTableCaptions`.
- **Table footnotes.** `Merge-TableFootnotes.ps1 -ListOnly` does the same. Watch for a table whose last row is a banner or a merged total. Turn it off with `-SkipTableFootnotes`.

## Options

### Convert-ToNISTWord.ps1

| Flag | What it does |
|------|-------------|
| `-File` / `-Folder` | One source file, or every supported file in a folder. |
| `-ReferenceDoc` | Reference document to use. Defaults to `NIST-reference.docx` beside the script. |
| `-OutputFolder` | Where the `.docx` files go. Defaults to the source folder. |
| `-TableStyle` | Table style for converted tables. Defaults to `NIST-Data`. Pass `''` to leave pandoc's table styling alone. |
| `-HeaderRow` / `-TotalRow` / `-FirstColumn` / `-LastColumn` / `-BandedRows` / `-BandedColumns` | Table Style Options written in stage four. Defaults are the converted-table set: everything off except Banded Columns. |
| `-RepeatHeaderRow` | Repeat the first row across page breaks. Default `$true`, independent of `-HeaderRow`. |
| `-TableTextStyle` | Paragraph style put on every paragraph inside a table. Defaults to `TableText`. Pass `''` to skip. |
| `-SkipInlineRepair` | Skip stage zero. For HTML that already carries real `<strong>`/`<em>`/`<sup>`. |
| `-SkipTableIndents` | Skip the indent rebuild in stage four. |
| `-FootnoteStyle` | Paragraph style for the merged table notes. Defaults to `TableandFigureFootnotes`. |
| `-AutoFitTables` | Leave tables on Word's autofit instead of pinning them to a width. |
| `-TableWidthInches` | Pin every table to this width, rescaling columns in proportion. The NIST text column is 6.5. |
| `-UnboldColumns` | Columns to strip direct bold from in the body rows. Off by default. |
| `-SkipTableBanners` / `-BannerFill` | Turn off, or recolor, the group-label banner shading. Fill defaults to `CCECE3`. |
| `-SkipTableFootnotes` | Do not merge each table's footnote rows. |
| `-SkipTableCaptions` | Do not turn the header above each table into a numbered caption. |
| `-KeepSourceNumbers` | Keep a number the source already put at the front of a table title. |
| `-CaptionFromStyle` | StyleIds the table headers are on. Defaults to `Heading1`; takes a list, e.g. `Heading1,Heading2,Caption`. |
| `-CaptionStyle` | Caption style to apply. Defaults to `CaptionTable`. |
| `-CaptionChapter` | Chapter number before the dash. Defaults to `1`. One value per run. |
| `-CaptionUseStyleRef` | Write the chapter number as a `STYLEREF` field instead of literal text. |
| `-RemapLists` | Move lists onto the NIST List Bullet and List Number ladders. Off by default, see the caveat below. |
| `-Bibliography` | Bibliography file (`.bib`, `.json`, `.yaml`). Turns on `--citeproc`. |
| `-Csl` | CSL style file for citeproc. |
| `-NoMedia` | Skip `--extract-media`. Images are embedded either way. |
| `-PandocArgs` | Extra arguments passed to pandoc verbatim. |
| `-PandocPath` | Path to `pandoc.exe` if it is not on PATH. |
| `-SkipRemap` | Skip stage two, leaving pandoc's own styles in place. |
| `-SkipPipeline` | Stop after stage two, so you can inspect the files before the NIST build and the table finish run. |
| `-LegacyStyleFix` | Run the style step in full legacy-fix mode instead of `-Clean`. Rarely needed. |
| `-Force` | Skip prompts and overwrite existing output. |

### Build-NISTReferenceDoc.ps1

| Flag | What it does |
|------|-------------|
| `-Template` | The NIST template to build from. Required. |
| `-Output` | Where to write the reference document. Defaults to `NIST-reference.docx` beside the script. |
| `-ApplyStyleArchitecture` | Run the style architecture on the copy first, before the pandoc styles are added. |
| `-Force` | Skip the prompt and overwrite an existing output file. |

## Worked examples

A single HTML chapter, inspected before the build runs:

```powershell
.\Convert-ToNISTWord.ps1 -File "C:\src\chapter-03.html" -OutputFolder "C:\out" -SkipPipeline
# open C:\out\chapter-03.docx, check it, then:
.\Build-NISTDocuments.ps1 -Folder "C:\out" -Clean
```

A LaTeX paper with citations:

```powershell
.\Convert-ToNISTWord.ps1 -File "C:\src\paper.tex" -OutputFolder "C:\out" `
    -Bibliography "C:\src\refs.bib" -Csl "C:\src\nist.csl"
```

A batch, unattended, with lists moved onto the NIST ladder:

```powershell
.\Convert-ToNISTWord.ps1 -Folder "C:\src\report" -OutputFolder "C:\out\report" -RemapLists -Force
```

Statistical tables with notes at the foot of them. Check what the merge would take before turning it on, then convert:

```powershell
.\Convert-ToNISTWord.ps1 -File "C:\src\descriptivetabs.html" -OutputFolder "C:\out" -Force
```

That is the whole job: Table Style Options, widths pinned, cell text on `TableText`, row indents rebuilt, banner rows shaded, footnote rows merged, and the headers above the tables turned into numbered captions with a `STYLEREF` chapter. To see what the two shape-based steps would take before committing to them:

```powershell
.\Merge-TableFootnotes.ps1  -File "C:\out\descriptivetabs.docx" -ListOnly
.\Set-NISTTableCaptions.ps1 -File "C:\out\descriptivetabs.docx" -ListOnly
```

Open the result and press Ctrl+A then F9 so the `SEQ` fields settle, and the numbers and any table of tables read correctly.

## What maps to what

Stage one creates these in the reference document. Stage two folds most of them into NIST styles.

| Pandoc style | Based on | Fate in the finished document |
|--------------|----------|-------------------------------|
| First Paragraph | Body Text | folded into Body Text |
| Compact | Body Text | folded into Body Text |
| Block Text | Quote | folded into Quote |
| Image Caption | Caption | folded into Caption |
| Table Caption | Caption | folded into Caption |
| Figure | _graphic | folded into _graphic |
| Captioned Figure | _graphic | folded into _graphic |
| Author | Body Text | folded into Body Text |
| Date | Body Text | folded into Body Text |
| Abstract | Body Text | folded into Body Text |
| Abstract Title | Front_matter_head 1 | folded into Front_matter_head 1 |
| Source Code | Body Text, Consolas | **kept** |
| Verbatim Char | Default Paragraph Font, Consolas | **kept** |
| Bibliography | Body Text | kept |
| Definition Term | Body Text | kept |
| Definition | Body Text | kept |
| TOC Heading | Front_matter_head 1 | kept |

`Source Code` and `Verbatim Char` are the only two meant to survive into finished documents. Nothing in the NIST set covers code. Every fold falls back sensibly if your template lacks the target: `Quote` falls back to Body Text, `_graphic` falls back to Body Text, `Front_matter_head 1` falls back to Heading 1.

## Known limits

**Tables.** Pandoc's table style is named `Table`, which the table-style allowlist in `Set-NISTStyleArchitecture.ps1` would delete. The reference document deliberately does not define it. The converter assigns `NIST-Data` instead. If you skip the pipeline and your reference document has no `NIST-Data`, the converter warns and leaves tables unstyled rather than writing a reference to a style that does not exist.

**Lists.** Without `-RemapLists`, lists keep pandoc's direct numbering, which looks close but is not the NIST ladder. With `-RemapLists`, bullets convert cleanly, but ordered lists lose pandoc's explicit start values and restarts, so two separate numbered lists can run on as one sequence. Check numbered lists after using it.

**LaTeX figures in PDF or EPS.** They cannot be embedded in Word. The converter lists the offending files before it runs and pandoc drops them. Convert to PNG or SVG first.

**Heavy LaTeX.** Pandoc's reader ignores custom macros and environments it does not know. The `raw_tex` extension is on, so what it could not interpret shows up as raw text rather than vanishing, which makes the gaps visible. For journal-grade sources with heavy macro use, run `latexml` and `latexmlpost` to HTML5 first and convert that HTML, so both source types go through the same path.

**R `flextable` LaTeX cannot be read at all.** Quarto and R Markdown emit tables as hand-drawn `longtable*` blocks full of `\global\setlength`, `\ascline`, and per-cell `\multicolumn{1}{>{}l}{\textcolor[HTML]{...}{\fontsize{7}{7}...}}`. Pandoc's reader stops on the first `\global\setlength` with `unexpected \setlength` and exit code 64. There is no flag that fixes this. Convert the Quarto `.docx` output instead, where the same tables are already native Word tables. If only the `.tex` exists, re-render the Quarto source to `docx`.

**File names.** Two sources with the same stem (`report.html` and `report.tex`) both want `report.docx`. The converter detects this before doing any work and stops, rather than letting the second overwrite the first. Rename one or convert them one at a time.

## Preparing Quarto sources

Quarto output carries three things that fight the NIST template. Strip them from the source before converting, not from the finished Word file.

**Baked-in section numbers.** NIST `Heading1` through `Heading9` auto-number from `numId 11`, so Quarto's own numbers render twice: `1. 1. Recovery model`. In HTML they sit in `<span class="header-section-number">`; in a Quarto `.docx` they are literal text at the front of each heading paragraph. Remove them and let the styles number.

**A flattened table of contents.** Quarto writes a static TOC whose page numbers are all `1`, on `TOC1` through `TOC3` and `TOCHeading`. Delete those paragraphs; Word builds a live one from the headings.

**The title block.** `<header id="title-block-header">` in HTML, and a Title plus author and date paragraphs in `.docx`, become stray body text because the NIST front matter supplies all of it. Drop the block in HTML with the sidebar `<nav id="TOC">` alongside it.

**All inline formatting done in CSS.** This is the one that costs you a review cycle, because nothing warns you. `flextable` emits **no semantic tags at all**: no `<strong>`, no `<em>`, no `<sup>`. Every run is a `<span class="cl-xxxxxxxx">` whose generated class carries `font-weight: bold` or `vertical-align: super`. Pandoc reads tags, not CSS, so all of it is silently flattened to plain body text.

Parse the `<style>` blocks, map each generated class to the tags its rule implies, and rewrite the spans in a single pass:

| CSS in the class rule | Tag to emit |
|---|---|
| `font-weight: bold` (or 600-900) | `<strong>` |
| `font-style: italic` | `<em>` |
| `vertical-align: super` | `<sup>` |
| `vertical-align: sub` | `<sub>` |

Do it in one pass, not one pass per tag. A class is often both bold and superscript, and consuming the span in a first pass loses whatever the second pass would have added. In the Chapter 5 source, 72 of the 469 formatted spans were both.

Grep the source for `font-weight` and `vertical-align` to see how much applies.

**SVG figures in a Quarto `.docx`.** Word stores each figure as a PNG `<a:blip>` plus an `asvg:svgBlip` extension. Pandoc's reader prefers the SVG, and its writer then emits a blip with no `r:embed` at all, so Word renders nothing. Delete the `<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">` elements from the source `document.xml` first; that leaves the PNG as the only representation and it round-trips cleanly. Quarto's PNGs are usually 2400 px wide, which is 369 dpi at full text width, well past the print threshold.

## Troubleshooting

**"Reference doc not found"** - you have not run `Build-NISTReferenceDoc.ps1` yet, or you moved `NIST-reference.docx`. Point at it with `-ReferenceDoc`.

**"pandoc not found"** - pandoc is not on PATH. Pass `-PandocPath "C:\Program Files\Pandoc\pandoc.exe"`. Remember it lives on Windows, not in WSL.

**Text comes out in Normal instead of a NIST style** - a style reference is pointing at a style that does not exist. This is what stage two prevents, so check whether it was skipped with `-SkipRemap`.

**The pipeline stops at a step** - `Build-NISTDocuments.ps1` fails fast and names the step and file. Each underlying script backs up before writing, so the file is recoverable from its `.bak`. Fix the cause and re-run from that step.

**Table captions all read "Table 5-0"** - the `SEQ` fields have not been evaluated yet. Ctrl+A then F9 in Word, which also refreshes any table of tables. The cached values written into the file are correct, so this only shows up before the first update.

**A section heading turned into a table caption** - it sat directly above a table, which is the whole of the rule. Restore from the `.captions.bak` beside the file, then narrow `-FromStyle` to the level your captions actually use, or move that heading.

**A grey or colored block under a table that nothing removes** - those are the footnote rows. They are full-width merges, so the table style's banding paints them, and `Remove-Shading.ps1` will not touch them because the color is not direct formatting. Run `Merge-TableFootnotes.ps1`, which merges them into one row carrying `w:fill="auto"` and a zeroed `w:cnfStyle` so no Table Style Option applies.

**Bold and superscripts missing from converted HTML tables** - the HTML carries them in CSS classes rather than tags, and pandoc reads tags. Stage zero handles this; check it was not turned off with `-SkipInlineRepair`, and run `Repair-HtmlInlineFormatting.ps1 -ListOnly` on the source to confirm the classes are being read.

**Everything in the output folder got reformatted** - stage three processes every Word file in `-OutputFolder`. Use a dedicated folder.

## See also

- `readme.md` in this folder for the rest of the Word toolkit and the manual NIST build sequence
- `Build-NISTDocuments.ps1` for what stage three actually runs
- `publishing/md-docx/` for generic pandoc wrappers with no template awareness, useful for throwaway or interchange work
