# Word Template Toolkit

PowerShell scripts for working with Word templates at the OOXML layer. Reads files directly as zip archives of XML using `System.IO.Compression` and `System.Xml`, so Word does not need to be installed. Works on Windows, WSL, Linux, and macOS with PowerShell 5.1+.

Most scripts are client-agnostic and work on any Word document. The `Set-NIST*` / `Audit-NIST*` scripts are NIST-specific: they are thin wrappers that supply the NIST palette, style list, and margins and delegate to the generic engines in this same folder. The `ncstar/` and `nwirp/` subfolders hold the report-assembly tools for those two NIST document sets.

## Running the scripts

`cd` into this folder and run each script with `.\`. Do NOT copy the scripts into the document folder - the `Set-NIST*` wrappers find their generic engine and the `OoXml/` module by their own location, so they must run from here. Point `-Folder` at wherever the `.docx` files live.

```powershell
# cd into the toolkit folder (WSL path from Windows; adjust the distro name if yours differs)
cd \\wsl$\Ubuntu-24.04\home\artist\ovoco\tools\publishing\word

$docs = "$env:USERPROFILE\Documents\260726"
.\Set-PageLayout.ps1 -Folder $docs
```

(From another directory without `cd`, call by full path with the `&` operator instead: `& "\\wsl$\...\word\Set-PageLayout.ps1" -Folder $docs`.)

- Each script lists its targets and asks `Proceed? (y/N)`. Type `y`, or add `-Force` to skip.
- Swap `-Folder` for `-File "one.docx"` to test on a single file first.
- Scripts that modify files write a `.bak` backup before overwriting.

## NIST template build

Apply the NIST template to a folder of chapter files. Order matters: step 1 installs the palette that step 2's styles reference.

```powershell
cd \\wsl$\Ubuntu-24.04\home\artist\ovoco\tools\publishing\word
$docs = "$env:USERPROFILE\Documents\260726"

.\Set-NISTTheme.ps1               -Folder $docs
.\Set-NISTStyleArchitecture.ps1   -Folder $docs
.\Set-PageLayout.ps1              -Folder $docs
.\Set-TableLook.ps1               -Folder $docs
.\Set-DocTitle.ps1                -Folder $docs   # no -Title, so it prompts per file
.\Set-NISTHeaderFooterContent.ps1 -Folder $docs

# Header/footer: chapters get a first-page footer; front matter stays blank.
# Run the front-matter file with -FrontMatter, the chapters without it.
.\Set-NISTHeaderFooterContent.ps1 -File "$docs\NCSTAR_4-3_Chapter-00_Front-Matter.docx" -FrontMatter
Get-ChildItem "$docs\*.docx" -Exclude '*Front-Matter*' | ForEach-Object {
    .\Set-NISTHeaderFooterContent.ps1 -File $_.FullName -Force
}
```

**Entering the title.** Leave `-Title` off and `Set-DocTitle.ps1` walks each file one at a time, showing `[n/total] filename`, and prompts you for that file's title:

```
[1/7] NCSTAR_4-3_Chapter-00_Front-Matter.docx
Title (blank to skip):
```

Type the title and press Enter; it sets it and moves to the next file. Leaving a title blank skips that file. This is the usual mode when each chapter carries its own title. Pass `-Title "..."` to apply one title to every file instead (with a single y/N confirmation), or `-File "$docs\Chapter-00.docx"` to target just one.

Notes:

- **NIST margins are 1 inch on all four sides**, with header and footer at 0.3". That is the documented standard and the `Set-PageLayout.ps1` default, so the build line needs no margin flags. The script still takes `-InnerInches`/`-OuterInches` if a one-off document needs mirror margins.
- **Style colors are fixed hex**, not theme-linked, so they survive Word's Import Styles into a document with a different theme. Changing Design > Colors will not recolor them.
- `Reapply-Styles.ps1` (strip direct character formatting) and `Remove-Shading.ps1` (strip direct shading) are separate post-authoring cleanup steps, not part of this build. Both remove formatting that is often intentional, so run them per document when you know a file needs it.

## HTML, LaTeX, and Markdown into the NIST template

Two scripts turn non-Word source into NIST-formatted Word documents. They need `pandoc` on PATH; everything else in this folder works without it. Full instructions, options, and troubleshooting are in [`converting-html-latex.md`](converting-html-latex.md); the summary follows.

Build the reference document once per template revision:

```powershell
cd \\wsl$\Ubuntu-24.04\home\artist\ovoco\tools\publishing\word
.\Build-NISTReferenceDoc.ps1 -Template "$env:USERPROFILE\Documents\260806\HM-NCSTAR-Template.docx" -ApplyStyleArchitecture
```

That writes `NIST-reference.docx` beside the scripts, which is where the converter looks by default. Then convert:

```powershell
.\Convert-ToNISTWord.ps1 -Folder "C:\src\report" -OutputFolder "C:\out\report"
```

Five stages run in order: an inline-formatting repair on HTML sources, pandoc against the reference doc, a remap pass that folds pandoc's styles into NIST styles, `Build-NISTDocuments.ps1 -Clean` on the output folder, then the table finish (Table Style Options, cell text, unbolding, indents, footnote merge, captions). Header/footer content and document titles are still separate steps afterwards.

**Why a reference doc.** Pandoc resolves styles from the reference document **by name**, not by styleId. It looks for names like `Source Code`, `Block Text`, `Image Caption`, and `First Paragraph`. Any name it cannot find, it invents with its own look. `Build-NISTReferenceDoc.ps1` adds a definition for each of those names based on the closest NIST style, and leaves alone any name the template already defines. The theme, section properties, numbering, and header/footer parts come along automatically because the reference doc is a copy of the template.

**Why the remap stage.** It runs before the build pipeline on purpose. `Set-NISTStyleArchitecture.ps1` deletes styles it does not recognize but never edits `document.xml`, so a paragraph still pointing at a deleted pandoc style would quietly fall back to Normal. The remap rewrites those references first, so by the time the pipeline runs, only NIST styles are referenced.

Notes:

- **Tables.** Pandoc's table style is named `Table`, which the table-style allowlist would delete, so the reference doc deliberately does not define it. The converter assigns `NIST-Data` to every converted table instead (`-TableStyle` to change it, `-TableStyle ''` to leave pandoc's alone).
- **Inline formatting in HTML.** Pandoc reads tags, not CSS. HTML that carries its bold and superscripts only in a stylesheet arrives flattened, so the converter runs `Repair-HtmlInlineFormatting.ps1` first on a throwaway copy. See below.
- **Code.** `Source Code` and `Verbatim Char` are the only pandoc styles kept in finished documents. Nothing in the NIST set covers code.
- **Lists** stay on pandoc's direct numbering unless you pass `-RemapLists`, which moves them onto the NIST List Bullet / List Number ladders. Bullets convert cleanly; ordered lists lose pandoc's explicit start values, so separate numbered lists can run on as one sequence.
- **LaTeX figures** in PDF or EPS cannot be embedded in Word. The converter lists them and pandoc drops them; convert to PNG or SVG first.
- **Math** converts to native Word equations (OMML). The `MTDisplayEquation` style is not applied automatically.
- For heavy LaTeX (custom macros, theorem environments), consider `latexml` to HTML5 first and convert that HTML, so both source types share one path.

### Repairing CSS-only HTML formatting

Some generators emit no semantic tags at all. R flextable is the worst case: there is no `<strong>`, no `<em>`, no `<sup>` anywhere in the file. Every run of text is a `<span class="cl-xxxx">`, and the bold, the italic, and the superscript live only in the generated stylesheet. Pandoc reads tags, so all of it is dropped, and nothing warns you - the output simply comes out flat.

`Repair-HtmlInlineFormatting.ps1` reads the `<style>` blocks, works out what each class means, and rewrites the spans as real tags.

```powershell
# Look first: what does each class actually carry?
.\Repair-HtmlInlineFormatting.ps1 -File "C:\src\tables.html" -ListOnly

# Rewrite in place (backs up to <name>.inline.bak)
.\Repair-HtmlInlineFormatting.ps1 -File "C:\src\tables.html"

# A folder of HTML
.\Repair-HtmlInlineFormatting.ps1 -Folder "C:\src\html"
```

`Convert-ToNISTWord.ps1` runs this for you as stage 0 on any HTML source, on a working copy in the source folder so relative image paths still resolve; the copy is deleted afterwards and your file is never touched. Run it by hand only when you want to inspect the result, or when you are feeding the HTML to something other than the converter. `-SkipInlineRepair` turns the stage off for HTML that already has real tags.

The rewrite happens in **one pass**, which matters: a class is often both bold and superscript, and a pass per tag would consume the span on the first one and lose the second. It is idempotent, so a second run finds nothing to do.

## Scripts

### NIST template (this folder)

| Script | What it does |
|--------|-------------|
| `Set-NISTTheme.ps1` | Thin wrapper: passes the NIST NCSTAR palette to `Set-ThemeColors.ps1`. |
| `Set-NISTStyleArchitecture.ps1` | Imports `OoXml/OoXml.psd1` and rebuilds the NIST style set in `styles.xml` and `numbering.xml`: rebases styles onto the `xx*` engine, fixes basedOn, deletes junk styles, sets fixed-hex colors, and configures the headings, lists, and table styles. See "What Set-NISTStyleArchitecture configures" below for the specifics. |
| `Set-NISTHeaderFooterContent.ps1` | Installs even/odd/first header and footer parts with STYLEREF chapter number/title and TITLE+PAGE fields. Chapter first pages get a TITLE+PAGE footer; pass `-FrontMatter` to leave the front-matter first-page footer blank. |
| `Install-NISTThemeColors.ps1` | Thin wrapper: installs the NIST palette into the user's Word Theme Colors folder. Per-user; does not touch documents. |
| `Audit-NISTStyleSet.ps1` | Thin wrapper: passes the required/optional NIST style ids to `Audit-StyleSet.ps1`. Pre-assembly gate for `ncstar/Build-NISTReport.ps1`. |
| `Build-NISTReferenceDoc.ps1` | Copies the NIST template to `NIST-reference.docx` and adds a definition for every style name pandoc looks up, based on the closest NIST style. Names the template already defines are left untouched. |
| `Convert-ToNISTWord.ps1` | Converts .html/.htm/.xhtml/.tex/.md into NIST-formatted .docx in five stages: repair of CSS-only inline formatting (HTML sources), pandoc against the reference doc, remap of pandoc styles to NIST styles, `Build-NISTDocuments.ps1`, then the table finish (Table Style Options for converted tables, cell text onto `TableText`, widths pinned, row indents rebuilt, banner rows shaded, table note rows merged, and the table headers turned into numbered captions). Also takes a .docx source (`-File` only), which retargets an existing pandoc or Quarto Word file at the NIST template. Needs pandoc on PATH. |

#### What Set-NISTStyleArchitecture configures

All colors are written as fixed hex (no theme link) so they survive Word's Import Styles.

- **Rebasing** - Heading1-9, BodyText, Caption, and the List Bullet/Number family are rebased onto the `xx*` engine styles to break Normal inheritance. Junk styles are deleted; a table-style allowlist keeps only `TableNormal`, `TableGrid`, `HM-Cover`, `HM-Data`, `NIST-Data`, and `SidebarTable`.
- **Headings** - Heading 1-4 and 6-9 are Resolution Blue `#00257A`; **Heading 5** is neutral gray `#404040`. **Heading 4 is numbered** as part of the heading list (renders `1.2.3.4`); use the separate **Heading 4 Sub-Head** style for an unnumbered level-4 heading.
- **Appendix / front-matter heads** - AppendixHead 1-4 and Front_matter_head 1-4 inherit their matching Heading's font/size/spacing but are **unnumbered** (`numId=0`). Front-matter heads are forced **black**, not the heading blue.
- **Lists** - List, List Bullet, List Continue, and List Number share one ladder that lines up: the marker steps 0.2" per level from 0.20", the text hangs 0.15" past it, and the numbering tab stops are synced to the text position. **Table Bullet** (for bullets in table cells) sits tighter - bullet 0.1" in, 0.13" gap. The bullet-glyph ladder is set per level.
- **Paragraph tweaks** - **TableText** = 0.9 line multiple, 2pt space-after (auto-spacing flags stripped). **Table of Figures** spacing matches TOC 1. **_graphic** has keep-with-next / keep-lines-together removed and no basedOn. **Sidebar** / **Sidebar Bullet** callout text styles have their floating frame removed.
- **Table styles** - **HM-Data** and **NIST-Data**: all rules/gridlines unified to `#A6A6A6`, default cell margins top/bottom 0.01" / left 0.04" / right 0.02"; HM-Data's first column is shaded `#F2F2F2`. **Sidebar Table** (new): a callout box with a whole-table amber `#FFF2CC` fill, 0.04" cell margins, and a fixed 3" width; `Set-TableLook.ps1` floats it at the right margin (the float can't live in a table style).

### Inspection (read-only)

| Script | What it does |
|--------|-------------|
| `Get-WordStyles.ps1` | Lists every style in a Word file as Markdown. One section per style with Identity, Font, Paragraph, and Table attributes. Optional HTML view. |
| `Get-WordTableStyles.ps1` | Table styles only, with full conditional formatting blocks (header row, banded rows, first/last column, corner cells). |
| `Get-WordParts.ps1` | Lists Building Blocks (Quick Parts, AutoText, cover pages) from the glossary with metadata and content preview. |
| `Compare-WordTemplates.ps1` | Compares every Word file in a folder for drift. Produces CSV + HTML showing where they disagree on styles, headers, footers, building blocks, content controls, fields, bookmarks, or custom XML parts. |
| `Audit-WordTheme.ps1` | Reports the theme palette and lists every place a style or run sets an explicit hex color instead of a theme reference. |

### Modification

| Script | What it does |
|--------|-------------|
| `Expand-WordPackage.ps1` | Unpacks a Word file into a folder of XML files for hand editing. Creates a backup first. |
| `Compress-WordPackage.ps1` | Repacks an unpacked folder back into a valid Word file. Validates the result before overwriting. |
| `Reapply-Styles.ps1` | Strips run-level direct character formatting so paragraph styles render cleanly. Preserves italic, superscript/subscript, and OLE/MathType equations. |
| `Set-TableParagraphStyle.ps1` | Puts every paragraph inside a table onto one paragraph style (`TableText` by default) and changes nothing else. Bold, italic, super/subscript, indentation, alignment, and numbering all survive. Handles nested tables. `-KeepStyles` exempts captions or footnotes living inside a table. `-UnboldColumns` additionally strips direct bold from the given columns in the body rows, for converted tables whose group labels arrive bold. |
| `Repair-HtmlInlineFormatting.ps1` | Rewrites CSS-only inline formatting in HTML into real tags before a converter sees it. R flextable emits no `<strong>`, `<em>`, or `<sup>` at all - every run is a `<span class="cl-xxxx">` whose formatting lives in the generated stylesheet, and pandoc reads tags, not CSS. Reads the `<style>` blocks, maps each class to bold/italic/underline/super/sub, and rewrites every affected span in one pass so a class that is both bold and superscript keeps both. `-ListOnly` reports the classes first. Idempotent. |
| `Set-TableBannerRows.ps1` | Shades and bolds the group-label banner rows inside a table - the full-width merged rows that title a group (`Ownership (n = 297)`) rather than carry data. Fill defaults to `CCECE3`. Tells banners from footnote rows by paragraph style, so run `Merge-TableFootnotes.ps1` first and the notes are already claimed; `-SkipTrailing` is the fallback when it has not been. `-ListOnly` first. Idempotent. |
| `Merge-TableFootnotes.ps1` | Collapses the note rows at the foot of each table into a single full-width row with no fill, on `TableandFigureFootnotes`. Zeroes the row's `w:cnfStyle` and sets `w:fill="auto"` so neither banding nor first/last-row emphasis can paint it, strips the bold a banner misdetection leaves behind, and drops the indent; superscript markers survive. Finds the rows by shape (trailing full-width merges), or exactly from the source HTML's `<tfoot>` blocks with `-SourceHtml`. Run `-ListOnly` first. Idempotent. |
| `Restore-TableIndents.ps1` | Rebuilds the cell indentation pandoc drops. Reads the level from leading tabs in a `-Source` original, or from leading spaces in the target alone, and writes a real `w:ind`. Verifies the table shapes match before writing. |
| `Set-NISTTableCaptions.ps1` | Converts the paragraphs sitting directly above tables (`Heading1` by default) into `Caption_Table` with the NIST number: `Table 1-1.  Title`, where the table number is a `SEQ` field. `-FromStyle` takes a list, so a converted document whose table headers sit at several heading levels (or on `Caption`, where pandoc puts an HTML `<caption>`) is done in one pass with the numbering still sequential. `-Chapter` sets the chapter digit, `-UseStyleRef` makes it a `STYLEREF` field instead. Run `-ListOnly` first. |
| `Remove-Shading.ps1` | Removes locally applied (direct) shading from runs and paragraphs, so text and bullets show no background. Keeps style-defined shading (styles.xml is never touched) and table-cell fills (unless `-IncludeTables`). |
| `Repair-Typography.ps1` | Converts straight quotes to curly (typographic) quotes, deciding open/close by context across run and paragraph boundaries. Only `<w:t>` text is changed - field codes and markup stay straight. Idempotent. `-Ellipses` also converts `...` to a single ellipsis. |
| `Set-DocTitle.ps1` | Sets `<dc:title>` in docProps/core.xml. Works for any title string in any Word document. |
| `Merge-WordDocuments.ps1` | Combines several Word files into one with a section break between each. Uses the first file as the host (theme, styles, headers), renumbers rIds/bookmarks/footnotes, copies and renames media, accepts revisions, and strips comments. Originals untouched. |
| `Set-PageLayout.ps1` | Sets page margins with mirror margins. Defaults to the 1" standard on all four sides, header and footer 0.3". Configurable via -InnerInches, -OuterInches, -TopInches, -BottomInches, -HeaderInches, -FooterInches. |
| `Set-TableLook.ps1` | Normalizes `<w:tblLook>` and adds `<w:tblHeader/>` to first rows so they repeat across page breaks. `-FixedWidth` pins each table to a real width with `<w:tblLayout w:type="fixed"/>` instead of leaving it on autofit, so two tables with the same grid render the same width; `-TableWidthInches` rescales the columns to a target (the NIST text column is 6.5"); `-ColumnWidthsInches 2.7,0.6,0.6 -MatchHeader Characteristic,N,%` sets exact columns on one family of tables and leaves the rest alone. Widths are written at all three levels Word reads - `tblGrid`, `tblW`, and every cell's `tcW`, with `gridSpan` honoured - because it re-derives them from the contents if any of the three disagrees. All six Table Style Options are parameters (`-HeaderRow`, `-TotalRow`, `-FirstColumn`, `-LastColumn`, `-BandedRows`, `-BandedColumns`) and the legacy `w:val` bitmask is computed from them; the defaults are the formal report look (Header Row + Banded Rows), and `-RepeatHeaderRow` is independent of the Header Row checkbox. Tables using the Sidebar Table style are instead floated as right-margin callout boxes (text wraps around, 3" wide, no repeat header) since the float can't live in a table style. |
| `Set-TableLook.ps1` (`-ReportWidths`) | Measures every cell with the real installed font and prints the narrowest each column can be, grouped by table shape, then suggests a `-ColumnWidthsInches` line to paste back. Writes nothing. Needs Windows (System.Drawing). Full-width rows on the footnote and caption styles are excluded, since those are prose and are meant to wrap. |
| `Set-ThemeColors.ps1` | Replaces the `<a:clrScheme>` in theme1.xml with a caller-supplied color scheme XML block. Validates all 12 theme slots are present. |
| `Install-ThemeColors.ps1` | Installs a theme color scheme file into the user's Word Theme Colors folder so Design > Colors displays it by name. |
| `Audit-StyleSet.ps1` | Generic style audit engine. Checks that caller-supplied required style ids exist in styles.xml. Returns non-zero on drift. JSON or table output. |

### Putting table text on Table Text

Converters leave cell paragraphs on whatever style they picked, usually Body Text, and pandoc puts anything it read as a block quote on Quote, which carries a border and a fill that look wrong inside a cell. `Set-TableParagraphStyle.ps1` moves them all onto `TableText`.

```powershell
.\Set-TableParagraphStyle.ps1 -Folder $docs

# Leave captions and table footnotes on their own styles
.\Set-TableParagraphStyle.ps1 -File "$docs\Chapter-05.docx" -KeepStyles CaptionTable,TableandFigureFootnotes
```

`-UnboldColumns` is the one thing it changes beyond the paragraph style, and only when asked. It strips direct bold from the given columns in the body rows, leaving the header row alone. Note that in a converted statistical table the bold you see in column 1 is usually a group-label banner row, which is *meant* to be bold - `Set-TableBannerRows.ps1` shades those to match. Reach for `-UnboldColumns` only for bold that is genuinely unwanted.

Otherwise it rewrites `<w:pStyle>` and nothing else, which is the opposite end of the scale from `Reapply-Styles.ps1`:

| | `Reapply-Styles.ps1` | `Set-TableParagraphStyle.ps1` |
|---|---|---|
| Run formatting (`w:rPr`) | wiped, except italic and super/subscript | untouched, so bold, super/subscript, and character styles all survive - except bold in `-UnboldColumns`, if given |
| Paragraph style | unchanged | rewritten |
| Indentation, alignment, numbering | unchanged | unchanged |
| Scope | whole document, headers, footers, notes | paragraphs inside `<w:tbl>` in the body only |

It refuses to run if the style is missing from `styles.xml`, since a reference to an undefined style renders as Normal with no warning. Backs up to `<name>.tablestyle.bak`, and is idempotent.

### Merging table footnotes

Statistical tables carry their notes at the foot of the table - the a/b/c qualifiers, the significance codes, the "Reference levels:" legend. A converter emits each note as its own row, and because those rows are full-width merges they collect whatever the table style paints on them: the banding stripe, or a header fill if something upstream mistook a merged row for a banner. That is the grey (or green) block under a table that nothing seems to remove.

`Merge-TableFootnotes.ps1` collapses that run of rows into one.

```powershell
# Look first. A table whose real last row is a banner or a merged total
# matches the same rule as a footnote, and only the text tells them apart.
.\Merge-TableFootnotes.ps1 -File "$docs\Chapter-05.docx" -ListOnly

# Write (backs up to <name>.footnotes.bak)
.\Merge-TableFootnotes.ps1 -File "$docs\Chapter-05.docx" -Force

# Whole folder
.\Merge-TableFootnotes.ps1 -Folder $docs -Force
```

The row it writes is one cell spanning every column, and it is armored against the table style in two ways: `w:fill="auto"` on the cell, and a zeroed `w:cnfStyle` on the row so none of the Table Style Options - banding, first row, last row - apply to it. It stays uncolored wherever it lands in the stripe pattern, and it stays uncolored if you change the banding later.

Inside the cell, one paragraph per note, stacked, on `TableandFigureFootnotes`. Bold is stripped, since a banner misdetection is what put it there, and the indent is dropped because notes sit flush left. Superscript and subscript markers survive - nothing here touches `w:rPr` beyond the bold.

**Which rows it takes.** By default, structural: it walks up from the last row and takes the run of full-width merged rows it finds, stopping at the first row that is not one. Two ways to narrow that:

```powershell
# Exact, from the <tfoot> blocks of the HTML the tables came from
.\Merge-TableFootnotes.ps1 -File "$docs\Chapter-05.docx" `
    -SourceHtml "C:\src\descriptivetabs.html" -Force

# Or require the text to look like a note
.\Merge-TableFootnotes.ps1 -File "$docs\Chapter-05.docx" -Marker '^[a-z][A-Z(]' -Force
```

`-KeepBold` and `-KeepIndent` leave those alone, and `-Style` puts the notes on a different paragraph style.

It refuses to flatten a row that contains a nested table, and skips a table where every row matched, which is a one-column table rather than a table with notes at the foot of it. Idempotent: the row it writes is recognized and rebuilt identically on a re-run.

`Convert-ToNISTWord.ps1` runs it in stage four by default; `-SkipTableFootnotes` turns it off. The structural rule is a guess about meaning, so on a new document set run it here with `-ListOnly` once and read the listing before trusting it on that shape of table.

### Measuring table column widths

Column widths are guesswork unless the text is measured. Two characters-per-inch estimates for the same content differed by 44%, which is exactly the difference between fitting and wrapping. `-ReportWidths` measures instead:

```powershell
# Every table shape in the file
.\Set-TableLook.ps1 -File "$docs\Chapter-05.docx" -ReportWidths

# One family
.\Set-TableLook.ps1 -File "$docs\Chapter-05.docx" -ReportWidths -MatchHeader Characteristic,N,%
```

```
  3-column tables [Characteristic | N | %]  x33
    column 1: widest text  1.453"  -> needs  1.514"
    column 2: widest text  0.186"  -> needs  0.247"
    column 3: widest text  0.342"  -> needs  0.402"
    full-width rows: widest  3.685" -> the TABLE must reach 3.745"
      longest: Received disaster assistance funding from federal sources (n = 155)
    suggested: -ColumnWidthsInches 2.7,0.6,0.6 -MatchHeader Characteristic,N,%   (total 3.90")
```

The line that usually decides the answer is **full-width rows**. A group-label banner spans every column, so one long banner sets how wide the whole table has to be, while each column on its own would have fitted in half of it. Here column 1 needs 1.514" for its own labels but gets 2.7", entirely because of the banners.

Rows on the footnote and caption styles are excluded from that: they are prose, they are supposed to wrap, and counting them reports a table 42" wide. `-ExcludeSpanStyles` changes which styles are treated that way.

The suggestion gives the trailing columns one common width (floored at `-MinColumnInches`, default 0.6") and lets column 1 absorb what the banners need, plus `-HeadroomInches` (default 0.15"). Nothing is ever suggested wider than `-MaxTableInches` (default 6.5", the NIST text column); when the content genuinely needs more, the space is shared out in proportion and the report names the columns whose text will wrap.

### Table captions

Converted documents often carry a heading above each table where a caption belongs. `Set-NISTTableCaptions.ps1` converts them and adds the number.

```powershell
# Look before writing. A section heading that happens to sit above a table
# matches the same rule as a caption, and only the title text tells them apart.
.\Set-NISTTableCaptions.ps1 -File "$docs\tables.docx" -ListOnly

.\Set-NISTTableCaptions.ps1 -File "$docs\tables.docx" -Chapter 1
```

`-FromStyle` takes a list, which converted HTML usually needs. The headers above the tables are rarely all one level, and an HTML `<caption>` element inside a table comes through pandoc on `Caption` rather than on a heading style:

```powershell
.\Set-NISTTableCaptions.ps1 -File "$docs\tables.docx" -FromStyle Heading1,Heading2,Caption -Chapter 5 -ListOnly
```

The numbering runs in document order across the whole set, so mixing the levels does not disturb the sequence. A paragraph on one of those styles is only taken when the next thing in the body is a table, ignoring empty paragraphs in between, so headings that introduce prose are left alone.

The output matches the production chapters character for character: the word Table, a non-breaking space, the chapter number, an en dash, a `SEQ Table \* ARABIC \s 1` field, a period, and two spaces. Only the table number is a field, so it renumbers on Ctrl+A then F9.

The chapter number is literal text from `-Chapter`, because a converted extract has no Heading 1 for a `STYLEREF` field to read. Pass `-UseStyleRef` once the content is merged into a real chapter and the chapter number tracks the heading it lands under.

Captions go above the table, which is the NIST convention for tables and the opposite of figures.

`-StripLeadingNumber` drops a number the source already put at the front of the title, so the caption reads `Table 1-1.  Sample Characteristics` rather than `Table 1-1.  1 Sample Characteristics`. It only strips when that number equals the number the caption is being given, so `2020 Enrollment Data` landing on table 3 keeps its 2020. The number is usually split across several runs in a converted file, which is handled.

`Convert-ToNISTWord.ps1` runs this in stage four by default, with `-StripLeadingNumber` on; `-CaptionFromStyle`, `-CaptionChapter`, and `-CaptionUseStyleRef` pass straight through, `-KeepSourceNumbers` turns the stripping off, and `-SkipTableCaptions` turns the whole step off. `-CaptionChapter` is one value for the whole run, so convert a multi-chapter batch a file at a time.

### Removing stray shading

`Remove-Shading.ps1` targets shading applied as direct formatting - the "Pattern: Clear (White)" white box (and any other pale fill) that text and bullets pick up from pasting or template imports. Shading that comes from a style belongs and is kept.

This is **not** part of `Build-NISTDocuments.ps1`. Run it deliberately, on documents you know have unwanted direct shading.

```powershell
# One file (backs up to <name>.shading.bak)
.\Remove-Shading.ps1 -File "$docs\Chapter-02.docx"

# Whole folder
.\Remove-Shading.ps1 -Folder $docs
```

Two kinds of shading it leaves alone by default, and both are usually intentional:

- **Style-defined fills.** `styles.xml` is never opened. `NIST-Data` and `HM-Data` (`00257A` header rows), `SidebarTable` (`FFF2CC`), `Quote`/`QuoteChar` (`E6E9F2`), `BlockText` (`F2F2F2`), and `CaptionFigure2` (`D9D9D9`) all keep their fills.
- **Table-cell fills** in `w:tcPr`. `-IncludeTables` strips these, and it is blunt: it takes every cell fill in every table in scope, not just the unwanted ones. Data tables in the NCSTAR chapters carry hundreds of deliberate cell fills, so check a document before using it.

### Fixing quotes and typography

`Repair-Typography.ps1` converts straight quotes (`"` `'`) to curly typographic quotes, choosing opening vs closing by context - tracked across run and paragraph boundaries so a quote that starts a run still opens correctly. Apostrophes (`it's`, `teachers'`, `the '90s`) resolve to the right form. Only `<w:t>` text is touched, so field codes (like `STYLEREF "Heading 1"`) stay straight. Already-curly quotes are left alone, so it is safe to re-run.

```powershell
# Whole folder (backs up each file to <name>.typography.bak)
.\Repair-Typography.ps1 -Folder $docs

# One file
.\Repair-Typography.ps1 -File "$docs\Chapter-02.docx"

# Also convert ... to a single ellipsis
.\Repair-Typography.ps1 -Folder $docs -Ellipses
```

### Sidebar callout boxes

A sidebar is a shaded box that floats at the right margin with the body text wrapping around it. It is two pieces because the float cannot live in a table style:

1. In Word, put the sidebar content in a table and apply the **Sidebar Table** style (Table Design gallery). Amber `#FFF2CC` fill, thin borders, sans text, 3" wide.
2. Run the pipeline (`Set-NISTStyleArchitecture.ps1` then `Set-TableLook.ps1`). `Set-TableLook.ps1` stamps the right-margin float + text wrapping onto every Sidebar-Table-styled table and forces the 3" width; the box is the same for every sidebar and only its height changes with the content.

For sidebar *text* (not in a table), the **Sidebar** and **Sidebar Bullet** paragraph styles are managed by `Set-NISTStyleArchitecture.ps1` (sans 9pt, framed).

### Combining files into one document

`Merge-WordDocuments.ps1` combines several Word files into a single document with a section break between each. The first file is the host: its theme, styles, numbering, and header/footer parts apply to the whole output. Each following file contributes body content; its own page geometry is kept as its section, but headers/footers are unified to the host's. Source files are never modified.

```powershell
# Every .docx in the folder, in file-name order, -> <folder>\Combined.docx
.\Merge-WordDocuments.ps1 -Folder $docs

# Choose the output path
.\Merge-WordDocuments.ps1 -Folder $docs -Output "$docs\Full-Draft.docx" -Force

# Explicit order / subset (first file is the host)
.\Merge-WordDocuments.ps1 -Files a.docx,b.docx,c.docx -Output out.docx

# Start each document on an odd (right-hand) page
.\Merge-WordDocuments.ps1 -Folder $docs -SectionBreak oddPage
```

It renumbers relationship, bookmark, and footnote/endnote ids, copies and renames each file's media, accepts tracked changes, and strips comments. Open the result and press **Ctrl+A** then **F9** to refresh the table of contents and other fields. For a full NIST deliverable with roman/decimal page numbering and a style-audit gate, use `ncstar/Build-NISTReport.ps1` instead.

### Reapply styles (Word macro)

After Word's Import Styles, some paragraphs keep a stale look until you re-click the style, even with no direct formatting to clear. A file edit cannot fix this - the file is already correct and it is Word's rendering that is stale. `ReapplyAllStyles.bas` re-applies each paragraph's own style across the whole document, the same as clicking the style, but on every paragraph.

To use: in Word press **Alt+F11**, Insert > Module, paste the contents of `ReapplyAllStyles.bas`, press **F5**. It does not clear direct formatting (neither does clicking a style); use `Remove-Shading.ps1` or `Reapply-Styles.ps1` for that.

### Report builds

| Folder | What it holds |
|--------|---------------|
| `ncstar/` | NCSTAR report assembly and front-matter validation (`Build-NISTReport.ps1`, `Test-NISTFrontMatter.ps1`). See `ncstar/readme.md`. |
| `nwirp/` | NWIRP building-block and snippet builders. See `nwirp/readme.md`. |

### OoXml Module

The `OoXml/` subdirectory is a PowerShell module for generic OOXML style and numbering manipulation. The NIST style script imports this module and passes its configuration data.

```powershell
Import-Module ./OoXml/OoXml.psd1
```

**Style functions**: `Edit-StyleBlock`, `Remove-StyleBlock`, `Get-AllStyleIdsByType`, `Set-StyleBasedOn`, `Set-StyleColor`, `Edit-RunRPrInStyle`, `Set-StyleRecommended`, `Set-StyleHidden`

**Numbering functions**: `Get-HeadingAbstractNumIds`, `Set-NumberingSuffix`, `Sync-ListIndents`, `Set-BulletGlyph`

**Scriptblock variables**: `$ListSingleSpacingEdit`, `$BodyTextSpacingEdit`, `$SuppressNumbering` (passed to `Edit-StyleBlock` as mutators)

## Workflows

### Find drift across templates

```powershell
Compare-WordTemplates.ps1 -Folder "C:\path\to\templates"
# Or with a reference file:
Compare-WordTemplates.ps1 -Folder "C:\templates" -Reference "master.dotx"
```

### Edit template XML by hand

```powershell
# 1. Unpack (backs up original)
Expand-WordPackage.ps1 -File "master.dotx" -OpenEditor
# 2. Edit XML in VSCode
# 3. Repack (validates before overwriting)
Compress-WordPackage.ps1 -SourceDir "master_unpacked"
```

## Word Wildcard Find and Replace

Text fixes done in Word itself (not at the OOXML layer). In Find and Replace
(`Ctrl+H`), click **More** and check **Use wildcards**, then use these patterns.

### Space between a number and the percent sign

Inserts a space where a digit sits directly against `%`.

| Field | Value |
|-------|-------|
| Find what | `([0-9])%` |
| Replace with (regular space) | `\1 %` |
| Replace with (non-breaking space) | `\1^s%` |

To fix values that already have a regular space (`5 %`), put the space in the Find
pattern. Use the optional-space variant to normalize both spaced and unspaced values in
one pass.

| Purpose | Find what | Replace with |
|---------|-----------|--------------|
| Convert an existing regular space to non-breaking | `([0-9]) %` | `\1^s%` |
| Catch spaced and unspaced together, normalize to non-breaking | `([0-9]) {0,1}%` | `\1^s%` |

- `([0-9])` captures the digit into group 1; `\1` puts it back. `%` is literal in wildcard mode.
- `^s` is a non-breaking space and keeps the number and `%` on the same line. It works only in the Replace field, never in Find.
- The space in Find is invisible: put exactly one regular space between `)` and `%` (or before `{0,1}`). `{0,1}` means zero or one of the preceding space.
- A regular space in Find does not match a non-breaking space, so values already using a non-breaking space are skipped. The patterns are idempotent - re-running changes nothing already correct.

## Requirements

- PowerShell 5.1 or PowerShell 7+
- No Word, no COM automation (except the optional PDF export in `ncstar/Build-NISTReport.ps1` and the `ReapplyAllStyles.bas` macro, which run inside Word)

Network paths (UNC like `\\server\share`) work directly. The scripts resolve them via `ProviderPath`.

## OOXML Reference

### Key files in an unpacked Word document

| Path | What it is |
|------|------------|
| `word/styles.xml` | Paragraph, character, table, and list style definitions |
| `word/theme/theme1.xml` | Theme color palette and theme fonts |
| `word/document.xml` | Main document body content |
| `word/numbering.xml` | List, bullet, and outline numbering definitions |
| `word/header*.xml`, `footer*.xml` | Header and footer content per section |
| `word/settings.xml` | Document-wide settings |
| `word/glossary/document.xml` | Building Blocks |

### Units cheat sheet

| Property | OOXML unit | Conversion |
|----------|-----------|------------|
| Font size (`<w:sz>`) | half-points | 24 = 12pt, 28 = 14pt |
| Spacing (`<w:spacing>`) | twips (1/20 pt) | 240 twips = 12pt |
| Indent (`<w:ind>`) | twips (1/1440 in) | 720 = 0.5 in |
| Line spacing (auto) | 240ths of a line | 240 = single, 480 = double |
| Color (`<w:color>`) | hex, no `#` | `000000` = black |
| Toggle (`<w:b/>`) | presence = true | absent = inherit from base |

### Theme color slots

| Slot | Use |
|------|-----|
| `dk1`, `lt1` | Text dark, Background light |
| `dk2`, `lt2` | Text dark 2, Background light 2 |
| `accent1` through `accent6` | Accent colors |
| `hlink`, `folHlink` | Hyperlink, followed hyperlink |

Use `<w:color w:themeColor="accent1"/>` instead of `<w:color w:val="HEXVAL"/>` to make styles follow the theme.
