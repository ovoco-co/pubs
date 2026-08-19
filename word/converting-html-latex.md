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

### What the three stages do

**Stage one, pandoc.** Converts each source file to `.docx` against the reference document, so the output starts with the NIST theme, page setup, headers and footers, numbering, and styles. Images are embedded, and `--extract-media` also drops the originals in a `media` folder beside the output. LaTeX math becomes native Word equations (OMML).

**Stage two, remap.** Folds pandoc's styles into NIST styles, assigns the `NIST-Data` table style to every table, and deletes the pandoc style definitions that nothing references any more.

**Stage three, build.** Runs `Build-NISTDocuments.ps1 -Clean` on the output folder: theme, styles, page layout (1" margins), table look, typography, and font de-embedding.

Stage two runs before stage three on purpose. `Set-NISTStyleArchitecture.ps1` deletes styles it does not recognize but never edits `document.xml`, so a paragraph still pointing at a deleted pandoc style would quietly fall back to Normal. The remap rewrites those references first.

## Step three: finish the document

The conversion does not do everything. After it runs:

- **Table indentation.** Run `Restore-TableIndents.ps1`. Statistical tables carry their row hierarchy as a leading tab (in a `.docx` source) or a leading space (from HTML), and pandoc drops the tabs entirely. Without this the group labels and their members all sit flush left and the grouping is gone. With a `.docx` source pass `-Source` so the levels are read from the original; from HTML, run it on the target alone and it reads the leading spaces.
- **Table text.** Run `Set-TableParagraphStyle.ps1` to move cell paragraphs onto `TableText`. The converter assigns the table style but leaves cell paragraphs on whatever pandoc chose, which is Body Text, or Quote for anything it read as a block quote. Quote carries a border and a fill that look wrong inside a cell. The script rewrites the paragraph style only, so bold, superscripts, subscripts, and indentation survive.
- **Table captions.** Run `Set-NISTTableCaptions.ps1 -ListOnly` first, then without it. It converts the headings sitting above tables into `Caption_Table` with a `SEQ` number. Check the listing: a section heading that happens to sit above a table matches the same rule as a caption.
- **Headers and footers.** Run `Set-NISTHeaderFooterContent.ps1` per file, with `-FrontMatter` on the front-matter file and without it on chapters.
- **Document title.** Run `Set-DocTitle.ps1`, which prompts per file.
- **Captions.** Pandoc's captions land on the `Caption` style but carry no Word caption fields, so cross-references and a table of figures need the fields inserted in Word.
- **Display equations.** Math converts correctly but does not pick up the `MTDisplayEquation` style. Apply it where you want NIST equation spacing.

## Options

### Convert-ToNISTWord.ps1

| Flag | What it does |
|------|-------------|
| `-File` / `-Folder` | One source file, or every supported file in a folder. |
| `-ReferenceDoc` | Reference document to use. Defaults to `NIST-reference.docx` beside the script. |
| `-OutputFolder` | Where the `.docx` files go. Defaults to the source folder. |
| `-TableStyle` | Table style for converted tables. Defaults to `NIST-Data`. Pass `''` to leave pandoc's table styling alone. |
| `-RemapLists` | Move lists onto the NIST List Bullet and List Number ladders. Off by default, see the caveat below. |
| `-Bibliography` | Bibliography file (`.bib`, `.json`, `.yaml`). Turns on `--citeproc`. |
| `-Csl` | CSL style file for citeproc. |
| `-NoMedia` | Skip `--extract-media`. Images are embedded either way. |
| `-PandocArgs` | Extra arguments passed to pandoc verbatim. |
| `-PandocPath` | Path to `pandoc.exe` if it is not on PATH. |
| `-SkipRemap` | Skip stage two, leaving pandoc's own styles in place. |
| `-SkipPipeline` | Stop after stage two, so you can inspect the files before the NIST build runs. |
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

**Everything in the output folder got reformatted** - stage three processes every Word file in `-OutputFolder`. Use a dedicated folder.

## See also

- `readme.md` in this folder for the rest of the Word toolkit and the manual NIST build sequence
- `Build-NISTDocuments.ps1` for what stage three actually runs
- `publishing/md-docx/` for generic pandoc wrappers with no template awareness, useful for throwaway or interchange work
