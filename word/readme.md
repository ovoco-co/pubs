# Word Tools

OOXML scripts for Word documents: generic engines that work on any client's files, plus
the NIST template wrappers that supply NIST configuration on top of them. The wrappers
find their engines by `$PSScriptRoot`, so both have to stay in this folder.

No Word installation is required. Everything reads and writes the OOXML directly through
`System.IO.Compression` and `System.Xml`. Pandoc is the only external binary, and only the
source-conversion scripts need it. It is installed on Windows, not in WSL, so run those
from Windows PowerShell.

This file is generated from what the entry points actually call. If a script here is not
listed below, nothing in the pipeline invokes it.

## Start here

Four scripts are entry points. Everything else is either called by one of them or is a
standalone inspection tool.

| Entry point | Use it when |
|---|---|
| `Convert-ToNISTWord.ps1` | Source (HTML, LaTeX, Markdown, or a foreign .docx) has to become a NIST document |
| `Build-NISTDocuments.ps1` | A .docx already exists and needs the NIST template applied |
| `ncstar/Build-NISTReport.ps1` | Finished chapters have to be assembled into one deliverable |
| `restore_zotero_fields.py` | An author still uses Zotero and the templated chapters have lost their citation links. See `zotero-restore.md` |

## Convert-ToNISTWord

Runs in five stages. Stage 3 is the whole of `Build-NISTDocuments.ps1`, so converting a
document does everything formatting one would, plus the conversion and table work around it.

| Stage | What runs | Why it sits here |
|---|---|---|
| 0 | `Repair-HtmlInlineFormatting.ps1` | HTML only, on a throwaway copy. flextable emits no semantic tags, so bold and superscript live in CSS classes that pandoc cannot see |
| 1 | pandoc against `NIST-reference.docx` | Build that reference doc first with `Build-NISTReferenceDoc.ps1`, once per template revision |
| 2 | remap pass (in-script) | Folds pandoc's style ids onto NIST ones. Must run before stage 3: the style step deletes styles it does not recognize but never edits `document.xml`, so a paragraph left pointing at a deleted pandoc style falls back to Normal |
| 3 | `Build-NISTDocuments.ps1` | The build pipeline below |
| 4 | table finishing | Optional polish, listed below |

Stage 4 runs in this order, and the order matters:

| Step | Script | Note |
|---|---|---|
| 1 | `Set-TableLook.ps1` | Re-run after the build so its options are the last thing written |
| 2 | `Set-TableParagraphStyle.ps1` | Before the indents: it preserves `w:ind` only on paragraphs it leaves in place |
| 3 | `Restore-TableIndents.ps1` | Rebuilds the row hierarchy pandoc drops |
| 4 | `Merge-TableFootnotes.ps1` | Claims the note rows before anything can shade them |
| 5 | `Set-TableBannerRows.ps1` | After the footnotes, so whatever is still full-width is a banner |
| 6 | `Set-TableAccentBar.ps1` | Only when `-AccentBarRules` or `-AccentBarColor` is passed |
| 7 | `Set-NISTTableCaptions.ps1` | Last: the caption style only exists once stage 3 has applied the style architecture |

Each stage 4 step is caught individually, so one failing does not abandon the rest.

## Build-NISTDocuments

Seven steps, in this order, all idempotent and safe to re-run:

| Step | Script |
|---|---|
| 1. Theme | `Set-NISTTheme.ps1` |
| 2. Styles | `Set-NISTStyleArchitecture.ps1` |
| 3. Page layout | `Set-PageLayout.ps1` |
| 4. Table look | `Set-TableLook.ps1` |
| 5. 3-column widths | `Set-NISTThreeColumnTableWidths.ps1` |
| 6. Typography | `Repair-Typography.ps1` |
| 7. De-embed fonts | `Remove-EmbeddedFonts.ps1` |

Step 1 must precede step 2: the styles reference the theme slots the palette creates.

This script takes `-Folder` only, not `-File`, so it formats every .docx in the folder.
Stage the files you want into their own folder first. Running it across a mixed folder is
not harmless: the style step deletes styles it does not recognize, so a file that was
already correct, or one carrying styles from pasted content, loses them. The individual
step scripts all accept `-File` if you need to treat one document on its own.

### Deliberately not in this pipeline

`Remove-Shading.ps1`, `Reapply-Styles.ps1`, and `Set-NISTHeaderFooterContent.ps1` are
excluded on purpose. They destroy intentional formatting or need a per-file decision. Run
them yourself when a specific file calls for it. Do not add them to the step list.

`Set-TableParagraphStyle.ps1` is also outside this pipeline, but `Convert-ToNISTWord.ps1`
does call it in stage 4, because pandoc leaves table cells on the wrong styles.

## NIST wrappers

| Script | What it does |
|---|---|
| `Set-NISTTheme.ps1` | Writes the NIST entity palette into `theme1.xml` |
| `Set-NISTStyleArchitecture.ps1` | Rebases styles onto the `xx*` family, deletes junk styles, applies theme colors, sets bullet glyphs |
| `Set-NISTTemplateStyles.ps1` | The same architecture for a CLEAN document, without the legacy-fix passes |
| `Set-NISTHeaderFooterContent.ps1` | Installs the six even/odd/first header and footer parts |
| `Set-NISTTableCaptions.ps1` | Writes NIST table and figure captions with SEQ fields. Run `-ListOnly` first on a new document set |
| `Set-NISTThreeColumnTableWidths.ps1` | Normalizes 3-column body tables to a fixed layout |
| `Build-NISTReferenceDoc.ps1` | Builds the pandoc reference doc. Pandoc resolves styles by NAME, not styleId |
| `Audit-NISTStyleSet.ps1` | Pre-assembly gate: checks the NIST style ids downstream scripts depend on |
| `Install-NISTThemeColors.ps1` | Per-user Word gallery entry. Does not touch documents |

## Generic tools

Table and document formatting:

| Script | What it does |
|---|---|
| `Set-TableLook.ps1` | Table style options and repeating header rows. `-ReportWidths` measures columns and suggests widths |
| `Set-TableBannerRows.ps1` | Shades and bolds the full-width group-label rows |
| `Set-TableAccentBar.ps1` | Colored bar down the first cell of every row, the primary population code |
| `Set-TableDecimalTabs.ps1` | Decimal tab stops for numeric columns. Only correct on fixed-layout tables |
| `Set-TableParagraphStyle.ps1` | Puts cell paragraphs on one style without touching run formatting |
| `Set-TableCrossReferences.ps1` | Turns typed "see Table 5-10" mentions into live REF fields on the caption. `-ListOnly` first |
| `Merge-TableFootnotes.ps1` | Collapses trailing note rows into one unshaded row |
| `Restore-TableIndents.ps1` | Restores table-cell indentation pandoc drops |
| `Set-PageLayout.ps1` | Margins and header/footer distance. Defaults are the NIST values |
| `Set-DocTitle.ps1` | Sets `dc:title` metadata |
| `Set-ThemeColors.ps1` | Replaces a color scheme from a supplied `clrScheme` block |
| `Install-ThemeColors.ps1` | Installs a theme color file into the user's Word gallery |
| `Repair-Typography.ps1` | Straight quotes to curly, on visible text only |
| `Repair-HtmlInlineFormatting.ps1` | CSS-only inline formatting in HTML into real tags |
| `Remove-EmbeddedFonts.ps1` | Drops bundled fonts |
| `Remove-Shading.ps1` | Direct run and paragraph shading only. `-IncludeTables` would strip hundreds of deliberate cell fills |
| `Reapply-Styles.ps1` | Strips direct character formatting. Post-authoring cleanup, not a build step |
| `Merge-WordDocuments.ps1` | Combines files with section breaks, originals untouched |

Inspection and audit, none of which modify a document:

| Script | What it does |
|---|---|
| `Get-WordStyles.ps1` | Every style in a file, as Markdown |
| `Get-WordTableStyles.ps1` | Every table style with its conditional formatting blocks |
| `Get-WordParts.ps1` | Every Building Block in a template's glossary |
| `Audit-StyleSet.ps1` | Checks a file for required style ids |
| `Audit-WordTheme.ps1` | Finds styles and runs that hardcode hex instead of using the theme |
| `Compare-WordTemplates.ps1` | Template drift, read from the OOXML |
| `Expand-WordPackage.ps1` | Unpacks a .docx into editable XML |
| `Compress-WordPackage.ps1` | Repacks that folder into a valid Word file |

## OoXml module

`OoXml/OoXml.psd1` is the manifest. `Convert-ToNISTWord.ps1` and
`Set-NISTStyleArchitecture.ps1` import it. It holds the low-level helpers:
`Styles/` edits style blocks, run properties, colors, and visibility.
`Numbering/` handles abstractNum definitions, bullet glyphs, suffixes, and list indents.

A function here that returns a collection must use `return ,$collection`. Without the
comma an empty collection comes back as `$null` and a one-element collection comes back
as the bare element. That has bitten this module twice.

## Document-set folders

`ncstar/` and `nwirp/` hold build scripts for the two NIST document sets. They are
PowerShell only. `ncstar/Build-NISTReport.ps1` assembles a deliverable,
`ncstar/Test-NISTFrontMatter.ps1` audits a Chapter-00 against the canonical sequence, and
the two `nwirp/Build-NWIRP*` scripts build the building-block set and the snippets doc.

## Conventions

Verb-noun names, PowerShell 5.1 and up. Client-specific scripts carry the client in the
noun (`Set-NISTTheme`), generic ones do not (`Set-ThemeColors`). Scripts back up before
writing, set `$ErrorActionPreference = 'Stop'`, and ask before acting unless passed
`-Force`. Most take `-File` or `-Folder`.

## Also in this folder

`converting-html-latex.md` covers the conversion route in more detail.
`ReapplyAllStyles.bas` and `word-table-format.py` predate the OOXML tools and are not
called by anything here.
