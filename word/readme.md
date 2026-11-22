# Word Tools

OOXML scripts for Word documents: generic engines that work on any client's files, plus
the NIST template wrappers that supply NIST configuration on top of them. The wrappers
find their engines by `$PSScriptRoot`, so both have to stay in this folder.

No Word installation is required. Everything reads and writes the OOXML directly through
`System.IO.Compression` and `System.Xml`. Pandoc is the only external binary, and only the
source-conversion scripts need it. It is installed on Windows, not in WSL, so run those
from Windows PowerShell.

## Start here

Five entry points. Everything else is either called by one of them or is a standalone
inspection tool.

| Entry point | Use it when |
|---|---|
| **author chapter set** (below) | Authors hand back a folder of chapters that has to go into the template |
| `Convert-ToNISTWord.ps1` | Source (HTML, LaTeX, Markdown, or a foreign .docx) has to become a NIST document |
| `Build-NISTDocuments.ps1` | A .docx already exists and needs the NIST template applied |
| `ncstar/Build-NISTReport.ps1` | Finished chapters have to be assembled into one deliverable |
| `restore_zotero_fields.py` | An author still uses Zotero and the templated chapters have lost their citation links. See `zotero-restore.md` |

## Author chapter set

Three steps. Do not improvise a fourth.

| Step | What runs |
|---|---|
| 1. Bring in the styles | `Build-NISTDocuments.ps1 -Folder <folder>` |
| 2. Clear the direct formatting | `Reapply-Styles.ps1 -Folder <folder>` |
| 3. Apply the styles | put every paragraph on a template style |

Step 3 in full, in this order. Each one is idempotent, so a step can be repeated after a
correction without undoing the others.

| | Runs | Does |
|---|---|---|
| 3- | `Import-WordStyles.ps1 -Folder <f> -Source <ref> -Style <ids>` | any style the architecture only EDITS that this chapter does not have |
| 3a | `Get-UnstyledContent.ps1 -Folder <f>` | read-only: classifies what has no style, and reports the classes with no safe default |
| 3b | `Set-StyleRemap.ps1 -Folder <f>` | author styles to template styles; unstyled body to `BodyText`, unstyled cells to `TableText`; leaves caption-shaped text for 3d |
| 3c | `Set-TableParagraphStyle.ps1 -Folder <f>` | every paragraph inside a table to `TableText` |
| 3d | `Set-ContentStyles.ps1 -Folder <f>` | figures to `graphic`, captions to `CaptionFigure`/`CaptionTable`, reference entries to `References`, list items to `ListBullet`/`ListNumber`, and the caption prefix to label + NBSP + number + period + TWO spaces |
| 3e | `Set-NISTChapterNumbering.ps1 -File <ch> -StripTypedNumber` | per chapter, with its own number: builds `HMList` when the file has none and wires `Heading1-4` to it |
| 3e1 | `Set-HeadingShift.ps1 -File <ch> -Shift 1 -Title "..."` | ONLY when Heading 1 is used as the section level. Run BEFORE 3e: with eight Heading 1 paragraphs the chapter cannot be numbered, and with no chapter number every caption's STYLEREF has nothing to read |
| 3e2 | `Set-NISTTableCaptions.ps1 -File <ch> -Style CaptionFigure -Label Figure -SeqName Figure -UseStyleRef -InPlace -StripSourceCaption`, then again with `-Style CaptionTable -Label Table -SeqName Table` | renumbers captions that ALREADY exist into canonical form. ONE LABEL PER RUN, so a chapter needs two. `-InPlace` trusts the caption style instead of adjacency, which matters because a figure rendered as a table defeats `-AfterGraphic`; it now defaults `-FromStyle` to `-Style`, and the earlier line here omitted both, which left the command looking for a Heading 1 beginning with "Figure", reporting `Nothing to convert` on all eight chapters and exiting 0. No `-Chapter`: the cached STYLEREF number is read from the document's own `HMList` |
| 3e3 | `Join-CaptionFragment.ps1 -Folder <f> -IncludeUnstyled` | a caption written in two halves, where the author marked the orphan half with a leading `^` |
| 3e4 | `Set-EmptyHeadingText.ps1 -Folder <f>` | a heading with no text is an invisible section boundary; this makes it a reported placeholder instead |
| 3f | `Set-DocTitle.ps1 -Folder <f> -Title "<volume>"` | `dc:title`, which the footer and the banner TITLE field both read |
| 3g | `Set-NISTHeaderFooterContent.ps1 -File <ch> [-FrontMatter]` | the six header/footer parts; `-FrontMatter` on the front-matter file |
| 3h | `Set-NISTBannerTable.ps1 -Folder <f>` | the page-anchored colour banner above each chapter's Heading 1 |
| 3i | `Repair-WordFields.ps1 -Folder <f>` | dead REFs, empty SEQ husks, instruction-less fields, Google content controls, stray `/>` |
| 3j | `Remove-DirectParagraphFormat.ps1 -Folder <f> -IncludeUnstyled -ParagraphMark` | direct spacing, indents and justification, and the paragraph mark's own run properties |
| 3k | `Set-TableLook.ps1 -Folder <f> -TableStyle NIST-Data -FixedWidth` | content tables onto `NIST-Data` at a pinned width; leaves the banner alone |
| 3k1 | `Set-TableCrossReferences.ps1 -File <ch> -Label Figure` then `-Label Table` | typed "see Table 5-10" mentions into live REF fields. Run AFTER the caption numbers are right: it matches a mention against the number the caption currently shows |
| 3k2 | `Set-FieldsDirty.ps1 -Folder <f>` and again with `-Part '^word/document\.xml$' -Instruction 'SEQ|STYLEREF|REF'` | so Word recomputes the running heads, the caption numbers and the cross-references on open, with no Word needed here |
| 3l | `Test-NISTDocument.ps1 -Folder <f>` | the suite |

**Section, appendix and chapter references wait for the ASSEMBLED document.** A REF field
reaches a bookmark in its own document only, and in a chapter set those targets are in
other files. NWIRP 2-1 carries 86 typed ones and only 9 are same-file (ch 4 to 4.5.10,
ch 6 to 6.2/6.3/6.3.3, ch 7 to 7.1-7.5); the other 77 are 19 cross-chapter sections, 44
`Chapter N`, and 13 `Appendix A`-`H` whose appendices are not in the chapter that cites
them - ch 4 asks for seven of them and holds two appendix headings, neither lettered.
Wiring those per chapter produces dead references. Do them after `Merge-WordDocuments.ps1`
or `ncstar/Build-NISTReport.ps1`, against the merged file, and note that
`Set-TableCrossReferences.ps1` handles Figure and Table only - a heading reference has to
resolve "Section 6.3.3" through the HMList ladder rather than by matching caption text.

Do NOT number a front-matter chapter (3e). Its Heading 1 is "Preface", and the chain would
render "Chapter 0". Do not number a chapter whose Heading 1 is used as its SECTION level
either - one chapter here has eight of them and would read "Chapter 4" through
"Chapter 11". Check with `Test-NISTDocument.ps1` first: one Heading 1 per chapter.

If a pass turns out to have damaged specific paragraphs, `Restore-WordParagraphs.ps1`
brings them back out of `orig\` by matching their text, then the idempotent steps above
are re-run over them. That is the second reason to keep `orig\`.

Step 2 is one call and it is the whole of the cleanup. It wipes `w:rPr` from every run
across document, headers, footers, footnotes and endnotes, preserving italic,
superscript/subscript, highlighting and footnote reference marks, and leaving MathType and
OLE runs entirely alone. It is idempotent. Dismantling direct formatting property by
property instead - sizes, then fonts, then justification, then indents, then colors - is
slower, misses things, and is how real damage gets done.

Before step 1:

- Copy the folder to `<folder>\orig\` once. Every later claim about what changed is
  measured against it, and `GetFiles()` is not recursive so the pipeline ignores it.
- Check for `~$` lock files. One means an author has that chapter open: a write either
  fails or loses their work, and leaves a half-written `.tmp` behind. Skip those files and
  say which.
- Confirm the toolkit is one version. Two copies of the same script in one tree is how an
  older `Set-NISTStyleArchitecture.ps1` got run and wrote a `framePr` into eight chapters
  in one pass: `Get-ChildItem -Recurse -Filter 'Set-NIST*.ps1' | Group-Object Name | Where-Object Count -gt 1`
- Survey with `Test-NISTDocument.ps1 -Folder <folder>`. The number that matters most is how
  many paragraphs carry no `pStyle` at all, split into table cells, empty spacers and real
  body text. Those need three different treatments.

### Styles the architecture only edits

`Set-NISTStyleArchitecture.ps1` writes full definitions for the `xx*` family,
Header/Footer, `Versohead`, `Title`, `MTDisplayEquation`, `NIST-Data`, `SidebarTable`, the
List Number family when it is missing, and the appendix ladder (AppendixHead 1-4 on
"Appendix A." / A.1. / A.1.1. / A.1.1.1.). It only *edits* these, and no-ops when they are absent:

`BodyText` · `CaptionTable` · `CaptionFigure` · `graphic` · `TableText` ·
`TableandFigureFootnotes` · `FootnoteText` · `HMList`

`Quote` is in that list too. An edit pass NO-OPS SILENTLY when the style is absent, so a
chapter missing one gets no warning and no style: Ch04 of NWIRP 2-1 had no `Quote`, and the
pass that would have moved its seven block quotations reported "nothing to change".

`Import-WordStyles.ps1` brings them across verbatim, following `basedOn`/`next`/`link`
transitively so a linked character style (`QuoteChar`) comes with its paragraph style and
no new dangling reference is left behind. It never imports `xxgraphic`. Run the
architecture afterwards: an imported style's `numPr` points at a `numId` that means
something else in the target.

**Then check what every style in use inherits FROM.** A paragraph style based on `Normal`,
or on nothing, inherits none of the template - not the face, not the size, not the
spacing - so text on it renders as Normal however carefully it was applied. Twenty-one
large blocks of body prose went out that way in one chapter with every other check
passing. The style is kept and given a parent, not replaced.

### Order dependencies that are not obvious

**Remap `document.xml` before the architecture runs.** The architecture never edits
`document.xml`, and it *skips* a rebase whose target style is absent - it warns
"Rebase SKIPPED", so read the output rather than tailing it. A paragraph pointing at an
author style the architecture deletes falls back to Normal.

**Identify captions before anything else claims them.** Putting unstyled paragraphs on
`BodyText` moves caption text onto `BodyText` too, and a later "does this figure have a
caption?" test that looks for a caption *style* then finds none and inserts placeholders
beside captions that were already there. Order: style the captions that exist, then insert
placeholders only where one is genuinely missing, then convert typed numbers to fields,
then normalise to two spaces after the number.

A caption is label + number + period/colon + title. "Figure 5 shows the outages" is prose
- in one chapter 25 of 55 candidates were. Letter-keyed captions ("Table B.1",
"Figure P-1") keep literal numbers: a chapter SEQ field cannot produce "B".

**Measure caption text with every text-producing node, not just `w:t`.** A caption in
correct form reads `Table ` + STYLEREF + `<w:noBreakHyphen/>` + SEQ + `.  Title`. Read
only the `w:t` elements and that comes back as "Table 32.  Title" - the hyphen missing -
so every offset past it is one short, and a prefix rewritten at those offsets lands inside
the number: "Table 3-1." became "Table 31.  -". It cost 12 captions across two chapters.
`<w:noBreakHyphen/>`, `<w:tab/>`, `<w:br/>` and `<w:softHyphen/>` all have to be in the
model, and none of them may be edited. Note also that `'<w:tab/>'.StartsWith('<w:t')` is
TRUE: match the element as `^<w:t[ >]` or a tab is taken for a text run.

**The number may continue in literal text after the field.** Both forms appear in the same
chapter set: `SEQ` alone ("Figure 1"), and STYLEREF + hyphen + typed digits ("Figure 2-9",
where only the "2" is a field). Treating the last field end as the end of the number put
the period inside it. Absorb a separator-plus-digits run after the field first - and only
that, so a title beginning with a year ("Figure 5.  2020 Census") keeps its year.

**Chapter numbering now builds `HMList` when it is missing.** Heading numbering belongs to
that named multilevel list, not to numbering applied per heading, and the `pStyle` links
inside its `ilvl 0-3` are what put it in Word's Multilevel List gallery. An author chapter
has its own multilevel list instead, numbering `Heading1` as "4.%1." - which is why the
chapter number has to come from the filename and `ilvl 0` has to start at it.

**An imported list style carries the TEMPLATE's numId.** This is the trap that made list
bullets and numbers wrong three times. A `numId` that RESOLVES is not a `numId` that
resolves to the right thing: in an author document that id already belongs to one of the
author's own lists, so an imported `List Bullet 2` silently inherited "Appendix %1.",
"Recommendation %1.  ", or a heading ladder's "Chapter %1.". Every one resolved cleanly,
so an existence-only check passed all of them. `Set-NISTStyleArchitecture.ps1` only
checks that a list style HAS numbering (creating the List Number family when it is
missing); it does not check what that numbering is. After importing list styles, look at
the marker each one renders in Word before trusting it.

**Direct formatting is two jobs, not one.** `Reapply-Styles.ps1` clears direct RUN
formatting and nothing else: its pattern only matches an `rPr` that follows a `<w:r>` open
tag. Direct PARAGRAPH formatting (`w:spacing`, `w:ind`, `w:jc`) and the paragraph MARK's
own `rPr` both survive it, which is what Word's inspector keeps reporting as
"Body Text + Justified, First line: 0.25"" and "Normal + 11 pt" on an empty paragraph.
`Remove-DirectParagraphFormat.ps1` is the second job and is deliberately not in any
blanket pipeline.

**The chapter number lives in the document, not in the filename.** `HMList`'s `ilvl 0`
`w:start` IS the number the headings render, so it survives a rename and makes a re-run a
no-op. `Set-NISTChapterNumbering.ps1` takes `-ChapterNumber` when the document has no
number yet, reads the document's own otherwise, and REFUSES rather than guess - and
`Test-NISTDocument.ps1` reads the same place.

Neither reads the filename. A filename is metadata anyone can change, and the number ends
up inside the document: an earlier version inferred it from `Chapter-NN` and the suite used
a stricter pattern than the writer did, so a chapter spelled `Chapter_01` was numbered from
one pattern and checked by neither. `ncstar/Build-NISTReport.ps1` still sorts files by name,
which is fine - that is ordering files, not stamping a number into one.

**Copy definitions, never compose them.** Styles, list ladders and numbering chains come
out of the authoritative source verbatim. A hand-built equivalent comes out almost right,
differs in ways nobody notices for a week, and has to be redone.

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

### Not in this pipeline, and why

These are excluded from the automated step list because each needs a decision, not because
they are off limits. `Reapply-Styles.ps1` in particular is **step 2 of the author chapter
set process above** - excluded here means "choose when to run it", not "never run it".

| Script | When to run it |
|---|---|
| `Reapply-Styles.ps1` | Step 2 on any set authors have edited |
| `Remove-Shading.ps1` | Only on a file with unwanted direct shading. `-IncludeTables` would strip hundreds of deliberate cell fills |
| `Set-NISTHeaderFooterContent.ps1` | When the running heads need rebuilding. Use `-SkipFieldUpdate` unless you have checked for runaway bookmarks |
| `Set-TableParagraphStyle.ps1` | Whenever cells are on the wrong style. `Convert-ToNISTWord.ps1` does call it in stage 4 |

## NIST wrappers

| Script | What it does |
|---|---|
| `Set-NISTTheme.ps1` | Writes the NIST entity palette into `theme1.xml` |
| `Set-NISTStyleArchitecture.ps1` | Rebases styles onto the `xx*` family, deletes junk styles, applies theme colors, sets bullet glyphs, curates Recommended/Hidden |
| `Set-NISTTemplateStyles.ps1` | Thin wrapper over `Set-NISTStyleArchitecture.ps1` |
| `Set-NISTChapterNumbering.ps1` | Restores the `HMList` chain and the `numPr` on Heading1-4. Needs `HMList` to exist |
| `Set-NISTHeaderFooterContent.ps1` | Installs the six even/odd/first header and footer parts |
| `Set-NISTTableCaptions.ps1` | Writes NIST table and figure captions with SEQ fields. Run `-ListOnly` first on a new document set |
| `Set-NISTThreeColumnTableWidths.ps1` | Normalizes 3-column body tables to a fixed layout |
| `Build-NISTReferenceDoc.ps1` | Builds the pandoc reference doc. Pandoc resolves styles by NAME, not styleId |
| `Test-NISTDocument.ps1` | The check suite. `-File` or `-Folder`. Run it before and after everything |
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
| `Set-DocTitle.ps1` | Sets `dc:title` metadata. Run before `Set-NISTHeaderFooterContent`, whose footer reads it |
| `Set-FootnoteStyles.ps1` | Footnote paragraphs onto `FootnoteText`, reference marks onto `FootnoteReference` |
| `Set-ThemeColors.ps1` | Replaces a color scheme from a supplied `clrScheme` block |
| `Install-ThemeColors.ps1` | Installs a theme color file into the user's Word gallery |
| `Repair-Typography.ps1` | Straight quotes to curly, on visible text only |
| `Repair-HtmlInlineFormatting.ps1` | CSS-only inline formatting in HTML into real tags |
| `Repair-NumberingDefinitions.ps1` | Duplicate abstractNum/num ids and the schema order. Fixes vanished bullets and a missing multilevel list |
| `Remove-EmbeddedFonts.ps1` | Drops bundled fonts |
| `Remove-DirectRunSize.ps1` | One direct run size, in named styles only. Keeps bold, italic and color |
| `Remove-Shading.ps1` | Direct run and paragraph shading only |
| `Reapply-Styles.ps1` | Clears direct character formatting, keeping italic, superscript/subscript, highlighting, footnote marks and equation runs. Idempotent |
| `Set-FieldsDirty.ps1` | Marks fields dirty so Word recomputes them on open. Per-field, so headers can be refreshed without touching the body |
| `Set-HeadingShift.ps1` | Shifts a heading ladder a level deeper or shallower, deepest-first. Refuses when the target style is absent |
| `Join-CaptionFragment.ps1` | Merges a marked caption half back into the caption above it, moving runs so inline formatting survives |
| `Set-EmptyHeadingText.ps1` | Gives an empty heading a reported placeholder. Inserts inside a `_Ref` bookmark so a cross-reference still resolves |
| `Restore-LostImages.ps1` | Puts back an image whose a:blip was dropped while its media part and relationship survived. Finds them as orphaned media |
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

## Traps

Each of these has produced a wrong conclusion or real damage.

**Whitespace in self-closing tags.** `<w:i/>` and `<w:i />` are the same element and one
document set contains both. This caused four separate wrong conclusions in a single
session: a missed style match, a stray `/>` left as text inside `w:pPr`, a false
"60 lost italic reference titles" alarm, and a stray `/>` written into `numbering.xml` in
six files. Match `<w:i(?:\s[^>]*?)?\s*/>`, and when replacing match the **whole element** -
a pattern ending at the last attribute quote leaves the original's `/>` behind.

**Run counts are not content.** Merging two adjacent italic runs lowers the run count
without losing a character. Counting runs reports damage that did not happen and hides
damage that did. Compare the concatenated text of runs carrying the property.

**Shared abstractNums.** `ListNumber` and `ListBullet2` pointed at the same `abstractNum`
in one chapter; setting `numFmt="bullet"` to fix the bullets turned the numbers into
bullets. Check whether anything else references an `abstractNum` before rewriting it.

**`ilvl 0` names a pStyle.** A level naming `TableBullet` or `Heading1` is bound to *that*
style, not to whichever style points at the `abstractNum`. Five of eight chapters had
bullets attached to the wrong style this way.

**Latent built-ins.** Word creates `FootnoteText` as `basedOn Normal` at 10pt the instant
someone inserts a footnote. If the architecture does not name it, nothing corrects it - and
the curation pass marks it `semiHidden`, so it is applied, working, and invisible in the
Styles pane, which reads as "this chapter has no footnote style".

**Empty paragraphs can be load-bearing.** An untitled `AppendixHead` is still a section
boundary. Deleting two merged a reference list into a nomenclature section and orphaned a
live cross-reference bookmark.

**Cloning XML carries its ids.** Copying a banner table between documents brought a stray
`_Hlk` bookmark with it and gave seven files a duplicate `w:id="0"`. Regenerate
`w14:paraId`/`w14:textId` and strip or renumber bookmarks when cloning.

**`-File` does not parse arguments.** `powershell.exe -File s.ps1 -ParagraphStyle A,B,C`
passes one string, so a `[string[]]` parameter matches nothing - silently. A removal
reported "none found" while 127 sat in plain sight. One call per value, or
`-Command "& '<path>' -ParagraphStyle A,B,C"`.

**`References` is not a list.** Based on `ListParagraph` it inherits that style's list
indents, pushing the first line in 0.2-0.25" while the Styles pane shows nothing on
`References` itself. It belongs on `xxbase` with the hanging indent written explicitly.

**`pageBreakBefore` on a style** forces a break before every paragraph wearing it. On
`Heading1` in a chapter with eight of them, that is eight page breaks.

## Guard every write

Snapshot before, compare after, refuse to save on a mismatch: visible text identical
(allowing typographic quote substitution), italic and superscript content unchanged **by
content**, XML parses, bookmark ids pair up, image count unchanged. A script that refuses
to write is worth more than one that is usually right.

Two things learned the hard way about the parse check itself.

**Trim the BOM for the probe, and only for the probe.** `$utf8NoBom.GetString()` decodes a
UTF-8 BOM into a leading `U+FEFF` - the `$false` only stops it EMITTING one - and
`XmlDocument.LoadXml` rejects that with `Data at the root level is invalid. Line 1,
position 1.` Twenty-six of the toolkit's twenty-nine probe sites never trimmed it, so on
any file Word itself last wrote, the guard meant to protect the write was the only thing
stopping it. It stayed hidden because a file the toolkit had already written carries no
BOM: the fault only shows on author-fresh input, which is the input that matters. Trimming
the string the file is WRITTEN from instead would silently drop the BOM from the output.

**Pick an invariant the script cannot legitimately break.** Visible-text equality is the
right one for a style remap and the wrong one for `Set-TableCrossReferences`, which
corrects a stale mention on purpose; that script guards on structure instead - protected
element counts, field begin/end balance against the number of edits, no stray markup - and
reports the text delta rather than asserting it away.

## Leave alone

- **Tracked changes and comments.** Authors work in these files; they are the normal
  mid-work state, not defects. They do constrain repairs: a citation field whose
  `instrText` sits inside a `w:del` cannot be unlinked without falsifying the revision.
- **A file with a `~$` lock.**
- **Importing the text into a clean document while authoring is live.** It discards every
  tracked change, comment, cross-reference bookmark and image anchor. That is the right
  move only once revisions are accepted.

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

`NIST-DOCUMENT-CHECKLIST.md` is the per-document checklist `Test-NISTDocument.ps1`
implements. `zotero-restore.md` is the process for putting Zotero citation and
bibliography fields back from the author's originals. `nist-color-palette.md` is the palette. `converting-html-latex.md` covers the
conversion route in more detail. `ReapplyAllStyles.bas` and `word-table-format.py` predate
the OOXML tools and are not called by anything here.
