# clean-template

Moves author chapters into a clean NIST template in a fixed, checked order,
keeping every tracked change and comment.

The rule: **every run starts from the untouched originals**, every decision
lives in `settings.json`, and nothing is fixed by hand between steps. A new
fix becomes a numbered step here, with a check after it.

## Run it

```
python run_pipeline.py "C:\Users\Artist\Documents\HM-261113"
```

Close Word first (steps 03, 11, 12 and 13 drive Word). Each Word step has
`word_step_timeout_minutes`: hidden Word can wait forever on a dialog no one
sees, so past that its Word is ended and the run stops, showing the last file
the step finished. Uses the pubs `.venv` Python (lxml,
pywin32). Output goes to `<originals>\run\`:

| Folder | What is in it |
|---|---|
| `originals` | the chapters the run used, copied from the source folder minus `exclude` patterns (compiled whole-report copies); every step reads these |
| `00-template` | `NIST-Template.docx`, built fresh from `template_source` |
| `01-strip` … `13-update-fields` | the chapters as each step left them |
| `final` | the finished chapters, plus the template they were built on |
| `run-report.md`, `run-log.txt` | what ran, and every command's output |

Each chapter's own document properties (title, custom properties) are carried
into its new file in step 03: the chapter banner and running heads print the
TITLE property.

`01-strip\strip-report.md` lists what the strip removed and what it kept for
review (boxed paragraphs, directly numbered paragraphs, extra styles).

## Steps

| # | Step | Script | What it changes |
|---|---|---|---|
| 00 | template | `..\word\Set-NISTStyleArchitecture.ps1`, `make_empty_template.py` | Style architecture onto the template source, then an empty copy: styles, numbering, theme, page setup, headers and footers |
| 01 | strip | `strip_formatting.py` | Hand formatting out (fonts, sizes, color, underline, spacing, indents, alignment, tracked formatting changes). Keeps styles, bold/italic/super/subscript, highlight, tables' indents and alignment, callout borders and shading, equations, symbol fonts, direct numbering. Maps styles per `style_map` |
| 02 | remove toc | `remove_toc.py` | Table of contents and lists of tables/figures out of chapter files, with their headings, the "Front Matter" label, the page breaks around them and the emptied front-matter section. Files matching `keep_toc` (front matter) keep theirs. Stops on a block holding a comment, footnote, image, equation or table |
| 03 | move | `move_into_template.py` | Each chapter inserted into a fresh template copy with Word; the template's style definitions win; track changes left on |
| 04 | table text | `set_table_text.py` | Data-table paragraphs onto Table Text. Run formatting untouched |
| 05 | table style | `apply_table_style.py` | `table_style` on every data table, same options everywhere (header row, banded rows; total rows an author set are kept), header row repeats, direct borders/shading/cell margins/text color removed, widths pinned to the grid (fit to 6.5"), Sidebar Table callouts floated right at 3" |
| 06 | table notes | `..\word\Merge-TableFootnotes.ps1` | Trailing full-width note rows -> one unshaded row on Table and Figure Footnotes |
| 07 | group rows | `..\word\Set-TableBannerRows.ps1` | Full-width group-label rows shaded `group_row_fill` and bold. After 06, so notes are not mistaken for groups |
| 08 | decimal tabs | `..\word\Set-TableDecimalTabs.ps1` | Numeric columns aligned on the decimal point. After 05, which pins the widths the tab positions depend on |
| 09 | chapter numbers | `..\word\Set-NISTChapterNumbering.ps1` | Each chapter's own number back on Heading 1, read from its original (the move puts every chapter on the template's list, which starts at 1) |
| 10 | captions | `caption_check.py` | A label-less Caption_Figure paragraph after a figure caption -> Caption_Figure_2. Reports empty and placeholder captions ("xx", "text"); front matter and appendix letter numbers ("P–1", "B–1") stay typed. Stops on a typed chapter number |
| 11 | update fields | `update_fields.py` | Word F9. First: caption bookmarks reset to label and number, fields in tracked deletions locked, REF guard. After: split tracked insertions rejoined, locks removed |
| 12 | cross-refs | `..\word\Set-TableCrossReferences.ps1` (Table, Figure), `section_xrefs.py` | Typed "Table 3–1", "Figure 3–2", "Section 3.3" -> live references, same chapter only. Mentions crossing a tracked-change boundary, in other chapters, or with no matching heading stay typed and are reported |
| 13 | update fields | `update_fields.py` | F9 again, so the new references show current numbers |
| 14 | three columns *(optional)* | `..\word\Set-NISTThreeColumnTableWidths.ps1` | Every 3-column table at `three_column_widths_inches`. Off: suits survey tables only |
| 15 | accent bar *(optional)* | `..\word\Set-TableAccentBar.ps1` | Colored bar down each table's left edge, color by caption. Off |

What counts as a table: a **callout** is a single-cell table or one on the
Sidebar Table style; the **chapter banner** is the table directly before the
chapter's Heading 1 (or one holding Header / Headerch paragraphs); everything
else is a **data table**, whatever style it arrived with.

## The check

`compare_content.py` runs after every step that touches chapters, against the
originals through step 10 and against the step before from 11 on (once fields
are updated, their results are current, not original). It compares, per
chapter: the text with every tracked change accepted and with every change
rejected, word for word; what each author inserted, deleted and moved, in
characters; comments, footnotes, endnotes, images, equations, links (in either
form: the move rewrites table-of-contents links as HYPERLINK fields), fields; and each
table's shape (skipped from step 06 on, which merges note rows).

Allowances, each limited to the step that needs it:
- **field update (11, 13):** what live fields display may change, and the
  number of fields; a deleted field's result may not (it must stay deleted).
- **cross-refs (12):** the number of fields may change, a dash between digits
  may change form (a linked "Figure 4-14" shows the caption's "Figure 4–14"),
  and field markers inside a tracked insertion are not counted as inserted text.
- **after remove toc (02 on):** the originals' tables of contents are set
  aside, except in `keep_toc` files; everything else must still match.
- **always:** an empty insertion Word drops on save, and invisible
  text-direction marks Word rewrites.

A failed check stops the run and names the file and the difference.

It can also be run by hand: `python compare_content.py <originals> <folder>`.
