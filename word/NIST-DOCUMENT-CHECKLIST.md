# NIST chapter checklist

What `Test-NISTDocument.ps1` checks, what the canonical value is, and where that
value comes from. Every entry below exists because the defect was found in a real
NWIRP 2-3 or NCSTAR 4-3 chapter.

```powershell
.\Test-NISTDocument.ps1 -Folder "C:\Users\Artist\Documents\HM-261102"
.\Test-NISTDocument.ps1 -File   "...\NWIRP_2-3_Chapter-04_Quantitative_Results.docx"
```

FAIL means a reader would see it. WARN means it is wrong but invisible, or it is a
judgement call.

## Authority

The base architecture is `C:\Users\Artist\Documents\nist-scripts` (the originals).
`publishing/word/` holds later evolution on top. Where they disagree, the originals
define the base and the newer value has to be justified. Two known divergences:

| | originals | toolkit | in the documents |
|---|---|---|---|
| NIST-Data cell margin, left | 29 | 58 | 58 - deliberate, the 2.7/0.6/0.6 widths are measured from it |
| NIST-Data `firstCol` fill | D9D9D9 | F2F2F2 | unresolved |

The theme-reference rule in `EDITING-OOXML.md` is **superseded**: the theme approach
was abandoned and colors are hardcoded. `Audit-WordTheme.ps1` will report thousands
of "explicit colors bypassing the theme" - that is expected, not a defect.

## Canonical values

| thing | value | source |
|---|---|---|
| `xxbase` (all body) | Source Serif Pro, sz 20 = 10 pt | `Set-NISTStyleArchitecture.ps1` |
| `xxheader` (all headings) | Source Sans Pro bold, sz 24 = 12 pt | same |
| `xxtable` (all table text) | Source Sans Pro, sz 18 = 9 pt | same |
| `Header` / `Footer` | Source Sans Pro Light, sz 18, `00257A` | same |
| `Caption` | Source Sans Pro **Semibold**, sz 20 = **10 pt**, no bold | same |
| page | Letter, 1 in all round, header/footer **0.3 in** (432 twips) | `Set-NISTPageLayout.ps1` |
| NIST-Data rules | `single sz=4 color=808080` + `themeColor=background1 themeShade=80` | `$NIST_DATA_CANONICAL` |
| NIST-Data cell margins | 14 / 58 / 14 / 29 | toolkit |
| accent bar | 4 pt = **`sz="32"`**, base strength, first cell of every row | `nist-color-palette.md` |
| banner tint | 80 % tint: schools F6DFCC, hospitals CCECE3, neither CCD3E4 | same |
| navy | **`00257A`** | theme accent1 |
| caption format | `Figure` + NBSP + chapter + en dash + SEQ + `.` + **two spaces** + title | tools/CLAUDE.md |
| base style family | `xxbase`, `xxheader`, `xxtable` - **three, not four** | originals |

`Source Sans Pro Semibold` - lowercase **b**. `SemiBold` does not resolve and falls
back silently. It broke PDF output and was fixed in the script; the Heading 6 writer
was missed and kept reintroducing it.

## Traps that cost real time

- **A bare `REF bookmark \h` pastes the entire bookmark span**, not a number. Six
  caption bookmarks in Chapter 4 never closed at their caption, so each spanned
  ~1.75 MB; one field update took the chapter from 6,257 paragraphs to 22,958.
  `\r` makes it insert a number instead.
- **Appendix tables are keyed by letter** (`Table C-1`, `E-4`). Converting those to
  a chapter `SEQ` field renumbers them 4-53 … 4-63. Check what SEQ would produce
  before converting any typed caption number.
- **A List of Tables looks like a run of captions on BodyText.** The giveaway is a
  trailing page number: `Table 3-1. Data Collection Instruments7`.
- **An empty paragraph inside a table cell is structural.** NCSTAR's 1,482 "extra
  paragraph returns" were 1,480 table cells and 2 real ones.
- **A paragraph whose text sits inside `<w:del>` looks empty** to a `<w:t>` scan.
  Deleting it destroys a pending tracked change.
- **Google Docs `goog_rdk` content controls nest paragraphs inside `w:sdt`**, so
  `findall` on a `w:tc` finds nothing. Use `.iter()`, or unwrap them first.
- **Banded rows and tinted columns cannot coexist** - the stripe paints only the
  untinted columns and you get a checkerboard.
- **Word COM: create the Application inside the `try`**, set `Visible`/`DisplayAlerts`
  inside it too, and wrap `Quit()` in its own `catch`. An orphaned headless WINWORD
  holds every file it opened locked until it is killed.

## The checks

**INTEGRITY** - runaway bookmark via a bare REF · bookmarks pair up (count,
duplicate starts and ends, orphans) · field characters balance · no Google Docs
export residue

**STYLES** - every used style defined · no dangling basedOn/link/next · xx* base
family (three) · docDefaults font · no non-NIST fonts applied directly · palette
colors canonical · no `xxgraphic` · no orphan character styles

**LAYOUT** - no text frames · page setup · empty paragraphs carry a style · figures
on the `graphic` style

**NUMBERING** - every numId resolves · heading numbering chain

**CAPTIONS** - numbers are fields · on the matching style · size is 10 pt and
style-owned · not italic · figure caption below its figure · table caption above
its table

**TABLES** - on NIST-Data · tblLook 0420 · header row repeats · accent bar · footnote
row unshaded with a zeroed cnfStyle and nil side/bottom rules · fixed width

**FIELDS** - every REF target has a bookmark · no empty SEQ husks · no
citation-manager linking

**REFERENCES** - entries on the References style · journal and book titles italic

**CONTENT** - alt text on every image (accepted document) · no review markers ·
`dc:title` set · no table of contents in the chapter · fonts de-embedded

**HEADER/FOOTER** - paragraphs styled · running head names the chapter

### Reading a running head

Two traps, both of which made a correct document look broken.

Find the header parts through `word/_rels/document.xml.rels`, not by filename.
Word's own naming is `header1.xml`, but a template may name them anything:
NCSTAR 4-3 uses `nistHeaderFirst.xml`, `nistHeaderEven.xml` and
`nistHeaderDefault.xml`, so a `header\d+\.xml` pattern finds no parts at all and
reports the header as empty. (`Get-WordParts` has to load `.rels` as well as
`.xml` for that lookup to see anything.)

Then judge the FIELD, not its cached result. The correct running head is
`STYLEREF "Heading 1"`, which Word recomputes per page from the heading above
it - odd pages take the heading text, even pages take `\n` for the number, which
the `Chapter %1` lvlText renders as "Chapter 4". The literal between
`fldCharType="separate"` and `"end"` is only the last value Word cached, so a
perfectly good header can read "Chapter Title" (the NCSTAR template default),
or name a different chapter, and still render correctly on open. Only a header
with NO STYLEREF is worth testing against its own text.

An empty first-page header is the convention, not a defect: no running head on
the opening page of a chapter. Test the default and even headers.

## Bringing a raw author set into the template

The order matters, and NWIRP 2-1 established it.

1. **Remap document.xml FIRST.** `Set-NISTStyleArchitecture.ps1` never edits
   document.xml, and it skips a rebase whose base style is absent (it warns:
   "Rebase SKIPPED"). Map the author styles it does not define - `Bibliography`
   to `References`, `AppendixText` to `BodyText` (it is defined as
   `basedOn BodyText` with no overrides), `NormalWeb` to `BodyText` - and give
   every unstyled body paragraph `BodyText`. 77 percent of 2-1's paragraphs had
   no `pStyle` at all; of 2,356, only 859 were real body text, 1,037 were table
   cells and 168 were empty spacers.
2. **Import the missing styles.** No script installs a style set. The
   architecture CREATES only `xxbase`/`xxheader`/`xxtable`, Header/Footer,
   `References`, `Title`, `MTDisplayEquation`, `NIST-Data`, `SidebarTable` and
   the List ladders; it only EDITS `BodyText`, `CaptionTable`, `CaptionFigure`,
   `graphic`, `TableText` and no-ops when they are absent. Donor is
   `projects\pubs\word\NIST-reference.docx`. Resolve basedOn/link/next
   transitively or the import leaves a new dangling reference, and never import
   `xxgraphic`. Deletion in the architecture is an explicit blacklist, not a
   whitelist, so unrecognized author styles survive - but every TABLE style
   outside `TableNormal/TableGrid/HM-Cover/HM-Data/NIST-Data/SidebarTable` is
   deleted, so repoint `PlainTable2` and friends first.
3. `Build-NISTDocuments.ps1` (all 7 steps), then `Set-TableParagraphStyle.ps1`
   for the cells, then `Set-DocTitle.ps1` before
   `Set-NISTHeaderFooterContent.ps1` (the footer reads `dc:title`). Use
   `-SkipFieldUpdate`: a Word field update is what corrupted NCSTAR Ch04.

### Citation managers have TWO forms

Zotero writes fields - `fldChar begin` / `instrText ADDIN CSL_CITATION` /
`separate` / result / `end`. Mendeley and newer Zotero write SDT content
controls - `<w:sdt><w:sdtPr><w:tag w:val="MENDELEY_CITATION_v3_<base64>"/>`.
A fldChar-only walker finds nothing in a Mendeley chapter: 2-1 Ch05 had 127
controls and zero fields. Unlinking a field means keeping the runs between
`separate` and `end`; unlinking a control means unwrapping it and keeping
`sdtContent`. Both need a depth-aware scan.

The auto-generated bibliography is its own control, tagged
`MENDELEY_BIBLIOGRAPHY`. While it stays wrapped its entries cannot be restyled,
which is why Ch05 read "0 of 87 on References" - unwrap it, then the entries
take `References`.

A citation whose `instrText` is a `delInstrText` inside `w:del` must be LEFT
ALONE. That is a tracked deletion, and rewriting it falsifies the revision.

### Measuring a new set against the certified one

Do not re-decide settled questions - measure the new set against the certified
one and reproduce it. 2-1 vs 2-3 found: banner table missing on 7 of 8 chapters
(the certified set has exactly one per chapter, page-anchored, first in
`w:body`); 92 figures not on `graphic`; 0 of 31 content tables on `NIST-Data`
(the certified set puts every table except the banner on it); 41
`ListParagraph`; 60 orphan character styles. It also showed two of the rules
need NO work here - 2-1 carries no accent bars, which is correct because the bar
codes schools against hospitals and 2-1 is infrastructure, and its reference
italics are already right because the citation manager generated them.

For orphan character styles, build the keep-list from the certified set rather
than judging each one: 2-3 keeps `Hyperlink`, `CommentReference` and
`FootnoteReference` and drops Word's other latent built-ins (`Emphasis`,
`IntenseEmphasis`, `LineNumber`, `PlaceholderText`, `UnresolvedMention`...).
Deleting a style must also remove the `w:link` on its paragraph partner.

The banner is identified by PAGE anchoring (`vertAnchor="page"` AND
`horzAnchor="page"`), never by the presence of `w:tblpPr`: authors float content
tables with `vertAnchor="text"`, and a bare `tblpPr` test skipped two real data
tables in Ch03.

`ListParagraph` items move to `List Bullet`/`List Number` by the numFmt of the
numbering they point at, and the paragraph's direct `numPr` comes OFF - the
certified set's 30 ListBullet paragraphs carry none, because the style's own
numId supplies the glyph.

## Inserting a missing caption

Build it to the certified shape, with BOTH numbers as fields:

```
CaptionFigure
  "Figure" + NBSP
  field  STYLEREF 1 \s                -> chapter number of the Heading 1 above
  en dash (U+2013)
  field  SEQ Figure \* ARABIC \s 1    -> figure number, restarting each chapter
  ". " then " <title>"                -> period then TWO spaces
```

A typed number goes stale the moment anything moves, because Word only advances
a SEQ counter where it evaluates a SEQ field. STYLEREF rather than a literal
chapter means the caption reads the chapter it lands under. Write NO `_Ref`
bookmark: an unreferenced one is what grew into the runaway span that corrupted
NCSTAR Ch04, and Word creates one when someone inserts a cross-reference.

Figure captions go BELOW the figure, table captions above. Group consecutive
image paragraphs into one figure - a blank paragraph between panels does not
break the group, a text paragraph does - or a multi-panel figure gets a caption
per panel.

Make the placeholder say what it is for (`[PLACEHOLDER - figure title]`) and
make sure the review-marker check actually matches it. The original pattern
required the brackets to hold nothing but the keyword, so 39 inserted
placeholders were reported as "none".

## Direct formatting that fights the style

A direct `w:sz` always beats the paragraph style, so a correct 10pt BodyText
still renders at 11 or 12pt. 2-1 Ch00 carried a hard size on 100 percent of its
body runs; Ch04 had genuine body paragraphs at 13 and 14pt alongside bold 18pt
figure labels and a 21pt title block sitting on BodyText.

A direct 10pt is redundant, not wrong - it matches what xxbase gives BodyText,
and the certified set carries plenty. Only a different size fights.

`Remove-DirectRunSize.ps1` strips only `w:sz`/`w:szCs` and leaves bold, italic
and color alone. Note it also removes `w:szCs`, the COMPLEX-SCRIPT size, which
has no effect on Latin text - that is why it reported removing 1,378 elements
from Ch05 where a scan for `w:sz` found none. Count the two separately before
concluding anything about what changed.

Compute point sizes carefully: `w:val="21"` is 10.5pt, and integer division
renders it as "10pt", which hides it among the legitimate redundant 10pt runs.

Highlighter (271 runs across 2-1, including 120 white and 15 green in Ch04) and
direct spacing/indent on BodyText (Ch05 85 of 89 paragraphs, Ch06 108 of 110)
are REPORTED, never stripped in a blanket pass. Highlight is emphasis and
spacing may be a deliberate draft setting; both are the author's to remove.

## Direct run formatting: what comes off and what stays

Measure the certified set before stripping anything. 2-3 carries just as much
direct run formatting as a raw set does - rFonts 9,933 vs 11,320, color 4,549 vs
3,525, 9,594 rFonts holding nothing but `w:cs`/`w:eastAsia`. A blanket strip
makes a new set DIVERGE from the standard rather than meet it, which is the
same reason `Reapply-Styles.ps1` is excluded from the pipeline.

The one dimension that was genuinely out of line in 2-1 was the typeface
actually rendered: **1,026 off-palette direct `ascii`/`hAnsi` overrides against
6 in the certified set.**

REMOVE - Word defaults and paste residue carrying ordinary prose: Calibri (408),
Times New Roman (364), Aptos Narrow (82), Arial (56), Georgia, Aptos, Segoe UI,
Roboto. These were sitting on text like "Hurricane Maria made landfall" and
"Recovery of Business and Supply Chain".

KEEP - the glyph depends on the font: **Segoe UI Symbol** carries the real
U+2713 checkmark and U+2731/U+273B asterisks (20 of them in Ch05 alone);
Cambria Math carries the equations (1,426 runs); Symbol and Wingdings carry
bullet glyphs. Keep every Source face.

KEEP - `w:cs` and `w:eastAsia` only affect complex-script runs and have no
visible effect on Latin text. Including them turned a 1,026-override job into a
16,119-element one on the first attempt.

NEVER TOUCH: `w:vertAlign` (66 in 2-1 - footnote markers, ordinals "th"/"st",
and math subscripts), `w:i`, `w:b`, `w:u`, `w:strike`, `w:highlight`, `w:shd`,
and `w:sym`, which names its own font per character.

Direct SIZES are the exception worth removing, because a direct `w:sz` always
beats the style. A direct size equal to what the style gives (10pt on BodyText,
10pt on a caption) is redundant rather than wrong.

### Verify with whitespace-tolerant patterns

`<w:i/>` and `<w:i />` are the same element, and a set can hold both forms -
some of these files were re-serialized with the spaced form mid-pass. Searching
for the literal `'<w:i/>'` reported 60 lost italic reference titles in Ch02 that
were sitting there untouched. Use `<w:i(?:\s[^>]*?)?\s*/>`. This trap has now
produced a wrong conclusion three times: a missed style match, a stray `/>` text
node, and a false formatting-loss alarm in the verification code itself.

## Still open

- Reference italics. The NIST template italicises the standalone work (book,
  journal, proceedings, standard) but not the piece inside it. The chapters are APA
  and `NIST-Template-Comparison.md` records the reference format as an unresolved
  question with NIST. Ch1 223 entries, Ch2 119, Ch3 16.
- Auto-width tables: NCSTAR 22/45, Ch4 16/64, Ch5 3/5, Ch6 1/4.
- Orphan character styles: ~29 per chapter, mostly pandoc `*Tok` syntax-highlight
  leftovers. Deleting a `*Char` style needs its paragraph partner's `w:link`
  removed in the same pass.
- Running heads are built two ways across the set. NCSTAR 4-3 uses a STYLEREF
  field; the NWIRP chapters carry a literal ("Findings  Chapter 6"). Both name
  the right chapter, so neither is wrong, but a literal does not follow a
  heading that gets renamed.
