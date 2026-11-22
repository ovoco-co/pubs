# Restoring Zotero links after the template pass

The NIST template pipeline keeps every citation's text but loses the Zotero fields around
it, and it drops the document's Zotero preferences. After that, the Zotero Word plugin no
longer recognizes the chapter: it can't refresh it, edit a citation, or rebuild the
reference list. When an author still works in Zotero, the links have to come back.

The author's original still has everything. Each citation field carries its full item data
and the group library URI (`zotero.org/groups/<id>/items/<key>`), so copying the field
restores a working link without anyone picking items from the library again.

## Steps

1. **Keep the originals.** Download the author's last Zotero-linked versions into their own
   folder, for example `HM-261122`. They are read, never written.
2. **Close the masters in Word.** The script skips any chapter with a `~$` lock file.
3. **Dry run** (writes nothing):

   ```powershell
   .venv\Scripts\python word\restore_zotero_fields.py <originals> <masters>
   ```

   For each chapter it reports what it would restore and what it would leave alone.
4. **Read the exceptions** (below), then run again with `--write`. Use `--only 02,06` to
   limit the run to some chapters.
5. **Refresh in Word.** With Zotero running, open each chapter and click
   *Zotero > Refresh*. This is the real test, and it is a manual step.

## What the script does

| Part | Original | Master |
|---|---|---|
| Citation | `ADDIN ZOTERO_ITEM CSL_CITATION {...}` field | The same text, now wrapped in that field. The instruction is copied unchanged. |
| Reference list | `ADDIN ZOTERO_BIBL` field over the entries | The master's existing entries, first to last, wrapped in that field. The NIST styling stays. |
| Preferences | `ZOTERO_PREF_1`, `_2`... custom properties (style, field type) | Copied across, replacing any already there. |

**Matching.** Citations are matched in document order. Each search starts after the previous
match, so a repeated citation such as "(GAO, 2019a)" lands on the right occurrence. A match
counts only if the 60 characters before it (and the 30 after) are similar enough to the
original's. The dry run lists any match accepted with a context score below 0.9.

**What changes in the master.** Only field markers are inserted. The text, styles and tracked
changes stay as they are. A run that a citation starts or ends inside is split at that
point, with its formatting kept.

## Exceptions it reports and leaves alone

- **Deleted in the master as a tracked change.** An editor cut or rewrote the sentence. The
  old field is still intact inside the deletion, so rejecting the change brings it back.
  Nothing to do.
- **Not found.** The sentence was rewritten without tracking. Re-add the citation by hand
  with *Zotero > Add/Edit Citation*.
- **Crosses a tracked change or hyperlink.** The citation's text is split between an
  insertion and plain text, or between a hyperlink and plain text, so no single field can
  wrap it. Accept the change, then re-run the script or add the citation by hand.
- **Reference entry edited in the master.** *Refresh* rebuilds the list from the library and
  overwrites the edit. Make the same correction to the Zotero item, or the fix is lost.

## What to expect on Refresh

- **Hand-edited citations.** Authors sometimes edit a citation's text inside the field, for
  example dropping the parentheses to fit a sentence. Zotero asks whether to keep each one.
  Answer **Yes** (keep the modification): the text is the author's, not the script's. The
  dry run's matching counts these as unchanged. Compare the counts with the original to
  confirm none were introduced.
- **Reference list styling.** The regenerated list uses Zotero's *Bibliography* paragraph
  style, not `References`. Set *Bibliography* to match `References` (or base it on
  `References`) so the list doesn't change appearance.
- **Different group library.** A citation from a different group library refreshes only for
  someone who belongs to that group.

## Checking without Zotero

Compare, per chapter, the number of `ZOTERO_ITEM` fields outside tracked deletions with the
original's count. Also compare the number of hand-edited citations, where the result text
differs from `plainCitation`. Both should match, except for citations the editors deleted.

## NWIRP 2-1 (2026-11-22)

Originals `HM-261122`, masters `HM-261119`:

| Chapter | Restored | Deleted by editors | Reference list | Entries edited in master |
|---|---|---|---|---|
| 1 | 12 of 13 | DOE 2018 (sentence rewritten) | wrapped, 14/14 identical | none |
| 2 | 100 of 101 | FHWA 2016 (sentence cut) | wrapped, 44/45 identical | PRASA 2019 |
| 3 | 9 of 9 | none | none in original | none |
| 6 | 64 of 64 | none | wrapped, 44/46 identical | Lei 2016, Lewis 2019 |
| 7 | 12 of 12 | none | none in original | none |

Chapters 0 and 4 never used Zotero. Chapter 5 had preferences but no citations; its
preferences were copied.

## Linking chapters that never used Zotero

`link_zotero_citations.py` handles the other case: citations and a reference list typed by
hand, or sources that sit in footnotes. Every change is a tracked revision under `--author`,
so the editor or author can review it and reject it.

1. **Get every source into the library.** Search the group library through Zotero's local API
   (Zotero > Settings > Advanced > *Allow other applications on this computer to communicate
   with Zotero*). Add the missing items, tagged so they can be reviewed as a set. The
   connector's `/connector/saveItems` endpoint creates them in the selected library.
2. **Write a map** from the citation text to the item key:
   - `cites`: parenthetical citations, e.g. `"FCC 2017-2018"`.
   - `narrative`: `"Cordova et al.|2021"`.
   - `replace`: one-off spans, such as a bare URL or a statute citation.
   - `edits`: plain-text edits that a replacement depends on.
   - `footnotes`: footnotes that are nothing but a source.
   - `bibliography`: replace the list, or append a new one.

   The script's docstring gives the format.
3. **Dry run, then `--write`.** Citations the map doesn't cover are listed as "left as typed".
   Settle them in the map, or flag them with a comment for the author.
4. **Refresh in Word.** The citation text the script inserts is provisional. Refresh rewrites
   it in the document's style.

The NWIRP 2-1 maps (Ch4 and the Preface) and the list of items added for them are in
`HM-261122`.

## House citation style

`zotero/nist-author-date.csl` is the house style:
- **In the text:** Chicago author-date with no comma, e.g. (FCC 2020), (Maitland and Peha 2018),
  (Cordova et al. 2021), (FEMA 2018a, 2018b).
- **Reference list:** NIST Technical Series format, e.g. "Taylor BN (2011) Title. *Journal*
  116(6):797. https://doi.org/..."

To install it, open Zotero > Settings > Cite > Styles and click "+", or POST the file to
`/connector/installStyle`. To switch a chapter, open Zotero > Document Preferences in Word.

**Organization authors.** The style prints the author field as entered, so cite organizations
by their short form. Put the short form in the author field (FEMA) and the full name in
Publisher.
