"""Update every field in each chapter with Word: the equivalent of Ctrl+A, F9.

Usage: python update_fields.py <folder>

Edits the files in place. Caption numbers (STYLEREF, SEQ), cross-references
(REF) and any table of contents recompute from the document as it now stands.

BEFORE Word runs (XML), every bookmark a REF field points at is made safe,
because a REF displays whatever its bookmark contains:
  LABEL AND NUMBER  a bookmark in a caption is reset to exactly the label and
          number ("Figure 3-6"), what Word's Insert Cross-reference writes for
          "Only label and number". Authors' bookmarks that began before the
          figure pasted the image into every reference on update; ones that
          ended after the period showed "Figure 3-6.  " mid-sentence.
  GUARD   then every REF must point at a bookmark that exists, stays inside
          one paragraph and holds no image. (A REF whose bookmark ran thousands
          of paragraphs downstream once pasted ~1.75 MB into NCSTAR Ch04.) Any
          failure stops the run before Word opens anything.
  REPORT  a reviewer's tracked edit made INSIDE a field's displayed result
          (someone retyping a reference number by hand) is listed: the update
          recomputes that text, as any F9 would.

  LOCK    a field inside a tracked deletion is locked (w:fldLock) for the
          update and unlocked after. Word would otherwise rewrite its deleted
          result as live text, bringing a reference a reviewer deleted back
          into the document.

Word then updates the fields with track changes OFF, so the new numbers are
not recorded as revisions: body fields one at a time, then headers and
footers, then tables of contents and lists of tables/figures (see update_all).

AFTER Word runs (XML): when a field sits inside a tracked insertion, Word
moves its new displayed text out of the insertion and splits it in two. The
two halves are joined back around the result, same author and date, so
rejecting that insertion still removes the whole reference.

Needs Word, and Word must not already be running.
"""
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import win32com.client as win32
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
CAPTION_STYLES = {'CaptionFigure', 'CaptionTable', 'Caption'}


def local(el):
    return etree.QName(el).localname


def ref_targets(root):
    instr = [t.text or '' for t in root.iter(W + 'instrText')] + \
            [f.get(W + 'instr') or '' for f in root.iter(W + 'fldSimple')]
    return {m.group(1) for i in instr for m in [re.match(r'\s*REF\s+(\S+)', i)] if m}


def para_of(el):
    return next(el.iterancestors(W + 'p'), None)


def is_caption(p):
    ps = p.find(f'{W}pPr/{W}pStyle') if p is not None else None
    return ps is not None and ps.get(W + 'val') in CAPTION_STYLES


def label_and_number(p):
    """(first run of the label, element the number ends with) in a caption."""
    label = None
    for r in p.iter(W + 'r'):
        if re.match(r'\s*(Figure|Table)\b', ''.join(x.text or '' for x in r.iter(W + 't'))):
            label = r
            break
    if label is None:
        return None, None
    for fs in p.iter(W + 'fldSimple'):
        if re.search(r'\bSEQ\b', fs.get(W + 'instr') or ''):
            return label, fs
    depth, seq = 0, False
    for el in p.iter(W + 'fldChar', W + 'instrText'):
        if el.tag == W + 'instrText' and re.search(r'\bSEQ\b', el.text or ''):
            seq = True
        elif el.tag == W + 'fldChar' and el.get(W + 'fldCharType') == 'end' and seq:
            return label, el.getparent()
    return label, None


def span_problem(start, end):
    """Why the bookmark from start to end is unsafe for a REF to paste."""
    if para_of(start) is not para_of(end) or para_of(start) is None:
        return 'spans a paragraph break'
    p = para_of(start)
    order = list(p.iter())
    if order.index(start) > order.index(end):
        return 'ends before it starts'
    if any(x.tag in (W + 'drawing', W + 'pict') for x in order[order.index(start):order.index(end)]):
        return 'holds an image'
    return None


def prepare(root):
    """Reset caption bookmarks, then guard. Returns (reset, problems)."""
    starts = {b.get(W + 'name'): b for b in root.iter(W + 'bookmarkStart')}
    ends = {b.get(W + 'id'): b for b in root.iter(W + 'bookmarkEnd')}
    reset, problems = 0, []
    for name in sorted(ref_targets(root)):
        start = starts.get(name)
        end = ends.get(start.get(W + 'id')) if start is not None else None
        if start is None or end is None:
            problems.append(f'REF {name}: bookmark missing')
            continue
        cap = next((p for p in (para_of(end), para_of(start)) if is_caption(p)), None)
        if cap is not None:
            label, number = label_and_number(cap)
            if label is not None and number is not None:
                if label.getprevious() is not start or number.getnext() is not end:
                    start.getparent().remove(start)
                    label.addprevious(start)
                    end.getparent().remove(end)
                    number.addnext(end)
                    reset += 1
        problem = span_problem(start, end)
        if problem:
            problems.append(f'REF {name}: {problem}')
    return reset, problems


def hand_edited_results(root):
    """Field results that hold a reviewer's tracked edit."""
    found, stack = [], []
    for el in root.iter(W + 'fldChar', W + 'instrText', W + 'ins', W + 'del'):
        if el.tag == W + 'fldChar':
            kind = el.get(W + 'fldCharType')
            if kind == 'begin':
                stack.append({'instr': '', 'result': False})
            elif kind == 'separate' and stack:
                stack[-1]['result'] = True
            elif kind == 'end' and stack:
                stack.pop()
        elif el.tag == W + 'instrText' and stack and not stack[-1]['result']:
            stack[-1]['instr'] += el.text or ''
        elif (el.tag in (W + 'ins', W + 'del') and stack and stack[-1]['result']
              and not any(True for _ in el.iter(W + 'fldChar'))):
            found.append(f'{local(el)} by {el.get(W + "author")} in "{stack[-1]["instr"].strip()[:30]}"')
    return found


def deleted_fields(root):
    """Field starts (complex fldChar begin, or fldSimple) inside a tracked
    deletion or move-from."""
    out = []
    for el in root.iter(W + 'fldChar', W + 'fldSimple'):
        if el.tag == W + 'fldChar' and el.get(W + 'fldCharType') != 'begin':
            continue
        if any(a.tag in (W + 'del', W + 'moveFrom') for a in el.iterancestors()):
            out.append(el)
    return out


def lock_deleted(root):
    """Lock deleted fields that are not locked already; returns how many."""
    n = 0
    for el in deleted_fields(root):
        if el.get(W + 'fldLock') not in ('1', 'true', 'on'):
            el.set(W + 'fldLock', '1')
            n += 1
    return n


def unlock_deleted(root):
    n = 0
    for el in deleted_fields(root):
        if el.get(W + 'fldLock') is not None:
            del el.attrib[W + 'fldLock']
            n += 1
    return n


def rejoin_insertions(root):
    """Join <ins A>…separate</ins> result-runs <ins B>end…</ins> back into one
    insertion when A and B have the same author and date."""
    joined = 0
    for a in list(root.iter(W + 'ins')):
        if a.getparent() is None or local(a.getparent()) != 'p':
            continue
        marks = [f.get(W + 'fldCharType') for f in a.iter(W + 'fldChar')]
        if not marks or marks[-1] != 'separate':
            continue
        between, nxt = [], a.getnext()
        while nxt is not None and nxt.tag == W + 'r' and not list(nxt.iter(W + 'fldChar')):
            between.append(nxt)
            nxt = nxt.getnext()
        if (nxt is None or nxt.tag != W + 'ins' or not between
                or nxt.get(W + 'author') != a.get(W + 'author') or nxt.get(W + 'date') != a.get(W + 'date')):
            continue
        first = [f.get(W + 'fldCharType') for f in nxt.iter(W + 'fldChar')]
        if not first or first[0] != 'end':
            continue
        for r in between + list(nxt):
            a.append(r)
        nxt.getparent().remove(nxt)
        joined += 1
    return joined


WD_FIELD_TOC = 13   # TOC fields: tables of contents and lists of tables/figures


def update_all(doc):
    """Ctrl+A, F9, in an order Word finishes. One doc.Fields.Update() over a
    chapter whose contents and lists of tables/figures hold ~640 PAGEREF and
    HYPERLINK entries ran past 5 minutes without returning (NWIRP 2-02 Ch02);
    this takes under a minute. Body fields first, one at a time in document
    order (so SEQ numbers count up correctly), skipping the entries inside a
    TOC field; then headers and footers; then each TOC field, which rebuilds
    its entries from the updated captions and headings."""
    fields = list(doc.Fields)
    tocs = [x for x in fields if x.Type == WD_FIELD_TOC]
    spans = [(x.Code.Start, x.Result.End) for x in tocs]
    body = [x for x in fields if x.Type != WD_FIELD_TOC and not any(s <= x.Code.Start <= e for s, e in spans)]
    for x in body:
        x.Update()
    for section in doc.Sections:
        for hf in list(section.Headers) + list(section.Footers):
            hf.Range.Fields.Update()
    for x in tocs:
        x.Update()
    return len(fields)


def rewrite(path, fn):
    zin = zipfile.ZipFile(path)
    root = etree.fromstring(zin.read('word/document.xml'))
    result = fn(root)
    new = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    tmp = path.with_suffix('.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, new if info.filename == 'word/document.xml' else zin.read(info.filename))
    zin.close()
    tmp.replace(path)
    return result


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    folder = Path(sys.argv[1])
    files = sorted(p for p in folder.glob('*.docx') if not p.name.startswith('~$'))
    bad, locked = {}, {}
    for f in files:
        root = etree.fromstring(zipfile.ZipFile(f).read('word/document.xml'))
        for line in hand_edited_results(root):
            print(f'{f.name}: reviewer edit inside a field result, recomputed: {line}')
        already = sum(1 for el in deleted_fields(root) if el.get(W + 'fldLock') is not None)
        if already:
            bad[f.name] = [f'{already} field(s) in tracked deletions are locked already; '
                           'cannot tell those apart from the ones this step locks']
            continue
        locked[f.name] = rewrite(f, lock_deleted)
        reset, problems = rewrite(f, prepare)
        if reset:
            print(f'{f.name}: {reset} caption bookmark(s) reset to label and number')
        if problems:
            bad[f.name] = problems
    if bad:
        sys.exit('Field update NOT run, cross-reference problems:\n' +
                 '\n'.join(f'  {n}: ' + '; '.join(p) for n, p in bad.items()))
    if 'WINWORD.EXE' in subprocess.run(['tasklist', '/FI', 'IMAGENAME eq WINWORD.EXE'],
                                       capture_output=True, text=True).stdout.upper():
        sys.exit('Word is running. Close it first: this script would attach to the open copy.')
    word = win32.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for f in files:
            doc = word.Documents.Open(str(f.resolve()), False, False, False)
            try:
                doc.TrackRevisions = False
                n = update_all(doc)
                doc.TrackRevisions = True
                doc.Save()
                print(f'{f.name}: {n} fields updated')
            finally:
                doc.Close(0)
    finally:
        word.Quit(0)
    for f in files:
        joined = rewrite(f, rejoin_insertions)
        if joined:
            print(f'{f.name}: {joined} tracked insertion(s) rejoined around an updated reference')
        if locked.get(f.name):
            unlocked = rewrite(f, unlock_deleted)
            print(f'{f.name}: {unlocked} field(s) in tracked deletions held back from the update')


if __name__ == '__main__':
    main()
