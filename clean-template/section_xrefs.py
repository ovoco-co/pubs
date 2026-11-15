r"""Turn typed "Section 3.3" references into live cross-references, with Word.

Usage: python section_xrefs.py <folder>

Edits the files in place. For each chapter Word reports the number it shows
on every numbered heading (Heading 2 "3.3.", Heading 3 "3.3.1.", ...). Each
typed "Section N.N" whose number matches a heading IN THIS CHAPTER has its
number replaced by { REF _Ref... \w \h }, the field Word's own Insert
Cross-reference writes for "Paragraph number (full context)": it shows "3.3"
(no trailing period) and follows the heading if it is renumbered. "Section "
stays as text.

Left as typed text, and reported:
  - a number whose chapter is not this one (one Word file cannot point into
    another);
  - a number that matches no heading in this chapter (a stale reference);
  - a mention inside a tracked change, or already inside a field.
A chapter whose Heading 1 is unnumbered (front matter) is skipped.

Track changes is OFF while the fields go in (they are a conversion of text
that is already there, not an author's edit) and back ON before saving.
Needs Word, and Word must not already be running.
"""
import re
import subprocess
import sys
from pathlib import Path

import win32com.client as win32

PATTERN = 'Section [0-9]@.[0-9.]@'      # Word wildcards: "Section " then a dotted number


def heading_map(doc):
    """Chapter number, and {"3.3": paragraph} for every numbered heading 2-9."""
    chapter, heads = None, {}
    for para in doc.Paragraphs:
        style = para.Style.NameLocal.lower()
        if not style.startswith('heading'):
            continue
        num = para.Range.ListFormat.ListString
        if not num:
            continue
        if style == 'heading 1':
            m = re.search(r'\d+', num)
            chapter = chapter or (m.group(0) if m else None)
        else:
            heads.setdefault(num.rstrip('.'), para)
    return chapter, heads


def bookmark(doc, para, n):
    name = f'_Ref{900000000 + n}'
    rng = para.Range.Duplicate
    rng.MoveEnd(1, -1)                   # the heading text, not its paragraph mark
    doc.Bookmarks.Add(name, rng)
    return name


def process(doc):
    chapter, heads = heading_map(doc)
    report = {'linked': [], 'other chapter': [], 'no such heading': [], 'in a tracked change or field': []}
    if not chapter:
        return None, report
    marks = {}
    rng = doc.Content
    find = rng.Find
    find.ClearFormatting()
    find.Text = PATTERN
    find.MatchWildcards = True
    find.Forward = True
    find.Wrap = 0
    while find.Execute():
        hit = rng.Text
        num = hit[len('Section '):].rstrip('.')
        start = rng.Start
        after = rng.End
        if rng.Fields.Count or rng.Revisions.Count or rng.Paragraphs(1).Style.NameLocal.lower().startswith('toc'):
            report['in a tracked change or field'].append(hit)
        elif num.split('.')[0] != chapter:
            report['other chapter'].append(hit)
        elif num not in heads:
            report['no such heading'].append(hit)
        else:
            if num not in marks:
                marks[num] = bookmark(doc, heads[num], len(marks))
            target = doc.Range(start + len('Section '), start + len('Section ') + len(num))
            fld = doc.Fields.Add(target, -1, f'REF {marks[num]} \\w \\h', False)
            report['linked'].append(hit.rstrip('.'))
            after = fld.Result.End
        rng.SetRange(after, doc.Content.End)
    return chapter, report


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    folder = Path(sys.argv[1])
    files = sorted(p for p in folder.glob('*.docx') if not p.name.startswith('~$'))
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
                chapter, report = process(doc)
                doc.TrackRevisions = True
                if report['linked']:
                    doc.Save()
                if chapter is None:
                    print(f'{f.name}: skipped (Heading 1 not numbered)')
                    continue
                print(f'{f.name}: chapter {chapter}, linked {len(report["linked"])}')
                for k, v in report.items():
                    if k != 'linked' and v:
                        print(f'    left as text, {k}: {", ".join(v)}')
            finally:
                doc.Close(0)
    finally:
        word.Quit(0)


if __name__ == '__main__':
    main()
