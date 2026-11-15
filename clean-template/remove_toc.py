"""Remove the table of contents and the lists of tables and figures from
chapter files: chapters do not get them, the front matter does.

Usage: python remove_toc.py <folder> [--keep PATTERN ...]

Edits the files in place. Files matching a --keep pattern (front matter:
"*Chapter-00*", "*Preface*") are left alone.

Removed, for each TOC field (contents, list of tables, list of figures):
  - the paragraphs from its begin mark to its end mark (the entries);
  - next to them, empty paragraphs (page breaks included) and label
    paragraphs: "Table of Contents", "List of Tables", "List of Figures",
    "Front Matter";
  - the section break after them, when nothing is left before it (the
    front-matter section had nothing else in it).
Bookmarks with either end removed go too (the headings' _Toc bookmarks).
Tracked changes inside the entries are stale field results and go with them;
they are counted in the report.

A block holding a comment, footnote, endnote, image, equation or table is
real content, not a table of contents: that file is left unchanged and the
run stops (exit code 1).
"""
import argparse
import fnmatch
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
LABEL = re.compile(r'^(front matter|(table of )?contents|list of (tables|figures|exhibits|illustrations))$', re.I)
CONTENT = [W + 'commentRangeStart', W + 'commentReference', W + 'footnoteReference', W + 'endnoteReference',
           W + 'drawing', W + 'pict', W + 'object', M + 'oMath', W + 'tbl']


def text(el):
    return ''.join(t.text or '' for t in el.iter(W + 't', W + 'delText')).strip()


def is_empty(p):
    return not text(p) and not any(True for _ in p.iter(*CONTENT))


def toc_spans(body):
    """(first, last) body-child indexes of each outermost TOC field."""
    spans, stack = [], []   # stack entries: [instr, separated, start index]
    for i, child in enumerate(body):
        if child.tag != W + 'p':
            continue
        for el in child.iter(W + 'fldChar', W + 'instrText'):
            if el.tag == W + 'instrText':
                if stack and not stack[-1][1]:
                    stack[-1][0] += el.text or ''
                continue
            kind = el.get(W + 'fldCharType')
            if kind == 'begin':
                stack.append(['', False, i])
            elif kind == 'separate' and stack:
                stack[-1][1] = True
            elif kind == 'end' and stack:
                instr, _, start = stack.pop()
                outer_toc = any(s[0].split()[:1] == ['TOC'] for s in stack)
                if instr.split()[:1] == ['TOC'] and not outer_toc:
                    spans.append((start, i))
    return spans


def blocks(body):
    """Merged, extended runs of body-child indexes to remove."""
    kids = list(body)
    removable = set()
    for first, last in toc_spans(body):
        lo, hi = first, last
        while lo > 0 and kids[lo - 1].tag == W + 'p' and (is_empty(kids[lo - 1]) or LABEL.match(text(kids[lo - 1]))) \
                and kids[lo - 1].find(f'{W}pPr/{W}sectPr') is None:
            lo -= 1
        while hi + 1 < len(kids) and kids[hi + 1].tag == W + 'p' and is_empty(kids[hi + 1]) \
                and kids[hi + 1].find(f'{W}pPr/{W}sectPr') is None:
            hi += 1
        removable.update(range(lo, hi + 1))
    # the section break right after, when nothing is kept before it
    for i, k in enumerate(kids):
        if (k.tag == W + 'p' and k.find(f'{W}pPr/{W}sectPr') is not None and is_empty(k)
                and i not in removable and removable and all(j in removable for j in range(i))):
            removable.add(i)
            break
    return [kids[i] for i in sorted(removable)]


def remove_toc(root):
    """Remove the tables of contents from a document.xml root, in place.
    Returns (report dict, problem or None). On a problem nothing is removed."""
    body = root.find(W + 'body')
    gone = blocks(body)
    report = {'fields': len(toc_spans(body)), 'paragraphs': len(gone),
              'labels': [text(p) for p in gone if LABEL.match(text(p))],
              'tracked': sum(1 for p in gone for _ in p.iter(W + 'ins', W + 'del'))}
    found = sorted({etree.QName(x).localname for p in gone for x in p.iter(*CONTENT)})
    if found:
        return report, 'holds ' + ', '.join(found)
    ids = {b.get(W + 'id') for p in gone for b in p.iter(W + 'bookmarkStart', W + 'bookmarkEnd')}
    for p in gone:
        body.remove(p)
    for b in list(root.iter(W + 'bookmarkStart', W + 'bookmarkEnd')):
        if b.get(W + 'id') in ids:
            b.getparent().remove(b)
    return report, None


def strip_xml(xml):
    """document.xml bytes with the tables of contents removed (for the check)."""
    root = etree.fromstring(xml)
    _, problem = remove_toc(root)
    return xml if problem else etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('folder', type=Path)
    ap.add_argument('--keep', nargs='*', default=[])
    a = ap.parse_args()
    problems = []
    for f in sorted(p for p in a.folder.glob('*.docx') if not p.name.startswith('~$')):
        if any(fnmatch.fnmatch(f.name, k) for k in a.keep):
            print(f'{f.name}: front matter, kept')
            continue
        zin = zipfile.ZipFile(f)
        root = etree.fromstring(zin.read('word/document.xml'))
        report, problem = remove_toc(root)
        if problem:
            problems.append(f'{f.name}: table of contents block {problem}; left unchanged')
            continue
        if not report['fields']:
            print(f'{f.name}: no table of contents')
            continue
        print(f'{f.name}: removed {report["fields"]} TOC field(s), {report["paragraphs"]} paragraphs '
              f'({", ".join(report["labels"]) or "no headings"}); '
              f'{report["tracked"]} tracked change(s) inside the entries')
        new = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        tmp = f.with_suffix('.tmp')
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                zout.writestr(info, new if info.filename == 'word/document.xml' else zin.read(info.filename))
        zin.close()
        tmp.replace(f)
    if problems:
        sys.exit('\n'.join(problems))


if __name__ == '__main__':
    main()
