"""Step 3 of the clean-template migration: check that a processed chapter has
exactly the content of its original.

Usage: python compare_content.py <original folder> <processed folder>

For each .docx present in both folders it compares:
  - the text with every tracked change accepted, and with every change rejected,
    word for word (read from the XML; invisible direction marks aside);
  - what each author inserted, deleted and moved, in characters (so Word
    splitting or joining revision elements does not count as a change);
  - comments, footnotes, endnotes, images, equations, links, fields;
  - every table's shape: rows, cells per row, merged cells.
Prints PASS or the differences. Reads only.
"""
import difflib
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

from remove_toc import strip_xml

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
# Invisible direction marks: Word rewrites w:dir/w:bdo containers as marks.
BIDI = re.compile('[‎‏‪-‮⁦-⁩]')
# Elements that make a revision non-empty.
DIGIT_DASH = re.compile(r'(\d)[‐‑–—-](\d)')   # a dash between digits, any form
CONTENT = {'t', 'delText', 'tab', 'br', 'drawing', 'pict', 'sym', 'instrText', 'fldChar',
           'footnoteReference', 'endnoteReference', 'object'}


def has_content(el):
    """True for an element that puts something on the page. A w:t holding
    no characters does not: Word drops a revision made only of those."""
    name = etree.QName(el).localname
    if name in ('t', 'delText'):
        return bool(el.text)
    return name in CONTENT


def result_elements(root):
    """The text, deleted text and images that sit inside a field's displayed
    result (between its separate and end marks, or inside a fldSimple)."""
    inside, stack = set(), []
    for el in root.iter(W + 't', W + 'delText', W + 'drawing', W + 'pict', W + 'fldChar'):
        if el.tag == W + 'fldChar':
            kind = el.get(W + 'fldCharType')
            if kind == 'begin':
                stack.append(False)
            elif kind == 'separate' and stack:
                stack[-1] = True
            elif kind == 'end' and stack:
                stack.pop()
        elif any(stack) or any(a.tag == W + 'fldSimple' for a in el.iterancestors()):
            inside.add(el)
    return inside


def read(z, name, drop_toc=False):
    """A part's XML; with drop_toc, document.xml without its tables of
    contents (what remove_toc.py takes out of a chapter)."""
    xml = z.read(name)
    return strip_xml(xml) if drop_toc and name == 'word/document.xml' else xml


def text(path, mode, skip_field_results=False, drop_toc=False):
    """Every word in the body, footnotes and endnotes, read from the XML with
    every tracked change accepted (mode 'accept') or rejected ('reject').
    Read directly rather than through pandoc, whose figure-caption guessing
    depends on paragraph properties the strip removes. skip_field_results
    leaves out what fields display (numbers, references, a table of contents),
    which is exactly what a field update is allowed to change."""
    z = zipfile.ZipFile(path)
    drop = ('del', 'moveFrom') if mode == 'accept' else ('ins', 'moveTo')
    words = []
    for name in ('word/document.xml', 'word/footnotes.xml', 'word/endnotes.xml'):
        if name not in z.namelist():
            continue
        root = etree.fromstring(read(z, name, drop_toc))
        # one entry per open complex field: [showing its result?, field deleted?]
        open_fields = []
        for el in root.iter(W + 't', W + 'delText', W + 'fldChar'):
            if el.tag == W + 'fldChar':
                kind = el.get(W + 'fldCharType')
                if kind == 'begin':
                    deleted = any(a.tag in (W + 'del', W + 'moveFrom') for a in el.iterancestors())
                    open_fields.append([False, deleted])
                elif kind == 'separate' and open_fields:
                    open_fields[-1][0] = True
                elif kind == 'end' and open_fields:
                    open_fields.pop()
                continue
            if el.tag == W + 'delText' and mode == 'accept':
                continue
            ancestors = list(el.iterancestors())
            if any(etree.QName(a).localname in drop and a.tag.startswith(W) for a in ancestors):
                continue
            # A deleted field's result is never set aside: it must stay deleted,
            # and an update that rewrote it as live text has to show up here.
            if skip_field_results and (any(r and not d for r, d in open_fields)
                                       or any(a.tag == W + 'fldSimple' for a in ancestors)):
                continue
            words.append(el.text or '')
    return ' '.join(BIDI.sub('', ''.join(words)).split()).split()


def part(z, name):
    return z.read(name).decode('utf-8', 'ignore') if name in z.namelist() else ''


def facts(path, skip_field_results=False, skip_field_marks=False, drop_toc=False):
    z = zipfile.ZipFile(path)
    xml = read(z, 'word/document.xml', drop_toc)
    d = xml.decode('utf-8', 'ignore')
    f = Counter()
    # Revisions per author. An empty run-level revision is not counted: Word
    # drops those on save. Paragraph-mark revisions (inside w:rPr) are always
    # counted, since they have no content by nature. With skip_field_results,
    # a revision whose whole content is a field's displayed result is not
    # counted either: a field update recomputes that text.
    root = etree.fromstring(xml)
    in_result = result_elements(root) if skip_field_results else set()
    # Measured in characters, not in revision elements: Word splits and joins
    # the elements freely (a field update splits an insertion around a
    # result), but what each author inserted and deleted must not change.
    # A text character counts 1, any other content (tab, break, field mark,
    # image) counts 1, a paragraph-mark revision counts 1.
    for kind in ('ins', 'del', 'moveFrom', 'moveTo'):
        for el in root.iter(W + kind):
            if etree.QName(el.getparent()).localname == 'rPr':
                f[f'{kind} by {el.get(W + "author")} (characters)'] += 1
                continue
            for x in el.iter():
                if not has_content(x) or (skip_field_results and x in in_result):
                    continue
                if skip_field_marks and etree.QName(x).localname in ('fldChar', 'instrText'):
                    continue
                n = len(x.text) if etree.QName(x).localname in ('t', 'delText') else 1
                f[f'{kind} by {el.get(W + "author")} (characters)'] += n
    f['comments'] = len(re.findall(r'<w:comment\b', part(z, 'word/comments.xml')))
    f['comment anchors'] = len(re.findall(r'<w:commentRangeStart\b', d))
    f['footnotes'] = len(re.findall(r'<w:footnoteReference\b', d))
    f['endnotes'] = len(re.findall(r'<w:endnoteReference\b', d))
    f['images'] = sum(1 for x in root.iter(W + 'drawing', W + 'pict') if x not in in_result)
    f['equations'] = len(re.findall(r'<m:oMath>', d))
    # Links counted in either form: Word's insert-from-file rewrites a
    # <w:hyperlink w:anchor> (a table of contents entry) as a HYPERLINK \l field.
    instr = [x.text or '' for x in root.iter(W + 'instrText')] + \
            [x.get(W + 'instr') or '' for x in root.iter(W + 'fldSimple')]
    hyperlink_fields = sum(1 for i in instr if re.match(r'\s*HYPERLINK\b', i))
    f['links'] = len(re.findall(r'<w:hyperlink\b', d)) + hyperlink_fields
    f['fields'] = (len(re.findall(r'w:fldCharType="begin"', d)) + len(re.findall(r'<w:fldSimple\b', d))
                   - hyperlink_fields)
    tables = []
    for t in re.finditer(r'(?s)<w:tbl>.*?</w:tbl>', d):
        rows = re.findall(r'(?s)<w:tr\b.*?</w:tr>', t.group(0))
        shape = tuple((len(re.findall(r'<w:tc>', r)), len(re.findall(r'<w:gridSpan\b', r)),
                       len(re.findall(r'<w:vMerge\b', r))) for r in rows)
        tables.append(shape)
    f['tables'] = len(tables)
    return f, tables


def compare(orig, new, allow_table_shape=False, field_update=False, new_fields=False, drop_toc=False):
    """Differences between two files. allow_table_shape skips the table-shape
    test, for a step whose job is to restructure table rows. field_update is
    for a Word field update: field results and the number of fields (a table
    of contents rebuilds its entry fields) may change, nothing else may.
    new_fields is for a step that turns typed text into fields showing the
    same text (cross-references): the number of fields may change, a dash
    between digits may change form (the field shows the caption's own), and
    field markers inside a tracked insertion are not counted as inserted text.
    drop_toc sets the tables of contents aside on both sides, for a chapter
    remove_toc.py has taken them out of: everything else must still match."""
    problems = []
    for mode in ('accept', 'reject'):
        a, b = text(orig, mode, field_update, drop_toc), text(new, mode, field_update, drop_toc)
        if new_fields:
            # a linked "Figure 4-14" shows the caption's own "Figure 4–14"
            a, b = [[DIGIT_DASH.sub(r'\1-\2', w) for w in x] for x in (a, b)]
        if a != b:
            diff = [d for d in difflib.ndiff(a, b) if d[0] in '+-'][:12]
            problems.append(f'text ({mode} all) differs: ' + ' '.join(diff))
    fa, ta = facts(orig, field_update, new_fields, drop_toc)
    fb, tb = facts(new, field_update, new_fields, drop_toc)
    for k in sorted(set(fa) | set(fb)):
        if (field_update or new_fields) and k == 'fields':
            continue
        if field_update and k == 'links':   # a table of contents rebuilds its entries
            continue
        if fa[k] != fb[k]:
            problems.append(f'{k}: {fa[k]} -> {fb[k]}')
    if ta != tb and not allow_table_shape:
        bad = [i + 1 for i, (x, y) in enumerate(zip(ta, tb)) if x != y]
        problems.append(f'table shape differs: tables {bad or "count"}')
    return problems, fa


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    a, b = Path(sys.argv[1]), Path(sys.argv[2])
    ok = True
    for orig in sorted(a.glob('*.docx')):
        new = b / orig.name
        if not new.exists() or orig.name.startswith('~$'):
            continue
        problems, fa = compare(orig, new)
        revs = sum(v for k, v in fa.items() if k.startswith(('ins ', 'del ', 'move')))
        if problems:
            ok = False
            print(f'FAIL {orig.name}')
            for p in problems:
                print(f'     {p}')
        else:
            print(f'PASS {orig.name}  ({revs} tracked characters, {fa["tables"]} tables, {fa["images"]} images, '
                  f'{fa["equations"]} equations, {fa["comments"]} comments)')
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
