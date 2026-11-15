"""Put the paragraphs of every data table onto Table Text, leaving callout
boxes and the chapter banner alone.

Usage: python set_table_text.py <file.docx | folder> [--dry-run]

Edits the files in place. Only w:pStyle changes: run formatting (italic,
superscript, subscript, bold, highlight) and paragraph indents and alignment
are untouched, and no revision is created.

Tables are told apart by shape and position, not by table style, since
authors' tables carry whatever style they were made with:
  callout box    a table with a single cell, or one on the Sidebar Table style
  chapter banner a table followed directly by a Heading 1 (empty paragraphs
                 between them aside), or one holding Header / Headerch
                 paragraphs
  data table     everything else
Inside a data table, paragraphs already on a table-specific style (Table Text,
Table Bullet, table footnotes, captions) keep it.
"""
import argparse
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TABLE_TEXT = 'TableText'
KEEP = {'TableText', 'TableBullet', 'TableandFigureFootnotes', 'CaptionTable', 'CaptionFigure',
        'CaptionFigure2', 'Caption'}
BANNER_STYLES = {'Header', 'Headerch', 'HeaderOdd', 'HeaderEven'}
CALLOUT_TABLE_STYLES = {'SidebarTable'}


def pstyle(p):
    ps = p.find(f'{W}pPr/{W}pStyle')
    return ps.get(W + 'val') if ps is not None else None


def opens_chapter(el):
    """True when the next non-empty paragraph after el is a Heading 1."""
    for nxt in el.itersiblings():
        if nxt.tag != W + 'p':
            return False
        if pstyle(nxt) == 'Heading1':
            return True
        if ''.join(t.text or '' for t in nxt.iter(W + 't')).strip():
            return False
    return False


def classify(tbl):
    ts = tbl.find(f'{W}tblPr/{W}tblStyle')
    if ts is not None and ts.get(W + 'val') in CALLOUT_TABLE_STYLES:
        return 'callout'
    if len(tbl.findall(f'{W}tr/{W}tc')) == 1:
        return 'callout'
    if opens_chapter(tbl) or any(pstyle(p) in BANNER_STYLES for p in tbl.iter(W + 'p')):
        return 'banner'
    return 'data'


def process(xml):
    root = etree.fromstring(xml)
    body = root.find(W + 'body')
    kinds = Counter()
    changed = Counter()
    for el in body:
        if el.tag != W + 'tbl':
            continue
        kind = classify(el)
        kinds[kind] += 1
        if kind != 'data':
            continue
        for p in el.iter(W + 'p'):       # nested tables included
            cur = pstyle(p)
            if cur in KEEP:
                continue
            ppr = p.find(W + 'pPr')
            if ppr is None:
                ppr = etree.SubElement(p, W + 'pPr')
                p.insert(0, ppr)
            ps = ppr.find(W + 'pStyle')
            if ps is None:
                ps = etree.Element(W + 'pStyle')
                ppr.insert(0, ps)
            ps.set(W + 'val', TABLE_TEXT)
            changed[cur or 'Normal'] += 1
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True), kinds, changed


def main():
    ap = argparse.ArgumentParser(description='Put data-table paragraphs onto Table Text.')
    ap.add_argument('src', type=Path)
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()
    files = [a.src] if a.src.is_file() else sorted(p for p in a.src.glob('*.docx') if not p.name.startswith('~$'))
    if not files:
        sys.exit(f'No .docx files in {a.src}')
    for f in files:
        if f.with_name('~$' + f.name[2:]).exists():
            print(f'SKIPPED (open in Word): {f.name}')
            continue
        zin = zipfile.ZipFile(f)
        new, kinds, changed = process(zin.read('word/document.xml'))
        summary = ', '.join(f'{v} {k}' for k, v in kinds.items()) or 'no tables'
        moved = ', '.join(f'{k} {v}' for k, v in changed.most_common()) or 'nothing'
        print(f'{f.name}: {summary}; moved to Table Text: {moved}')
        if a.dry_run or not changed:
            continue
        tmp = f.with_suffix('.tmp')
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                zout.writestr(info, new if info.filename == 'word/document.xml' else zin.read(info.filename))
        zin.close()
        tmp.replace(f)


if __name__ == '__main__':
    main()
