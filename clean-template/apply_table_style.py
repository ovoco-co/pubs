"""Give every data table the same table style and options, and remove the
direct table formatting that would override it.

Usage: python apply_table_style.py <file.docx | folder> [--style HM-Data] [--dry-run]

Edits the files in place; skips any file open in Word. Callout boxes and the
chapter banner are left alone (the same test as set_table_text.py).

For each data table:
  SET     the table style; the same options on every table (header row on,
          banded rows on, first/last column off); the header row repeats on
          each page. A total row (lastRow) an author switched on is kept.
  REMOVE  direct table and cell borders, cell shading, cell margins, the
          stale conditional-format flags Word caches on rows and cells
          (cnfStyle), and direct text color and shading on runs in the table,
          so the style's look shows through. White text left behind on a
          now-light row would be invisible, which is why run color goes too.
  PIN     the width: the table keeps its own column grid, scaled down to
          the text width if it is wider, on a fixed layout, so tables that
          share a grid render alike and decimal tabs (a later step) land
          where they are computed.
  KEEP    merged cells, vertical alignment, no-wrap, row heights, and all
          other run formatting (bold, italic, superscript, subscript,
          highlight).
Callout boxes on the Sidebar Table style are floated in the right margin,
3" wide, with text wrapping round them (the float cannot live in a table
style). Callouts on any other style, and the chapter banner, are untouched.
No revision is created.
"""
import argparse
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

from set_table_text import classify

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
LOOK = {'firstRow': '1', 'lastRow': '0', 'firstColumn': '0', 'lastColumn': '0',
        'noHBand': '0', 'noVBand': '1'}
# tblLook's legacy hex form: firstRow 0x20, lastRow 0x40, firstColumn 0x80,
# lastColumn 0x100, noHBand 0x200, noVBand 0x400.
LOOK_BITS = {'firstRow': 0x20, 'lastRow': 0x40, 'firstColumn': 0x80, 'lastColumn': 0x100,
             'noHBand': 0x200, 'noVBand': 0x400}
# CT_TblPr child order, so a new element goes in the right place.
TEXT_WIDTH = 9360          # twips: 6.5" between the template's 1" margins
SIDEBAR_FLOAT = {'leftFromText': '180', 'rightFromText': '180', 'vertAnchor': 'text',
                 'horzAnchor': 'margin', 'tblpXSpec': 'right', 'tblpY': '-5'}
SIDEBAR_WIDTH = '4325'     # twips, 3"
TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
               'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders',
               'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']


def child(parent, name, create=True):
    el = parent.find(W + name)
    if el is None and create:
        el = etree.Element(W + name)
        pos = TBLPR_ORDER.index(name)
        after = [c for c in parent if etree.QName(c).localname in TBLPR_ORDER[:pos]]
        if after:
            after[-1].addnext(el)
        else:
            parent.insert(0, el)
    return el


def remove_all(root, path, counter, label):
    for el in list(root.iter(*[W + p for p in path])):
        el.getparent().remove(el)
        counter[label] += 1


def pin_width(tbl, c, text_width=TEXT_WIDTH):
    """Fixed layout at the table's own grid, scaled to the text width if
    wider. Every cell's tcW is rewritten from the grid, honouring gridSpan."""
    grid = tbl.find(W + 'tblGrid')
    cols = grid.findall(W + 'gridCol') if grid is not None else []
    widths = [int(g.get(W + 'w') or 0) for g in cols]
    if not widths or sum(widths) <= 0:
        return
    total = sum(widths)
    if total > text_width:
        widths = [round(w * text_width / total) for w in widths]
        widths[-1] += text_width - sum(widths)
        c['tables scaled to the text width'] += 1
    for g, w in zip(cols, widths):
        g.set(W + 'w', str(w))
    for tr in tbl.findall(W + 'tr'):
        i = 0
        for tc in tr.findall(W + 'tc'):
            tcpr = tc.find(W + 'tcPr')
            if tcpr is None:
                tcpr = etree.Element(W + 'tcPr')
                tc.insert(0, tcpr)
            span = tcpr.find(W + 'gridSpan')
            n = int(span.get(W + 'val')) if span is not None else 1
            tcw = tcpr.find(W + 'tcW')
            if tcw is None:
                tcw = etree.Element(W + 'tcW')
                tcpr.insert(0, tcw)
            tcw.attrib.clear()
            tcw.set(W + 'w', str(sum(widths[i:i + n]) or widths[-1]))
            tcw.set(W + 'type', 'dxa')
            i += n
    tblpr = tbl.find(W + 'tblPr')
    tw = child(tblpr, 'tblW')
    tw.attrib.clear()
    tw.set(W + 'w', str(sum(widths)))
    tw.set(W + 'type', 'dxa')
    child(tblpr, 'tblLayout').set(W + 'type', 'fixed')
    c['widths pinned'] += 1


def float_callout(tbl, c):
    tblpr = tbl.find(W + 'tblPr')
    ts = tblpr.find(W + 'tblStyle') if tblpr is not None else None
    if ts is None or ts.get(W + 'val') != 'SidebarTable':
        return
    pp = child(tblpr, 'tblpPr')
    pp.attrib.clear()
    for k, v in SIDEBAR_FLOAT.items():
        pp.set(W + k, v)
    child(tblpr, 'tblOverlap').set(W + 'val', 'never')
    tw = child(tblpr, 'tblW')
    tw.attrib.clear()
    tw.set(W + 'w', SIDEBAR_WIDTH)
    tw.set(W + 'type', 'dxa')
    c['callouts floated'] += 1


def style_table(tbl, style, c):
    tblpr = tbl.find(W + 'tblPr')
    if tblpr is None:
        tblpr = etree.Element(W + 'tblPr')
        tbl.insert(0, tblpr)
    child(tblpr, 'tblStyle').set(W + 'val', style)

    old = tblpr.find(W + 'tblLook')
    total_row = old is not None and old.get(W + 'lastRow') == '1'
    look = child(tblpr, 'tblLook')
    look.attrib.clear()
    vals = dict(LOOK, lastRow='1' if total_row else '0')
    bits = 0
    for k, v in vals.items():
        look.set(W + k, v)
        bits |= LOOK_BITS[k] if v == '1' else 0
    look.set(W + 'val', '%04X' % bits)
    if total_row:
        c['total rows kept'] += 1

    for name in ('tblBorders', 'shd', 'tblCellMar'):
        el = tblpr.find(W + name)
        if el is not None:
            tblpr.remove(el)
            c[f'table {name}'] += 1

    for tcpr in tbl.iter(W + 'tcPr'):
        for name in ('tcBorders', 'shd', 'tcMar', 'cnfStyle'):
            el = tcpr.find(W + name)
            if el is not None:
                tcpr.remove(el)
                c[f'cell {name}'] += 1
    for trpr in tbl.iter(W + 'trPr'):
        el = trpr.find(W + 'cnfStyle')
        if el is not None:
            trpr.remove(el)
            c['row cnfStyle'] += 1
    for ppr in tbl.iter(W + 'pPr'):
        el = ppr.find(W + 'cnfStyle')
        if el is not None:
            ppr.remove(el)
            c['paragraph cnfStyle'] += 1
    for rpr in tbl.iter(W + 'rPr'):
        for name in ('color', 'shd'):
            el = rpr.find(W + name)
            if el is not None:
                rpr.remove(el)
                c[f'run {name}'] += 1

    rows = tbl.findall(W + 'tr')
    if rows:
        trpr = rows[0].find(W + 'trPr')
        if trpr is None:
            trpr = etree.Element(W + 'trPr')
            tblprex = rows[0].find(W + 'tblPrEx')
            (tblprex.addnext(trpr) if tblprex is not None else rows[0].insert(0, trpr))
        if trpr.find(W + 'tblHeader') is None:
            etree.SubElement(trpr, W + 'tblHeader')
            c['header row set to repeat'] += 1


def process(xml, style):
    root = etree.fromstring(xml)
    c = Counter()
    for tbl in root.find(W + 'body').findall(W + 'tbl'):
        kind = classify(tbl)
        c[f'{kind} tables'] += 1
        if kind == 'data':
            style_table(tbl, style, c)
            pin_width(tbl, c)
        elif kind == 'callout':
            float_callout(tbl, c)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True), c


def main():
    ap = argparse.ArgumentParser(description='Apply one table style to every data table.')
    ap.add_argument('src', type=Path)
    ap.add_argument('--style', default='HM-Data', help='table styleId (default HM-Data)')
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
        styles = zin.read('word/styles.xml').decode('utf-8')
        if f'w:styleId="{a.style}"' not in styles:
            print(f'SKIPPED (no {a.style} style in the file): {f.name}')
            continue
        new, c = process(zin.read('word/document.xml'), a.style)
        print(f'{f.name}: ' + ', '.join(f'{k} {v}' for k, v in c.items()))
        if a.dry_run or not c.get('data tables'):
            continue
        tmp = f.with_suffix('.tmp')
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                zout.writestr(info, new if info.filename == 'word/document.xml' else zin.read(info.filename))
        zin.close()
        tmp.replace(f)


if __name__ == '__main__':
    main()
