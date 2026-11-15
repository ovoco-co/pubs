"""Step 1 of the clean-template migration: strip hand formatting from Word
chapters, keeping the content, the styles and every tracked change.

Usage:
    python strip_formatting.py <file.docx | folder> <out_folder> [--template NIST-Template.docx]
                               [--map Old=New ...]

Writes <out_folder>/<name>.docx (originals are never touched) and
<out_folder>/strip-report.md. Works on the XML only; Word is not needed, and
no revision is ever created, so the existing insertions, deletions and
comments come through exactly as they were.

KEPT on runs:       style, bold, italic, superscript/subscript, highlight,
                    hidden, language, no-proof; inside tables also color and
                    shading. Equations and symbol-font characters are left
                    exactly as they are.
KEPT on paragraphs: style, borders and shading (callout boxes, reported),
                    tab stops, section breaks; inside tables also indents and
                    alignment (they carry row hierarchy).
DROPPED:            fonts, sizes, color, underline, spacing, indents,
                    alignment, keep-with-next and the like, and tracked
                    FORMATTING changes (the formatting they describe is gone).
LISTS:              numbering applied directly to a paragraph is kept as it
                    is and reported (it may be a numbered heading, a panel
                    label or a deliberate restart).
SECTIONS:           kept, with the template's margins; their own headers and
                    footers are dropped so the template's carry through.
"""
import argparse
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W, 'm': M}
q = lambda tag: '{%s}%s' % (W, tag)

RUN_KEEP = {'rStyle', 'b', 'bCs', 'i', 'iCs', 'vertAlign', 'highlight', 'vanish', 'specVanish',
            'webHidden', 'noProof', 'lang', 'rtl', 'cs', 'oMath',
            # paragraph-mark revision markers (pPr/rPr): a tracked paragraph split or join
            'ins', 'del', 'moveFrom', 'moveTo'}
RUN_KEEP_IN_TABLE = RUN_KEEP | {'color', 'shd'}
PARA_KEEP = {'pStyle', 'numPr', 'pBdr', 'shd', 'tabs', 'cnfStyle', 'rPr', 'sectPr'}
PARA_KEEP_IN_TABLE = PARA_KEEP | {'ind', 'jc'}
SYMBOL_FONTS = re.compile(r'^(Symbol|Wingdings.*|Webdings|MT Extra|Cambria Math|Marlett)$', re.I)
TEMPLATE_MARGINS = {'top': '1440', 'right': '1440', 'bottom': '1440', 'left': '1440',
                    'header': '432', 'footer': '432', 'gutter': '0'}
DEFAULT_MAP = {'NoSpacing': 'BodyText'}


def local(el):
    return etree.QName(el).localname


def in_table(el):
    return any(local(a) == 'tc' and a.tag == q('tc') for a in el.iterancestors())


def in_math(el):
    return any(a.tag.startswith('{%s}' % M) for a in el.iterancestors())


def symbol_font(rpr):
    f = rpr.find(q('rFonts'))
    if f is None:
        return False
    return any(SYMBOL_FONTS.match(f.get(q(a)) or '') for a in ('ascii', 'hAnsi', 'cs'))


class Stripper:
    def __init__(self, numbering_xml, style_map, template_ids):
        self.map = style_map
        self.template_ids = template_ids
        self.run_dropped = Counter()
        self.para_dropped = Counter()
        self.mapped = Counter()
        self.unknown_styles = Counter()
        self.lists = []
        self.callouts = []
        self.num_fmt = self._read_numbering(numbering_xml)

    @staticmethod
    def _read_numbering(xml):
        """numId -> {ilvl: numFmt}."""
        if xml is None:
            return {}
        root = etree.fromstring(xml)
        absfmt = {}
        for a in root.findall('w:abstractNum', NS):
            absfmt[a.get(q('abstractNumId'))] = {
                l.get(q('ilvl')): (l.find('w:numFmt', NS).get(q('val')) if l.find('w:numFmt', NS) is not None else 'decimal')
                for l in a.findall('w:lvl', NS)}
        out = {}
        for n in root.findall('w:num', NS):
            ref = n.find('w:abstractNumId', NS)
            if ref is not None:
                out[n.get(q('numId'))] = absfmt.get(ref.get(q('val')), {})
        return out

    def _remap(self, el):
        v = el.get(q('val'))
        if v in self.map:
            el.set(q('val'), self.map[v])
            self.mapped[f'{v} -> {self.map[v]}'] += 1
        elif self.template_ids is not None and v not in self.template_ids:
            self.unknown_styles[v] += 1

    def strip_rpr(self, rpr, table):
        if in_math(rpr):
            return
        keep = RUN_KEEP_IN_TABLE if table else RUN_KEEP
        sym = symbol_font(rpr)
        for child in list(rpr):
            name = local(child)
            if name == 'rStyle':
                self._remap(child)
            if name in keep or (sym and name == 'rFonts'):
                continue
            rpr.remove(child)
            self.run_dropped[name] += 1

    def strip_ppr(self, ppr, text):
        table = in_table(ppr)
        keep = PARA_KEEP_IN_TABLE if table else PARA_KEEP
        for child in list(ppr):
            name = local(child)
            if name in keep:
                continue
            ppr.remove(child)
            self.para_dropped[name] += 1
        ps = ppr.find(q('pStyle'))
        if ps is not None:
            self._remap(ps)
        if ppr.find(q('pBdr')) is not None or ppr.find(q('shd')) is not None:
            if not table:
                self.callouts.append(text[:70])
        self._fix_list(ppr, text)

    def _fix_list(self, ppr, text):
        """Direct numbering is KEPT and reported, never converted. It is not
        always a ribbon bullet: a numbered heading, "(a)"/"(b)" figure panel
        labels and a list deliberately restarted at 1 all look the same, and
        moving any of them onto a list style changes the numbers in the text."""
        np_ = ppr.find(q('numPr'))
        if np_ is None:
            return
        nid = np_.find(q('numId'))
        lvl = np_.find(q('ilvl'))
        nid = nid.get(q('val')) if nid is not None else None
        lvl = lvl.get(q('val')) if lvl is not None else '0'
        if nid in (None, '0'):
            return                      # numId 0 deliberately switches a list off
        fmt = self.num_fmt.get(nid, {}).get(lvl, 'decimal')
        ps = ppr.find(q('pStyle'))
        cur = ps.get(q('val')) if ps is not None else 'Normal'
        self.lists.append((cur, fmt, text[:60]))

    def strip_sect(self, sect):
        for child in list(sect):
            name = local(child)
            if name in ('headerReference', 'footerReference', 'pgNumType', 'titlePg'):
                sect.remove(child)
        mar = sect.find(q('pgMar'))
        if mar is not None:
            for k, v in TEMPLATE_MARGINS.items():
                mar.set(q(k), v)

    def run(self, xml):
        root = etree.fromstring(xml)
        # tracked formatting changes go first: the formatting they record is being removed
        for tag in ('rPrChange', 'pPrChange', 'sectPrChange'):
            for el in list(root.iter(q(tag))):
                el.getparent().remove(el)
                self.run_dropped[tag] += 1
        for p in root.iter(q('p')):
            text = ''.join(t.text or '' for t in p.iter(q('t')))
            ppr = p.find(q('pPr'))
            if ppr is not None:
                self.strip_ppr(ppr, text)
                mark = ppr.find(q('rPr'))
                if mark is not None:
                    self.strip_rpr(mark, in_table(p))
        for r in root.iter(q('r')):
            rpr = r.find(q('rPr'))
            if rpr is not None:
                self.strip_rpr(rpr, in_table(r))
        for sect in root.iter(q('sectPr')):
            self.strip_sect(sect)
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def template_style_ids(path):
    if not path:
        return None
    xml = zipfile.ZipFile(path).read('word/styles.xml').decode('utf-8')
    return set(re.findall(r'w:styleId="([^"]+)"', xml))


def strip_file(src, dst, style_map, template_ids):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    numbering = zin.read('word/numbering.xml') if 'word/numbering.xml' in names else None
    s = Stripper(numbering, style_map, template_ids)
    out = {}
    for part in ('word/document.xml', 'word/footnotes.xml', 'word/endnotes.xml'):
        if part in names:
            out[part] = s.run(zin.read(part))
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, out.get(info.filename, zin.read(info.filename)))
    return s


def write_report(path, results):
    with open(path, 'w', encoding='utf-8') as f:
        f.write('# Strip report\n\nFormatting removed from each chapter. Text, styles, tracked insertions and '
                'deletions, comments, tables, images and equations were not changed.\n')
        for name, s in results:
            f.write(f'\n## {name}\n\n')
            f.write('- Run formatting removed: ' + (', '.join(f'{k} {v}' for k, v in s.run_dropped.most_common()) or 'none') + '\n')
            f.write('- Paragraph formatting removed: ' + (', '.join(f'{k} {v}' for k, v in s.para_dropped.most_common()) or 'none') + '\n')
            if s.mapped:
                f.write('- Styles mapped: ' + ', '.join(f'{k} ({v})' for k, v in s.mapped.items()) + '\n')
            if s.unknown_styles:
                f.write('- **Styles not in the template** (will come across as extra styles): '
                        + ', '.join(f'{k} ({v})' for k, v in s.unknown_styles.items()) + '\n')
            if s.lists:
                f.write(f'- **Numbering applied directly, kept as is ({len(s.lists)}):**\n')
                for cur, fmt, text in s.lists:
                    f.write(f'  - {cur}, {fmt}: "{text}"\n')
            if s.callouts:
                f.write(f'- **Boxed or shaded paragraphs kept ({len(s.callouts)}):**\n')
                for t in s.callouts:
                    f.write(f'  - "{t}"\n')


def main():
    ap = argparse.ArgumentParser(description='Strip hand formatting from Word chapters, keeping tracked changes.')
    ap.add_argument('src', type=Path)
    ap.add_argument('out', type=Path)
    ap.add_argument('--template', type=Path, help='NIST-Template.docx, to report styles it does not have')
    ap.add_argument('--map', nargs='*', default=[], help='Extra style mappings, Old=New (style ids)')
    a = ap.parse_args()
    style_map = dict(DEFAULT_MAP)
    style_map.update(dict(m.split('=', 1) for m in a.map))
    files = [a.src] if a.src.is_file() else sorted(p for p in a.src.glob('*.docx')
                                                    if not p.name.startswith('~$') and 'Template' not in p.name)
    if not files:
        sys.exit(f'No .docx files in {a.src}')
    a.out.mkdir(parents=True, exist_ok=True)
    ids = template_style_ids(a.template)
    results = []
    for f in files:
        lock = f.with_name('~$' + f.name[2:])
        if lock.exists():
            print(f'SKIPPED (open in Word): {f.name}')
            continue
        s = strip_file(f, a.out / f.name, style_map, ids)
        results.append((f.name, s))
        print(f'{f.name}: {sum(s.run_dropped.values())} run and {sum(s.para_dropped.values())} paragraph '
              f'properties removed, {len(s.lists)} directly numbered, {len(s.callouts)} box(es) kept')
    write_report(a.out / 'strip-report.md', results)
    print(f'Report: {a.out / "strip-report.md"}')


if __name__ == '__main__':
    main()
