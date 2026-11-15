"""Check table and figure captions before their numbers are updated.

Usage: python caption_check.py <folder>

Edits the files in place, for one thing only:
  - a Caption_Figure paragraph with no "Figure" label that follows a figure
    caption is a description of that figure, so it moves to Caption_Figure_2.
Reported, never changed:
  - empty caption paragraphs, and captions whose title is a placeholder
    ("xx", "text", "TBD"): the field update numbers a placeholder like any
    other caption, so every table or figure after it moves up a number;
  - captions whose number is typed rather than a field. Front matter and
    appendix captions number by letter ("Figure P-1", "Table B-1") and stay
    typed by design; those are listed as kept.
A typed CHAPTER-numbered caption ("Table 3-4" with no SEQ field) would show a
stale number after the field update, so it stops the run (exit code 1) until
it is converted or excused.
"""
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
CAPTIONS = {'CaptionTable': 'Table', 'CaptionFigure': 'Figure'}
PLACEHOLDER = re.compile(r'^(x+|text|tbd|title|caption|placeholder|\[.*\])?$', re.I)
CAPTION_PREFIX = re.compile(r'^\s*(Table|Figure)[\s ]*[\w‐-—.-]*\.?\s*')
LETTER_NUMBER = re.compile(r'^\s*(Table|Figure)[\s ]+[A-Z][‐‑–—-]\d')


def pstyle(p):
    ps = p.find(f'{W}pPr/{W}pStyle')
    return ps.get(W + 'val') if ps is not None else None


def text(p):
    return ''.join(t.text or '' for t in p.iter(W + 't'))


def has_seq(p):
    instr = ' '.join(x.text or '' for x in p.iter(W + 'instrText'))
    instr += ' '.join(f.get(W + 'instr') or '' for f in p.iter(W + 'fldSimple'))
    return re.search(r'\bSEQ\s', instr) is not None


def process(xml):
    root = etree.fromstring(xml)
    paras = list(root.iter(W + 'p'))
    moved, empty, letter, typed, placeholder = [], [], [], [], []
    for i, p in enumerate(paras):
        style = pstyle(p)
        if style not in CAPTIONS:
            continue
        t = text(p)
        if not t.strip():
            empty.append(style)
            continue
        if re.match(rf'\s*{CAPTIONS[style]}\b', t):
            if PLACEHOLDER.match(CAPTION_PREFIX.sub('', t, count=1).strip()):
                placeholder.append(t[:60])
            if not has_seq(p):
                (letter if LETTER_NUMBER.match(t) else typed).append(t[:60])
            continue
        if style == 'CaptionFigure' and i and pstyle(paras[i - 1]) in ('CaptionFigure', 'CaptionFigure2'):
            p.find(f'{W}pPr/{W}pStyle').set(W + 'val', 'CaptionFigure2')
            moved.append(t[:60])
    return (etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True),
            moved, empty, letter, typed, placeholder)


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    folder = Path(sys.argv[1])
    blocking = []
    for f in sorted(p for p in folder.glob('*.docx') if not p.name.startswith('~$')):
        zin = zipfile.ZipFile(f)
        new, moved, empty, letter, typed, placeholder = process(zin.read('word/document.xml'))
        lines = []
        lines += [f'    moved to Caption_Figure_2: "{t}"' for t in moved]
        lines += [f'    empty caption paragraph ({s}), left for review' for s in empty]
        lines += [f'    placeholder caption, left for review: "{t}"' for t in placeholder]
        lines += [f'    typed letter number kept: "{t}"' for t in letter]
        lines += [f'    TYPED CHAPTER NUMBER, needs converting: "{t}"' for t in typed]
        print(f'{f.name}: ' + ('ok' if not lines else ''))
        for line in lines:
            print(line)
        blocking += [f'{f.name}: {t}' for t in typed]
        if moved:
            tmp = f.with_suffix('.tmp')
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    zout.writestr(info, new if info.filename == 'word/document.xml' else zin.read(info.filename))
            zin.close()
            tmp.replace(f)
    if blocking:
        sys.exit('Captions with typed chapter numbers would go stale on the field update:\n  ' + '\n  '.join(blocking))


if __name__ == '__main__':
    main()
