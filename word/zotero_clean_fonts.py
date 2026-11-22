"""Strip the font Zotero stamps on its output, so citations and the reference list take the style's font.

Usage:
    python zotero_clean_fonts.py DOCX [DOCX ...] [--font "Times New Roman"] [--write]

Zotero hands formatted text (italics in a title, say) to Word as rich text, and Word gives any run
without a stated font its own default, Times New Roman, as direct formatting. The reference list
therefore ignores the Bibliography/References style's font, and so does any citation that carries
italics. Zotero rewrites that text on every Refresh, so run this after refreshing.

It removes w:rFonts naming --font (default Times New Roman) from:
  - runs inside a Zotero field's result (ZOTERO_ITEM citations and the ZOTERO_BIBL list), and
  - runs and paragraph marks in paragraphs on the Bibliography style.
Everything else on those runs (italic, bold, superscript) stays. Dry run by default; --write saves
in place, refusing a file that is open in Word.
"""
import argparse
import sys
import zipfile
from pathlib import Path

from lxml import etree

import restore_zotero_fields as R

w = R.w


def strip(rpr, font):
    """Drop rFonts if every font it names is `font` (theme hints and other fonts are left alone)."""
    if rpr is None:
        return 0
    n = 0
    for rf in rpr.findall(w("rFonts")):
        names = {v for k, v in rf.attrib.items() if k in (w("ascii"), w("hAnsi"), w("cs"), w("eastAsia"))}
        if names and names <= {font}:
            rpr.remove(rf); n += 1
    if len(rpr) == 0 and not rpr.attrib:
        rpr.getparent().remove(rpr)
    return n


def clean_part(root, font):
    n = 0
    # paragraphs on the Bibliography style: paragraph mark and every run
    for p in root.iter(w("p")):
        st = p.find(f"{w('pPr')}/{w('pStyle')}")
        if st is not None and st.get(w("val")) == "Bibliography":
            n += strip(p.find(f"{w('pPr')}/{w('rPr')}"), font)
            for r in p.iter(w("r")):
                n += strip(r.find(w("rPr")), font)
    # results of Zotero fields anywhere (citations, and a list on another style)
    stack = []
    for r in root.iter(w("r")):
        fc = r.find(w("fldChar"))
        if fc is not None:
            kind = fc.get(w("fldCharType"))
            if kind == "begin":
                stack.append(["", False])
            elif kind == "separate" and stack:
                stack[-1][1] = True
            elif kind == "end" and stack:
                stack.pop()
            continue
        it = r.find(w("instrText"))
        if it is not None and stack:
            stack[-1][0] += it.text or ""
            continue
        if stack and stack[-1][1] and "ADDIN ZOTERO_" in stack[-1][0]:
            n += strip(r.find(w("rPr")), font)
    return n


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", nargs="+")
    ap.add_argument("--font", default="Times New Roman")
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    for f in a.docx:
        path = Path(f)
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            parts = {i.filename: z.read(i.filename) for i in infos}
        total = 0
        for name in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
            if name in parts:
                root = etree.fromstring(parts[name])
                n = clean_part(root, a.font)
                if n:
                    parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                    total += n
        print(f"{path.name}: {total} '{a.font}' font setting(s) {'removed' if a.write else 'to remove'}")
        if a.write and total:
            if R.locked(path):
                print("  SKIPPED: open in Word (~$ lock file)")
                continue
            tmp = path.with_suffix(".tmp")
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
                for i in infos:
                    out.writestr(i.filename, parts[i.filename])
            tmp.replace(path)


if __name__ == "__main__":
    sys.exit(main())
