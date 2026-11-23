"""Install the appendix numbering in chapters, whether or not they have appendices yet.

Usage:
    python set_appendix_numbering.py DOCX [DOCX ...] [--style-source DOCX] [--write]

Does only what the appendix-ladder pass of Set-NISTStyleArchitecture.ps1 does, without the rest of
the architecture, so it can be run on chapters whose other styles must not change:

  - adds the multilevel list "Appendix A." / A.1. / A.1.1. / A.1.1.1., linked to AppendixHead and
    AppendixHead2-4 (an existing list with "Appendix %1." on AppendixHead is reused), and points
    each of those styles at it;
  - copies Heading1NoNumber from --style-source when the chapter lacks it, since an unnumbered
    "References" heading must not sit on AppendixHead once that style numbers.

Paragraphs are not touched. Dry run by default; --write saves in place, refusing an open file.
"""
import argparse
import re
import sys
import zipfile
from pathlib import Path

import restore_zotero_fields as R

LADDER = [("upperLetter", "Appendix %1.", "AppendixHead"), ("decimal", "%1.%2.", "AppendixHead2"),
          ("decimal", "%1.%2.%3.", "AppendixHead3"), ("decimal", "%1.%2.%3.%4.", "AppendixHead4")]


def style_block(xml, sid):
    return re.search(r'(?s)<w:style\b[^>]*\bw:styleId="' + re.escape(sid) + r'"[^>]*>.*?</w:style>', xml)


def ladder_list(numbering):
    """(numId, abstractNumId) of an existing appendix list, or None."""
    for m in re.finditer(r'(?s)<w:abstractNum\b[^>]*w:abstractNumId="(\d+)".*?</w:abstractNum>', numbering):
        if re.search(r'(?s)<w:lvl w:ilvl="0"[^>]*>(?:(?!</w:lvl>).)*<w:pStyle w:val="AppendixHead"/>'
                     r'(?:(?!</w:lvl>).)*<w:lvlText w:val="Appendix %1\."/>', m.group(0)):
            n = re.search(r'(?s)<w:num\b[^>]*w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="' + m.group(1) + r'"/>', numbering)
            if n:
                return n.group(1), m.group(1)
    return None


def add_ladder(numbering):
    abs_id = 1 + max([int(x) for x in re.findall(r'<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"', numbering)] + [0])
    num_id = 1 + max([int(x) for x in re.findall(r'<w:num\b[^>]*w:numId="(\d+)"', numbering)] + [0])
    lvls = ""
    for i in range(9):
        if i < len(LADDER):
            fmt, text, sid = LADDER[i]
            lvls += (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:pStyle w:val="{sid}"/>'
                     f'<w:suff w:val="space"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                     '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl>')
        else:
            lvls += (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/>'
                     '<w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>')
    abstract = f'<w:abstractNum w:abstractNumId="{abs_id}"><w:multiLevelType w:val="multilevel"/>{lvls}</w:abstractNum>'
    num = f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{abs_id}"/></w:num>'
    # schema order: every abstractNum before every num
    m = re.search(r"<w:num\b", numbering)
    numbering = numbering[:m.start()] + abstract + numbering[m.start():] if m else numbering.replace("</w:numbering>", abstract + "</w:numbering>", 1)
    numbering = numbering.replace("</w:numbering>", num + "</w:numbering>", 1)
    return numbering, str(num_id)


def wire(styles, sid, lvl, num_id):
    m = style_block(styles, sid)
    if not m:
        return styles, False
    b = re.sub(r"(?s)<w:numPr>.*?</w:numPr>|<w:numPr\s*/>", "", m.group(0))
    num_pr = f'<w:numPr><w:ilvl w:val="{lvl}"/><w:numId w:val="{num_id}"/></w:numPr>'
    pp = re.search(r"(?s)<w:pPr>(.*?)</w:pPr>", b)
    if pp:
        lead = re.match(r"(?:<w:(?:keepNext|keepLines|pageBreakBefore|framePr|widowControl)\b[^>]*/>)*", pp.group(1))
        inner = pp.group(1)[:lead.end()] + num_pr + pp.group(1)[lead.end():]
        b = b[:pp.start(1)] + inner + b[pp.end(1):]
    elif "<w:rPr>" in b:
        b = b.replace("<w:rPr>", f"<w:pPr>{num_pr}</w:pPr><w:rPr>", 1)
    else:
        b = b.replace("</w:style>", f"<w:pPr>{num_pr}</w:pPr></w:style>")
    return styles[:m.start()] + b + styles[m.end():], True


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", nargs="+")
    ap.add_argument("--style-source", help="chapter to copy Heading1NoNumber from when missing")
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    src_h1nn = None
    if a.style_source:
        with zipfile.ZipFile(a.style_source) as z:
            m = style_block(z.read("word/styles.xml").decode("utf-8"), "Heading1NoNumber")
            src_h1nn = m.group(0) if m else None
    for f in a.docx:
        path = Path(f)
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            parts = {i.filename: z.read(i.filename) for i in infos}
        styles = parts["word/styles.xml"].decode("utf-8")
        numbering = parts["word/numbering.xml"].decode("utf-8")
        notes = []
        found = ladder_list(numbering)
        if found:
            num_id = found[0]; notes.append(f"list {num_id} reused")
        else:
            numbering, num_id = add_ladder(numbering); notes.append(f"list {num_id} added")
        wired = []
        for lvl, (_, _, sid) in enumerate(LADDER):
            styles, ok = wire(styles, sid, lvl, num_id)
            if ok:
                wired.append(sid)
        notes.append("wired " + ", ".join(wired) if wired else "no AppendixHead styles")
        if not style_block(styles, "Heading1NoNumber"):
            if src_h1nn and style_block(styles, "Heading1"):
                styles = styles.replace("</w:styles>", src_h1nn + "</w:styles>", 1); notes.append("Heading1NoNumber added")
            else:
                notes.append("Heading1NoNumber MISSING (no --style-source)")
        print(f"{path.name}: " + "; ".join(notes))
        if a.write:
            if R.locked(path):
                print("  SKIPPED: open in Word"); continue
            parts["word/styles.xml"] = styles.encode("utf-8")
            parts["word/numbering.xml"] = numbering.encode("utf-8")
            tmp = path.with_suffix(".tmp")
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
                for i in infos:
                    out.writestr(i.filename, parts[i.filename])
            tmp.replace(path)


if __name__ == "__main__":
    sys.exit(main())
