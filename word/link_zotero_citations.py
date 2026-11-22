"""Turn typed author-date citations into Zotero citation fields, as tracked changes.

Usage:
    python link_zotero_citations.py DOCX MAP.json --author "Name" [--write]

For a chapter whose citations and reference list were typed by hand (or whose footnotes hold
bare sources), once the sources are items in a Zotero library. MAP.json says which text is
which item; the script finds each citation in the document and footnotes and replaces it,
tracked under --author, with a ZOTERO_ITEM field whose instruction carries the item's URI and
CSL data, so Zotero > Refresh works at once:

  "cites"       {"FCC 2017-2018": KEY, ...}  parenthetical citations. A group "(A 2018; B 2019)"
                or "(A 2018, B 2019)" becomes one field; a part may carry a prefix
                ("downloaded from FCC 2020"). A group with any unknown part is left alone.
  "narrative"   {"Cordova et al.|2021": KEY}  "Cordova et al. (2021)": the "(2021)" becomes a
                suppress-author field and the name stays typed.
  "edits"       [{"find": text, "new": text}]  tracked plain-text edits, applied first.
  "replace"     [{"find": text, "key": KEY, "mode": "paren"|"narrative", "lead": text,
                  "locator": "...", "label": "section"}]  one-off spans, e.g. a bare URL or a
                statute citation; "lead" is typed text inserted before the field.
  "footnotes"   {footnote id: KEY}  a footnote that is only a source: its reference mark is
                deleted and a parenthetical citation inserted where it stood (before a
                preceding period or comma).
  "bibliography" {"mode": "replace"|"append", "keep": [prefix...], "heading": text,
                  "heading_style": style, "after": text, "uncited": [KEY...]}
                replace: the References-style paragraphs are deleted and a ZOTERO_BIBL field
                inserted in their place ("keep" spares paragraphs starting with a prefix).
                append: a heading and the field are inserted after the paragraph starting
                with "after". "uncited" keeps listed items that the text never cites.
  "prefs_from"  a Zotero document whose ZOTERO_PREF_* properties (style, field type) to copy.
  "person_from" a document whose people.xml entry for --author to copy (account identity).

Item data and the result text of the reference list come from Zotero's local API
(Settings > Advanced > "Allow other applications ... to communicate with Zotero"). Result text is
provisional; Refresh in Word rewrites it in the document's style.

Dry run by default; --write saves in place (refuses if a ~$ lock file exists).
"""
import argparse
import copy
import datetime
import html
import json
import random
import re
import string
import sys
import urllib.request
import zipfile
from pathlib import Path

from lxml import etree

import restore_zotero_fields as R

W = R.W
w = R.w
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
API = "http://127.0.0.1:23119/api"
SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
XML_SPACE = R.XML_SPACE


# ---------------------------------------------------------------- Zotero
def api(path):
    with urllib.request.urlopen(f"{API}/{path}", timeout=60) as r:
        return r.read().decode("utf-8")


def fetch_items(library, keys):
    csl, bib = {}, {}
    keys = sorted(set(keys))
    for i in range(0, len(keys), 50):
        chunk = ",".join(keys[i:i + 50])
        for item in json.loads(api(f"{library}/items?itemKey={chunk}&format=csljson")):
            csl[item["id"].rsplit("/", 1)[1]] = item
        for item in json.loads(api(f"{library}/items?itemKey={chunk}&format=json&include=bib&style=apa")):
            text = re.sub(r"<[^>]+>", "", item.get("bib", ""))
            bib[item["key"]] = html.unescape(re.sub(r"\s+", " ", text)).strip()
    missing = [k for k in keys if k not in csl]
    if missing:
        raise SystemExit(f"not in the library: {missing}")
    return csl, bib


def rand_id(n=8):
    return "".join(random.choice(string.ascii_letters + string.digits) for _ in range(n))


# ---------------------------------------------------------------- tracked-change builders
class Track:
    def __init__(self, author, start_id):
        self.author = author
        self.date = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:00Z")
        self.next_id = start_id

    def mark(self, tag):
        el = etree.Element(w(tag))
        el.set(w("id"), str(self.next_id)); self.next_id += 1
        el.set(w("author"), self.author); el.set(w("date"), self.date)
        return el


def to_deleted(run):
    for t in run.findall(w("t")):
        t.tag = w("delText")
    for t in run.findall(w("instrText")):
        t.tag = w("delInstrText")


def ptext(p):
    return "".join(R.run_text(r) for r in R.runs_of(p))


def text_run(text, rpr):
    r = etree.Element(w("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, w("t")); t.text = text; t.set(XML_SPACE, "preserve")
    return r


def field_runs(instr, result, rpr):
    return [R.field_run("begin", rpr), R.field_run(None, rpr, instr), R.field_run("separate", rpr),
            text_run(result, rpr), R.field_run("end", rpr)]


def place_after(anchor, ins, tr):
    """Put an insertion right after `anchor`. Inside another author's w:ins the remainder of
    that insertion is split off, since insertions cannot nest."""
    parent = anchor.getparent()
    if parent.tag == w("ins"):
        tail = [s for s in anchor.itersiblings()]
        parent.addnext(ins)
        if tail:
            shell = copy.deepcopy(parent)
            for k in list(shell):
                shell.remove(k)
            shell.set(w("id"), str(tr.next_id)); tr.next_id += 1
            for s in tail:
                shell.append(s)
            ins.addnext(shell)
    else:
        anchor.addnext(ins)


def delete_runs(runs, tr):
    d = tr.mark("del")
    runs[0].addprevious(d)
    for r in runs:
        to_deleted(r); d.append(r)
    parent = d.getparent()
    # a whole hyperlink deleted: insert after the hyperlink, not inside it
    if parent.tag == w("hyperlink") and not [r for r in parent.iter(w("r")) if r.getparent() is not d]:
        return parent
    return d


# ---------------------------------------------------------------- paragraph text model
def visible_runs(p):
    """Runs whose text is visible and editable: not deleted, not inside a field result or code."""
    out, depth = [], 0
    for r in R.runs_of(p):
        fc = r.find(w("fldChar"))
        if fc is not None:
            k = fc.get(w("fldCharType"))
            depth += 1 if k == "begin" else -1 if k == "end" else 0
            continue
        anc = r.getparent()
        deleted = False
        while anc is not None and anc is not p:
            if anc.tag in (w("del"), w("moveFrom")):
                deleted = True
            anc = anc.getparent()
        if depth == 0 and not deleted and r.find(w("t")) is not None:
            out.append(r)
    return out


def model(p):
    runs = visible_runs(p)
    spans, pos = [], 0
    for r in runs:
        n = len(R.run_text(r)); spans.append((r, pos, pos + n)); pos += n
    return spans, "".join(R.run_text(r) for r in runs)


def split_visible(p, offset):
    for r, a, b in model(p)[0]:
        if a < offset < b:
            pieces = R.explode(r)
            pos = a
            for piece in pieces:
                t = piece.find(w("t"))
                m = len(t.text or "") if t is not None else 0
                if pos < offset < pos + m:
                    k = offset - pos
                    tail = copy.deepcopy(piece)
                    t.text, tail.find(w("t")).text = t.text[:k], t.text[k:]
                    t.set(XML_SPACE, "preserve"); tail.find(w("t")).set(XML_SPACE, "preserve")
                    piece.addnext(tail)
                    return
                pos += m
            return


def span_runs(p, s, e):
    split_visible(p, s); split_visible(p, e)
    return [r for r, a, b in model(p)[0] if a >= s and b <= e and b > a]


# ---------------------------------------------------------------- citations
class Linker:
    def __init__(self, m, csl, library, tr, log):
        self.m, self.csl, self.lib, self.tr, self.log = m, csl, library, tr, log
        self.ids = {k: 100000 + i for i, k in enumerate(sorted(csl))}
        self.cited = set()
        self.count = 0
        keys = sorted(m.get("cites", {}), key=len, reverse=True)
        self.part_re = re.compile(r"^(?P<prefix>.*?)(?P<key>" + "|".join(re.escape(k) for k in keys) + r")$") if keys else None

    def item(self, key, suppress=False, prefix=None, locator=None, label=None):
        data = dict(self.csl[key]); data["id"] = self.ids[key]
        ci = {"id": self.ids[key], "uris": [f"http://zotero.org/{self.lib}/items/{key}"], "itemData": data}
        if suppress:
            ci["suppress-author"] = True
        if prefix:
            ci["prefix"] = prefix
        if locator:
            ci["locator"] = locator; ci["label"] = label or "page"
        self.cited.add(key)
        return ci

    def instr(self, items, result, note=0):
        body = {"citationID": rand_id(), "properties": {"formattedCitation": result, "plainCitation": result,
                "noteIndex": note}, "citationItems": items, "schema": SCHEMA}
        return " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(body, ensure_ascii=False) + " "

    def replace_span(self, p, s, e, items, result, lead=""):
        runs = span_runs(p, s, e)
        if not runs or len({r.getparent() for r in runs}) != 1:
            self.log.append(f"  SKIPPED (crosses a tracked change or field)  {result}")
            return False
        rpr = runs[0].find(w("rPr"))
        anchor = delete_runs(runs, self.tr)
        ins = self.tr.mark("ins")
        if lead:
            ins.append(text_run(lead, rpr))
        for el in field_runs(self.instr(items, result), result, rpr):
            ins.append(el)
        place_after(anchor, ins, self.tr)
        self.count += 1
        return True

    def parse_group(self, g):
        parts = [x.strip() for x in g.split(";")]
        out = []
        for part in parts:
            sub = [part] if self.part_re.match(part) else [x.strip() for x in re.split(r",\s+(?=[A-Z])", part)]
            for x in sub:
                m = self.part_re.match(x) if self.part_re else None
                if not m:
                    return None
                out.append((m.group("prefix"), m.group("key")))
        return out

    def edit_text(self, p, s, e, new):
        runs = span_runs(p, s, e)
        if not runs:
            return False
        rpr = runs[0].find(w("rPr"))
        anchor = delete_runs(runs, self.tr)
        ins = self.tr.mark("ins"); ins.append(text_run(new, rpr))
        place_after(anchor, ins, self.tr)
        return True

    def run_paragraph(self, p):
        # typed-text edits that a replacement depends on
        for ed in self.m.get("edits", []):
            spans, text = model(p)
            i = text.find(ed["find"])
            if i >= 0 and self.edit_text(p, i, i + len(ed["find"]), ed["new"]):
                ed["done"] = ed.get("done", 0) + 1
        # explicit replacements (they may contain parentheses)
        for rep in self.m.get("replace", []):
            spans, text = model(p)
            i = text.find(rep["find"])
            if i < 0:
                continue
            key = rep["key"]
            if rep.get("mode") == "narrative":
                y = self.csl[key].get("issued", {}).get("date-parts", [[None]])[0][0]
                result = f"({y or 'n.d.'})"
                items = [self.item(key, suppress=True)]
            else:
                result = rep.get("result") or f"({rep.get('name', '')})"
                items = [self.item(key, locator=rep.get("locator"), label=rep.get("label"))]
            if self.replace_span(p, i, i + len(rep["find"]), items, result, rep.get("lead", "")):
                rep["done"] = rep.get("done", 0) + 1
        # parenthetical groups
        start = 0
        while True:
            spans, text = model(p)
            m = re.compile(r"\(([^()]+)\)").search(text, start)
            if not m:
                break
            start = m.end()
            if not re.search(r"\b(?:19|20)\d{2}|n\.d\.", m.group(1)):
                continue
            parts = self.parse_group(m.group(1)) if self.part_re else None
            if not parts:
                before = text[max(0, m.start() - 40):m.start()]
                nk = self.narrative_key(before, m.group(1))
                if not nk:
                    self.log.append(f"  left as typed  ({m.group(1)})   ...{before[-30:]}")
                continue
            items, shown = [], []
            for prefix, key in parts:
                name, year = key.rsplit(" ", 1)
                items.append(self.item(self.m["cites"][key], prefix=prefix.strip() or None))
                shown.append(f"{prefix}{name}, {year}")
            if self.replace_span(p, m.start(), m.end(), items, "(" + "; ".join(shown) + ")"):
                start = 0
        # narrative: Name (year)
        start = 0
        while True:
            spans, text = model(p)
            m = re.compile(r"\(((?:19|20)\d{2}[a-z]?(?:-\d{4})?|n\.d\.)\)").search(text, start)
            if not m:
                break
            start = m.end()
            key = self.narrative_key(text[max(0, m.start() - 60):m.start()], m.group(1))
            if key and self.replace_span(p, m.start(), m.end(), [self.item(key, suppress=True)], f"({m.group(1)})"):
                start = 0

    def narrative_key(self, before, year):
        for k, key in self.m.get("narrative", {}).items():
            name, y = k.split("|")
            if y == year and before.rstrip().endswith(name) and before.endswith(" "):
                return key
        return None


# ---------------------------------------------------------------- footnotes -> citations
def footnotes_to_citations(body, fmap, linker, log):
    for p in body.iter(w("p")):
        for r in list(R.runs_of(p)):
            ref = r.find(w("footnoteReference"))
            if ref is None or ref.get(w("id")) not in fmap:
                continue
            key = fmap[ref.get(w("id"))]
            data = linker.csl[key]
            name = (data.get("author") or [{}])[0]
            name = name.get("literal") or name.get("family") or data.get("title", "")
            if len(data.get("author") or []) > 2:
                name += " et al."
            year = (data.get("issued", {}).get("date-parts") or [[None]])[0][0] or "n.d."
            result = f"({name}, {year})"
            # position of the mark in the visible text; step back over a period/comma before it
            spans, text = model(p)
            visible = {x: b for x, a, b in spans}
            pos = 0
            for x in R.runs_of(p):
                if x is r:
                    break
                pos = visible.get(x, pos)
            if pos and text[pos - 1] in ".,;":
                pos -= 1
            split_visible(p, pos)
            after = [x for x, a, b in model(p)[0] if a == pos]
            rpr = r.find(w("rPr"))
            rpr = copy.deepcopy(rpr) if rpr is not None else None
            if rpr is not None:
                for v in rpr.findall(w("vertAlign")) + rpr.findall(w("rStyle")):
                    rpr.remove(v)
            ins = linker.tr.mark("ins")
            ins.append(text_run(" ", rpr))
            for el in field_runs(linker.instr([linker.item(key)], result), result, rpr):
                ins.append(el)
            target = after[0] if after else r
            if target.getparent().tag == w("ins"):
                target = target.getparent()
            target.addprevious(ins)
            d = linker.tr.mark("del"); r.addprevious(d); d.append(r)
            linker.count += 1
            log.append(f"  footnote {ref.get(w('id'))} -> {result}")


# ---------------------------------------------------------------- bibliography
def bib_paragraphs(instr, entries, ppr_template, tr):
    paras = []
    for i, text in enumerate(entries):
        p = etree.Element(w("p"))
        ppr = copy.deepcopy(ppr_template) if ppr_template is not None else etree.Element(w("pPr"))
        for old in ppr.findall(w("rPr")):
            ppr.remove(old)
        rpr = etree.SubElement(ppr, w("rPr")); rpr.append(tr.mark("ins"))
        p.append(ppr)
        ins = tr.mark("ins")
        if i == 0:
            for el in [R.field_run("begin", None), R.field_run(None, None, instr), R.field_run("separate", None)]:
                ins.append(el)
        ins.append(text_run(text, None))
        if i == len(entries) - 1:
            ins.append(R.field_run("end", None))
        p.append(ins)
        paras.append(p)
    return paras


def style_ppr(style):
    ppr = etree.Element(w("pPr")); s = etree.SubElement(ppr, w("pStyle")); s.set(w("val"), style)
    return ppr


def delete_paragraph(p, tr):
    for r in R.runs_of(p):
        anc, deleted = r.getparent(), False
        while anc is not p:
            deleted |= anc.tag == w("del"); anc = anc.getparent()
        if not deleted:
            d = tr.mark("del"); r.addprevious(d); to_deleted(r); d.append(r)
    ppr = p.find(w("pPr"))
    if ppr is None:
        ppr = etree.Element(w("pPr")); p.insert(0, ppr)
    rpr = ppr.find(w("rPr"))
    if rpr is None:
        rpr = etree.SubElement(ppr, w("rPr"))
    rpr.append(tr.mark("del"))


def do_bibliography(body, spec, linker, bib, tr, log):
    keys = sorted(linker.cited | set(spec.get("uncited", [])))
    entries = sorted((bib[k] for k in keys), key=lambda s: s.lower())
    uncited = [[f"http://zotero.org/{linker.lib}/items/{k}"] for k in spec.get("uncited", []) if k not in linker.cited]
    instr = " ADDIN ZOTERO_BIBL " + json.dumps({"uncited": uncited, "omitted": [], "custom": []}) + " CSL_BIBLIOGRAPHY "
    if spec["mode"] == "replace":
        refs = [p for p in body.iter(w("p")) if (p.find(f"{w('pPr')}/{w('pStyle')}") is not None and
                p.find(f"{w('pPr')}/{w('pStyle')}").get(w("val")) == "References")]
        refs = [p for p in refs if ptext(p).strip() and
                not any(ptext(p).strip().startswith(k) for k in spec.get("keep", []))]
        template = refs[0].find(w("pPr"))
        for p in refs:
            delete_paragraph(p, tr)
        anchor = refs[-1]
        new = bib_paragraphs(instr, entries, template, tr)
    else:
        anchor = next(p for p in body.iter(w("p")) if ptext(p).strip().startswith(spec["after"]))
        head = etree.Element(w("p")); ppr = style_ppr(spec["heading_style"])
        rpr = etree.SubElement(ppr, w("rPr")); rpr.append(tr.mark("ins")); head.append(ppr)
        ins = tr.mark("ins"); ins.append(text_run(spec["heading"], None)); head.append(ins)
        new = [head] + bib_paragraphs(instr, entries, style_ppr(spec.get("entry_style", "References")), tr)
    for p in new:
        anchor.addnext(p); anchor = p
    log.append(f"  bibliography: {len(entries)} entries ({len(linker.cited)} cited, {len(uncited)} uncited), mode {spec['mode']}")


# ---------------------------------------------------------------- people.xml
def copy_person(parts, src, author):
    with zipfile.ZipFile(src) as z:
        if "word/people.xml" not in z.namelist():
            return "no people.xml in source"
        people = etree.fromstring(z.read("word/people.xml"))
    person = next((p for p in people if p.get(f"{{{W15}}}author") == author), None)
    if person is None:
        return f"{author} not in source people.xml"
    if "word/people.xml" not in parts:
        return "target has no people.xml (Word adds it on save)"
    tgt = etree.fromstring(parts["word/people.xml"])
    if any(p.get(f"{{{W15}}}author") == author for p in tgt):
        return "already present"
    tgt.append(copy.deepcopy(person))
    parts["word/people.xml"] = etree.tostring(tgt, xml_declaration=True, encoding="UTF-8", standalone=True)
    return "added"


# ---------------------------------------------------------------- main
def max_id(parts):
    ids = [0]
    for name, data in parts.items():
        if name.startswith("word/") and name.endswith(".xml"):
            ids += [int(x) for x in re.findall(rb'w:id="(-?\d+)"', data)]
    return max(ids)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx"); ap.add_argument("map")
    ap.add_argument("--author", required=True)
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    path = Path(a.docx)
    m = json.load(open(a.map, encoding="utf-8"))
    lib = m["library"]
    keys = set(m.get("cites", {}).values()) | set(m.get("narrative", {}).values()) | {r["key"] for r in m.get("replace", [])} \
        | set(m.get("footnotes", {}).values()) | set(m.get("bibliography", {}).get("uncited", []))
    csl, bib = fetch_items(lib, keys)

    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        parts = {i.filename: z.read(i.filename) for i in infos}
    tr = Track(a.author, max_id(parts) + 1000)
    log = []
    linker = Linker(m, csl, lib, tr, log)
    roots = {}
    for part in ("word/document.xml", "word/footnotes.xml"):
        if part in parts:
            roots[part] = etree.fromstring(parts[part])
    body = roots["word/document.xml"].find(w("body"))
    if m.get("footnotes"):
        footnotes_to_citations(body, m["footnotes"], linker, log)
    for part, root in roots.items():
        for p in list(root.iter(w("p"))):
            st = p.find(f"{w('pPr')}/{w('pStyle')}")
            if st is not None and st.get(w("val")) == "References":
                continue
            linker.run_paragraph(p)
    for rep in m.get("replace", []) + m.get("edits", []):
        if not rep.get("done"):
            log.append(f"  replace NOT FOUND: {rep['find'][:60]}")
    if m.get("bibliography"):
        do_bibliography(body, m["bibliography"], linker, bib, tr, log)
    print(f"{path.name}: {linker.count} citations linked to {len(linker.cited)} items, author {a.author}")
    for line in log:
        print(line)
    if not a.write:
        return
    if R.locked(path):
        raise SystemExit("open in Word (~$ lock file); not written")
    for part, root in roots.items():
        parts[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if m.get("prefs_from"):
        with zipfile.ZipFile(m["prefs_from"]) as z:
            prefs = R.zotero_prefs(z)
        for p in prefs:
            v = p.find(f"{{{R.VT}}}lpwstr")
            if v is not None and v.text:
                v.text = re.sub(r'<session id="[^"]*"', f'<session id="{rand_id()}"', v.text)
        R.merge_prefs(parts, prefs)
    if m.get("person_from"):
        print("  people.xml:", copy_person(parts, m["person_from"], a.author))
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        names = [i.filename for i in infos] + [n for n in parts if n not in {i.filename for i in infos}]
        for n in names:
            out.writestr(n, parts[n])
    tmp.replace(path)
    print("  written")


if __name__ == "__main__":
    sys.exit(main())
