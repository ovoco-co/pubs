"""Restore Zotero citation and bibliography fields from author originals into edited masters.

Usage:
    python restore_zotero_fields.py ORIG_DIR MASTER_DIR [--write] [--only 02,06]

A chapter that went through the NIST template pipeline keeps its citation TEXT but loses the
Zotero fields around it, so Zotero no longer sees the document. The original still holds every
field, with the full item data and library URI in its instruction. This copies them back:

  1. Citations - each ZOTERO_ITEM field in the original, in document order, is matched to the same
     text in the master, searching forward from the previous match and confirmed by the words
     around it. The master's runs are wrapped in the original instruction, unchanged. Text,
     styles and tracked changes in the master are not altered; only field markers are added.
  2. Bibliography - the master's reference list (first to last entry of the original's
     ZOTERO_BIBL result) is wrapped in the original ZOTERO_BIBL instruction, so the NIST-styled
     paragraphs stay as the field result.
  3. Document preferences - the ZOTERO_PREF_* custom properties (style, field type) are copied,
     without which Zotero treats the document as unlinked.

Anything that cannot be placed safely is reported and left alone: a citation whose sentence was
rewritten, one that spans a tracked-change boundary, a reference entry edited in the master. A
citation already inside a live Zotero field in the master counts as linked and is not touched.

Dry run by default: reports what would happen. --write saves the masters in place (skipping any
with a ~$ lock file). Afterwards open each chapter with Zotero running and click Refresh.
"""
import argparse
import copy
import difflib
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
PREF_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
CONTEXT = 60          # characters of surrounding text compared to confirm a match
MIN_SCORE = 0.6       # context similarity needed to accept a match
WINDOW = 250          # paragraphs searched ahead of the previous match


def w(tag):
    return f"{{{W}}}{tag}"


# ---------------------------------------------------------------- reading the run stream
def para_of(el):
    p = el.getparent()
    while p is not None and p.tag != w("p"):
        p = p.getparent()
    return p


def runs_of(p):
    """Runs belonging to this paragraph (not to a nested text box paragraph), in order."""
    return [r for r in p.iter(w("r")) if para_of(r) is p]


def run_text(r):
    return "".join(t.text or "" for t in r.findall(w("t")))


def walk_fields(body):
    """Map each run to the instruction of the field whose RESULT it sits in (None if none).
    Complex fields can span paragraphs (a bibliography does), so this walks the whole body."""
    owner, stack, collecting = {}, [], None
    for r in body.iter(w("r")):
        fc = r.find(w("fldChar"))
        if fc is not None:
            kind = fc.get(w("fldCharType"))
            if kind == "begin":
                stack.append([""]); collecting = stack[-1]
            elif kind == "separate":
                collecting = None
            elif kind == "end" and stack:
                stack.pop(); collecting = None
            continue
        it = r.find(w("instrText"))
        if it is not None and collecting is not None:
            collecting[0] += it.text or ""
            continue
        owner[r] = stack[-1][0] if stack else None
    return owner


def paragraphs(body):
    out = []
    for p in body.iter(w("p")):
        rs = runs_of(p)
        out.append((p, rs, "".join(run_text(r) for r in rs)))
    return out


# ---------------------------------------------------------------- the original's fields
def zotero_fields(body):
    """Each Zotero field: instruction, result text, and the paragraph text around it."""
    fields, cur = [], None
    for p in body.iter(w("p")):
        before = ""
        for r in runs_of(p):
            fc = r.find(w("fldChar"))
            if fc is not None:
                kind = fc.get(w("fldCharType"))
                if kind == "begin" and cur is None:
                    cur = {"instr": "", "result": "", "state": "instr", "depth": 1, "before": before, "p": p}
                elif cur is not None:
                    if kind == "begin":
                        cur["depth"] += 1
                    elif kind == "separate" and cur["depth"] == 1:
                        cur["state"] = "result"
                    elif kind == "end":
                        cur["depth"] -= 1
                        if cur["depth"] == 0:
                            cur["after_p"] = p
                            cur["after_start"] = len(before)
                            fields.append(cur); cur = None
                continue
            it = r.find(w("instrText"))
            if cur is not None and it is not None and cur["state"] == "instr" and cur["depth"] == 1:
                cur["instr"] += it.text or ""
            t = run_text(r)
            if cur is not None and cur["state"] == "result":
                cur["result"] += t
            before += t
        if cur is not None and cur["state"] == "result":
            cur["result"] += "\n"
    out = []
    for f in fields:
        kind = "BIBL" if "ZOTERO_BIBL" in f["instr"] else "ITEM" if "ZOTERO_ITEM" in f["instr"] else None
        if not kind:
            continue
        ptext = "".join(run_text(r) for r in runs_of(f["after_p"]))
        f["after"] = ptext[f["after_start"]:][:CONTEXT // 2] if f["after_p"] is f["p"] else ""
        f["before"] = f["before"][-CONTEXT:]
        f["kind"] = kind
        out.append(f)
    return out


# ---------------------------------------------------------------- editing the master
def explode(r):
    """Split a run into one run per content child, each carrying the run properties."""
    rpr = r.find(w("rPr"))
    kids = [k for k in r if k.tag != w("rPr")]
    if len(kids) <= 1:
        return [r]
    parent, idx, new = r.getparent(), r.getparent().index(r), []
    for k in kids:
        nr = etree.Element(w("r"))
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        nr.append(k)
        new.append(nr)
    parent.remove(r)
    for i, nr in enumerate(new):
        parent.insert(idx + i, nr)
    return new


def split_at(p, offset):
    """Make `offset` (in the paragraph's run text) fall on a run boundary."""
    pos = 0
    for r in runs_of(p):
        n = len(run_text(r))
        if pos < offset < pos + n:
            pieces = explode(r)
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
        pos += n


def runs_in(p, start, end):
    out, pos = [], 0
    for r in runs_of(p):
        n = len(run_text(r))
        if n and pos >= start and pos + n <= end:
            out.append(r)
        pos += n
    return out


def field_run(kind, rpr, instr=None):
    r = etree.Element(w("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    if instr is None:
        fc = etree.SubElement(r, w("fldChar")); fc.set(w("fldCharType"), kind)
    else:
        it = etree.SubElement(r, w("instrText")); it.text = instr; it.set(XML_SPACE, "preserve")
    return r


def wrap(first, last, instr):
    rpr = first.find(w("rPr"))
    first.addprevious(field_run("begin", rpr))
    first.addprevious(field_run(None, rpr, instr))
    first.addprevious(field_run("separate", rpr))
    last.addnext(field_run("end", rpr))


def score(a, b):
    if not a and not b:
        return 1.0
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def deleted_ids(body):
    """citationIDs of Zotero fields that sit inside a tracked deletion in the master."""
    text = "".join(t.text or "" for t in body.iter(w("delInstrText")))
    return set(re.findall(r'"citationID":"([^"]+)"', text))


def place_citations(fields, body, log):
    owner = walk_fields(body)
    gone = deleted_ids(body)
    paras = paragraphs(body)
    index = {p: i for i, (p, _, _) in enumerate(paras)}
    cur_p, cur_off = 0, 0
    done = linked = 0
    for f in fields:
        text = f["result"].strip()
        best = None
        for pi in range(cur_p, min(len(paras), cur_p + WINDOW)):
            p, rs, pt = paras[pi]
            start = cur_off if pi == cur_p else 0
            for m in re.finditer(re.escape(text), pt[start:]):
                s, e = m.start() + start, m.end() + start
                sc = (score(f["before"], pt[max(0, s - CONTEXT):s]) * 2 + score(f["after"], pt[e:e + CONTEXT // 2])) / 3
                if sc >= MIN_SCORE:
                    best = (pi, s, e, sc); break
            if best:
                break
        if not best:
            cid = re.search(r'"citationID":"([^"]+)"', f["instr"])
            why = "deleted in the master as a tracked change (field intact inside the deletion)"                 if cid and cid.group(1) in gone else f"...{f['before'][-40:]}[here]"
            log.append(f"  NOT PLACED  {text}   {why}")
            continue
        pi, s, e, sc = best
        p, rs, pt = paras[pi]
        spanned = runs_in(p, s, e) if s >= 0 else []
        owners = {owner.get(r) for r in spanned if r in owner}
        if any(o and "ZOTERO_ITEM" in o for o in owners):
            linked += 1
        elif any(owners):
            log.append(f"  NOT PLACED  {text}   inside another field ({next(o for o in owners if o).split()[0]})")
        else:
            split_at(p, s); split_at(p, e)
            spanned = runs_in(p, s, e)
            if not spanned or len({r.getparent() for r in spanned}) != 1:
                log.append(f"  NOT PLACED  {text}   crosses a tracked change or hyperlink   ...{f['before'][-40:]}[here]")
            else:
                wrap(spanned[0], spanned[-1], f["instr"])
                done += 1
                if sc < 0.9:
                    log.append(f"  placed, context {sc:.2f}  {text}   ...{pt[max(0, s - 40):s]}[here]")
        cur_p, cur_off = pi, e
        paras[pi] = (p, runs_of(p), "".join(run_text(r) for r in runs_of(p)))
    return done, linked


def place_bibliography(f, body, log):
    entries = [x.strip() for x in f["result"].split("\n") if x.strip()]
    paras = paragraphs(body)
    texts = [pt.strip() for _, _, pt in paras]

    def find(entry, start=0):
        for i in range(start, len(texts)):
            if texts[i] == entry:
                return i
        cands = [(score(entry[:120], t[:120]), i) for i, t in enumerate(texts) if i >= start and t]
        sc, i = max(cands, default=(0, -1))
        return i if sc >= 0.9 else -1

    a = find(entries[0])
    b = find(entries[-1], max(a, 0))
    if a < 0 or b < 0:
        log.append("  BIBLIOGRAPHY NOT PLACED: first or last reference entry not found in the master")
        return False
    owner = walk_fields(body)
    if any(owner.get(r) and "ZOTERO_BIBL" in owner[r] for r in paras[a][1]):
        log.append("  bibliography already a Zotero field")
        return True
    same = sum(1 for x in entries if x in set(texts[a:b + 1]))
    first_p, last_p = paras[a][0], paras[b][0]
    rs = runs_of(first_p)
    rpr = rs[0].find(w("rPr")) if rs else None
    ppr = first_p.find(w("pPr"))
    at = 1 if ppr is not None else 0
    for k, el in enumerate([field_run("begin", rpr), field_run(None, rpr, f["instr"]), field_run("separate", rpr)]):
        first_p.insert(at + k, el)
    last_p.append(field_run("end", rpr))
    log.append(f"  bibliography wrapped: {b - a + 1} paragraphs, {same}/{len(entries)} entries identical to the original")
    for x in entries:
        if x not in set(texts[a:b + 1]):
            log.append(f"    edited in master (Refresh will overwrite): {x[:90]}")
    return True


# ---------------------------------------------------------------- custom properties
def zotero_prefs(z):
    if "docProps/custom.xml" not in z.namelist():
        return []
    root = etree.fromstring(z.read("docProps/custom.xml"))
    return [p for p in root.findall(f"{{{CP}}}property") if p.get("name", "").startswith("ZOTERO_PREF")]


def merge_prefs(parts, prefs):
    if "docProps/custom.xml" in parts:
        root = etree.fromstring(parts["docProps/custom.xml"])
    else:
        root = etree.Element(f"{{{CP}}}Properties", nsmap={None: CP, "vt": VT})
        ct = parts["[Content_Types].xml"].decode("utf8")
        if "/docProps/custom.xml" not in ct:
            ct = ct.replace("</Types>", f'<Override PartName="/docProps/custom.xml" ContentType="{CUSTOM_CT}"/></Types>')
            parts["[Content_Types].xml"] = ct.encode("utf8")
        rels = parts["_rels/.rels"].decode("utf8")
        if CUSTOM_REL not in rels:
            rels = rels.replace("</Relationships>",
                                f'<Relationship Id="rIdZoteroCustom" Type="{CUSTOM_REL}" Target="docProps/custom.xml"/></Relationships>')
            parts["_rels/.rels"] = rels.encode("utf8")
    for p in root.findall(f"{{{CP}}}property"):
        if p.get("name", "").startswith("ZOTERO_PREF"):
            root.remove(p)
    pid = max([int(p.get("pid")) for p in root.findall(f"{{{CP}}}property")] + [1])
    for p in prefs:
        pid += 1
        q = copy.deepcopy(p); q.set("pid", str(pid)); q.set("fmtid", PREF_FMTID)
        root.append(q)
    parts["docProps/custom.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------- per chapter
def locked(path):
    d = path.parent
    return any((d / n).exists() for n in ("~$" + path.name, "~$" + path.name[2:]))


def restore(orig, master, write):
    log = []
    zo = zipfile.ZipFile(orig)
    obody = etree.fromstring(zo.read("word/document.xml")).find(w("body"))
    fields = zotero_fields(obody)
    prefs = zotero_prefs(zo)
    cites = [f for f in fields if f["kind"] == "ITEM"]
    bibl = [f for f in fields if f["kind"] == "BIBL"]
    print(f"{master.name}: original has {len(cites)} citations, {len(bibl)} bibliography, {len(prefs)} pref properties")
    if not fields and not prefs:
        return
    with zipfile.ZipFile(master) as zm:
        parts = {i.filename: zm.read(i.filename) for i in zm.infolist()}
        infos = zm.infolist()
    root = etree.fromstring(parts["word/document.xml"])
    body = root.find(w("body"))
    done, linked = place_citations(cites, body, log)
    print(f"  citations: {done} restored, {linked} already linked, {len(cites) - done - linked} not placed")
    for f in bibl:
        place_bibliography(f, body, log)
    for line in log:
        print(line)
    if prefs:
        merge_prefs(parts, prefs)
    if not write:
        return
    if locked(master):
        print("  SKIPPED: open in Word (~$ lock file)")
        return
    parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = master.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        names = [i.filename for i in infos] + [n for n in parts if n not in {i.filename for i in infos}]
        for n in names:
            out.writestr(n, parts[n])
    tmp.replace(master)
    print("  written")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("orig_dir"); ap.add_argument("master_dir")
    ap.add_argument("--write", action="store_true", help="save the masters (default: report only)")
    ap.add_argument("--only", help="comma-separated chapter numbers, e.g. 02,06")
    a = ap.parse_args()
    only = set(a.only.split(",")) if a.only else None
    for orig in sorted(Path(a.orig_dir).glob("*.docx")):
        if orig.name.startswith("~$"):
            continue
        m = re.search(r"Chapter-(\d+)", orig.name)
        if only and (not m or m.group(1) not in only):
            continue
        master = Path(a.master_dir) / orig.name
        if not master.exists():
            print(f"{orig.name}: no master of that name")
            continue
        restore(orig, master, a.write)


if __name__ == "__main__":
    sys.exit(main())
