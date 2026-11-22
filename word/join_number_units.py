"""Join a number to its unit with a non-breaking space, as the NIST Guide to the SI asks.

Usage:
    python join_number_units.py DOCX [DOCX ...] --author "Name" [--write]

"95.2 %", "500 ft", "35 m/s", "230 kV" must not break across a line. Two cases:

  - An ordinary space between a number and a unit becomes U+00A0. The page looks the same, so
    this is not tracked.
  - No space at all ("20%", "35m/s") gets a U+00A0 inserted. That is a visible edit, so it is a
    tracked insertion under --author.

Only digits count as the number (so "five km" and "Section 4 m" are not touched), the unit
must end at a word boundary, and the text must be visible: field codes, field results
(citations, cross-references) and deleted text are skipped. Covers the body, footnotes and
endnotes. "in" is not in the unit list: "20 in Ponce" is far commoner than inches.

Dry run by default: lists every change by chapter. --write saves in place, refusing a file open
in Word. Pairs with Repair-NonBreakingSpaces.ps1, which keeps these joins and clears the
accidental ones.
"""
import argparse
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

import link_zotero_citations as L
import restore_zotero_fields as R

w = R.w
NB = " "
UNITS = ["%", "km/h", "m/s", "mi/h", "km²", "m²", "ft²", "km2", "m2", "°C", "°F", "kWh", "MWh", "GWh", "kVA",
         "MVA", "km", "mm", "cm", "m", "ft", "mi", "mph", "kph", "kV", "V", "GW", "MW", "kW", "W", "Hz", "kg",
         "lb", "ha", "ms", "min", "s", "h"]
# number, optional ordinary space, unit; the unit must not run on into a word or another number
UNIT_RX = re.compile(r"(?<![\w.,])(\d[\d,]*(?:\.\d+)?)( ?)(" + "|".join(re.escape(u) for u in UNITS) + r")(?![\w/²])")


def process(root, tr, log, label):
    joined = inserted = 0
    for p in list(root.iter(w("p"))):
        spans, text = L.model(p)
        # right to left, so earlier offsets stay valid as the paragraph changes
        for m in reversed(list(UNIT_RX.finditer(text))):
            num, space, unit = m.groups()
            ctx = text[max(0, m.start() - 30):m.end() + 10].replace("\n", " ")
            # inserting a space changes the words; before a capitalized word it is usually a
            # title or product name ("Transitions to 100% Renewable Energy", "Global 500m SIN
            # Grid"), published wording to leave as is. Making an existing space non-breaking
            # changes nothing visible, so that is done everywhere.
            if not space and re.match(r"\s+[A-Z]", text[m.end():]):
                log.append(f"  {label}  LEFT     ...{ctx}   (title or name?)")
                continue
            if space:                                   # ordinary space -> NBSP, untracked
                pos = m.start(2)
                for r, a, b in L.model(p)[0]:
                    if a <= pos < b:
                        t = r.find(w("t")); k = pos - a
                        t.text = t.text[:k] + NB + t.text[k + 1:]
                        t.set(R.XML_SPACE, "preserve")
                        joined += 1
                        log.append(f"  {label}  join     ...{ctx}")
                        break
            else:                                       # no space -> tracked NBSP insertion
                pos = m.start(3)
                L.split_visible(p, pos)
                anchor = next((r for r, a, b in L.model(p)[0] if b == pos and b > a), None)
                if anchor is None:
                    log.append(f"  {label}  SKIPPED  ...{ctx}")
                    continue
                ins = tr.mark("ins")
                ins.append(L.text_run(NB, anchor.find(w("rPr"))))
                L.place_after(anchor, ins, tr)
                inserted += 1
                log.append(f"  {label}  INSERT   ...{ctx}")
    return joined, inserted


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", nargs="+")
    ap.add_argument("--author", required=True)
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--quiet", action="store_true", help="totals only")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    for f in a.docx:
        path = Path(f)
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            parts = {i.filename: z.read(i.filename) for i in infos}
        tr = L.Track(a.author, L.max_id(parts) + 1000)
        log, J, I = [], 0, 0
        for name, label in (("word/document.xml", "body"), ("word/footnotes.xml", "note"), ("word/endnotes.xml", "endn")):
            if name not in parts:
                continue
            root = etree.fromstring(parts[name])
            j, i = process(root, tr, log, label)
            if j or i:
                parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            J += j; I += i
        print(f"{path.name}: {J} spaces made non-breaking, {I} non-breaking spaces inserted (tracked)")
        if not a.quiet:
            for line in log:
                print(line)
        if a.write and (J or I):
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
