"""Make an empty copy of a .docx: keep styles, numbering, theme, settings,
fonts, glossary and the FINAL section's page setup and headers/footers; drop
all body content and every part only the body used.

Usage: python make_empty_template.py <source.docx> <dest.docx>
"""
import posixpath
import re
import sys
import zipfile

if len(sys.argv) != 3:
    sys.exit(__doc__)
src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}

# ---- document.xml: one empty Body Text paragraph + the final sectPr ----
doc = parts['word/document.xml'].decode('utf-8')
body_open = re.search(r'<w:body>', doc)
final_sect = re.search(r'(?s)<w:sectPr\b(?:(?!<w:sectPr\b).)*?</w:sectPr>\s*</w:body>', doc)
sect = re.sub(r'\s*</w:body>$', '', final_sect.group(0))
new_body = '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p>' + sect
doc = doc[:body_open.end()] + new_body + '</w:body></w:document>'
parts['word/document.xml'] = doc.encode('utf-8')

# ---- footnotes / endnotes: keep only the separators ----
for name, tag in (('word/footnotes.xml', 'footnote'), ('word/endnotes.xml', 'endnote')):
    if name in parts:
        x = parts[name].decode('utf-8')
        x = re.sub(r'(?s)<w:%s\b(?![^>]*w:type="(?:separator|continuationSeparator|continuationNotice)")[^>]*>.*?</w:%s>' % (tag, tag), '', x)
        parts[name] = x.encode('utf-8')

# ---- document.xml.rels: keep package-level parts and what the body still uses ----
KEEP_TYPES = ('/styles', '/numbering', '/settings', '/webSettings', '/fontTable', '/theme',
              '/footnotes', '/endnotes', '/customXml', '/glossaryDocument', '/people',
              '/stylesWithEffects', '/commentsExtensible')
rels = parts['word/_rels/document.xml.rels'].decode('utf-8')
used_ids = set(re.findall(r'r:id="([^"]+)"', doc)) | set(re.findall(r'r:embed="([^"]+)"', doc))
kept = []
for m in re.finditer(r'<Relationship\b[^>]*/>', rels):
    r = m.group(0)
    rid = re.search(r'Id="([^"]+)"', r).group(1)
    typ = re.search(r'Type="([^"]+)"', r).group(1)
    if rid in used_ids or typ.endswith(KEEP_TYPES):
        kept.append(r)
head = rels[:rels.index('<Relationship ')]
parts['word/_rels/document.xml.rels'] = (head + ''.join(kept) + '</Relationships>').encode('utf-8')

# ---- drop every part no relationship reaches any more ----
def reachable():
    seen = set()
    todo = ['_rels/.rels']
    while todo:
        relname = todo.pop()
        if relname not in parts:
            continue
        base = posixpath.dirname(posixpath.dirname(relname))  # folder of the source part
        for m in re.finditer(r'<Relationship\b[^>]*/>', parts[relname].decode('utf-8')):
            r = m.group(0)
            if 'TargetMode="External"' in r:
                continue
            tgt = re.search(r'Target="([^"]+)"', r).group(1)
            path = tgt.lstrip('/') if tgt.startswith('/') else posixpath.normpath(posixpath.join(base, tgt))
            if path in parts and path not in seen:
                seen.add(path)
                rp = posixpath.join(posixpath.dirname(path), '_rels', posixpath.basename(path) + '.rels')
                if rp in parts:
                    seen.add(rp)
                    todo.append(rp)
    return seen

live = reachable() | {'[Content_Types].xml', '_rels/.rels'}
dropped = sorted(p for p in parts if p not in live)
for p in dropped:
    del parts[p]

# ---- [Content_Types].xml: no Override for a part that is gone ----
ct = parts['[Content_Types].xml'].decode('utf-8')
ct = re.sub(r'<Override PartName="/([^"]+)"[^>]*/>', lambda m: m.group(0) if m.group(1) in parts else '', ct)
parts['[Content_Types].xml'] = ct.encode('utf-8')

with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
    for i in zin.infolist():
        if i.filename in parts:
            zout.writestr(i, parts[i.filename])

print('dropped %d parts:' % len(dropped))
for p in dropped:
    print('   ', p)
