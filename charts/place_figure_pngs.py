"""Swap figure pictures in a Word document for rebuilt PNGs, located by caption.

Usage:
    python place_figure_pngs.py DOCX PNG_DIR FIG=GFX [FIG=GFX ...]
    e.g. python place_figure_pngs.py Chapter-06.docx ...\\png 6-2=1370 6-3=1371

For each FIG ("6-2"), finds the CaptionFigure paragraph "Figure 6-2." and the single picture in the
nearest picture paragraph above it, then points that picture at PNG_DIR\\<GFX>a-NWIRP.png. The picture
is placed at its natural size (PNG pixels / 300 dpi, capped at the text-column width), uncropped,
so a chart built at its placed size keeps its 8 pt text. Nothing else in the document changes.
The old image part is removed when no other relationship uses it. Writes the document in place.
"""
import re
import sys
import zipfile
from pathlib import Path

from PIL import Image

EMU_PER_TWIP = 635
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def fit_picture(xml, cx, cy):
    """Size a picture drawing to cx x cy, uncropped and undistorted.

    Clears any a:srcRect crop the original picture carried (a crop left on a new image trims it and
    stretches what is left), and sets both the frame (wp:extent) and the picture's own transform
    (a:ext inside pic:spPr, which may carry attributes such as bwMode) to the new size."""
    xml = re.sub(r"<a:srcRect\b[^>]*/>", "<a:srcRect/>", xml)
    xml = re.sub(r"<a:srcRect\b[^>]*>.*?</a:srcRect>", "<a:srcRect/>", xml, flags=re.S)
    xml = re.sub(r'<wp:extent cx="\d+" cy="\d+"/>', f'<wp:extent cx="{cx}" cy="{cy}"/>', xml)
    xml = re.sub(r'(<pic:spPr\b[^>]*>.*?<a:ext )cx="\d+" cy="\d+"', rf'\1cx="{cx}" cy="{cy}"', xml, flags=re.S)
    return xml


def main(docx, png_dir, pairs):
    docx = Path(docx)
    z = zipfile.ZipFile(docx)
    parts = {n: z.read(n) for n in z.namelist()}
    infos = {i.filename: i for i in z.infolist()}
    z.close()

    doc = parts["word/document.xml"].decode("utf-8")
    rels = parts["word/_rels/document.xml.rels"].decode("utf-8")
    ct = parts["[Content_Types].xml"].decode("utf-8")

    pg = re.search(r'<w:pgSz w:w="(\d+)"', doc)
    mar = re.search(r'<w:pgMar [^>]*w:right="(\d+)"[^>]*w:left="(\d+)"', doc)
    text_w = (int(pg.group(1)) - int(mar.group(1)) - int(mar.group(2))) * EMU_PER_TWIP

    paras = [(m.start(), m.end()) for m in re.finditer(r"<w:p[ >].*?</w:p>", doc, re.S)]
    report = []
    for fig, gfx in pairs:
        png = Path(png_dir) / f"{gfx}a-NWIRP.png"
        if not png.exists():
            raise SystemExit(f"missing {png}")
        num = fig.replace("-", r"\D")
        cap_i = None
        for i, (a, b) in enumerate(paras):
            p = doc[a:b]
            text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)).replace("\xa0", " ")
            if "CaptionFigure" in p and re.match(rf"\s*Figure\s*{num}\.", text):
                cap_i = i
                break
        if cap_i is None:
            raise SystemExit(f"caption for Figure {fig} not found")
        # the figure's own picture: search upward, but never past the previous figure caption
        pic_i = None
        for i in range(cap_i - 1, -1, -1):
            q = doc[paras[i][0]:paras[i][1]]
            if "CaptionFigure" in q and re.search(r"<w:t[^>]*>\s*Figure", q):
                break
            if "r:embed=" in q or "drawingml/2006/chart" in q:
                pic_i = i
                break
        if pic_i is None:
            raise SystemExit(f"Figure {fig}: no picture or chart between its caption and the previous one")
        a, b = paras[pic_i]
        p = doc[a:b]
        embeds = re.findall(r'r:embed="([^"]+)"', p)        # any id, including rIdGfxN from an earlier run
        is_chart = not embeds and "drawingml/2006/chart" in p
        if not is_chart and len(embeds) != 1:
            raise SystemExit(f"Figure {fig}: expected one picture above the caption, found {len(embeds)}")
        old = embeds[0] if embeds else None

        # new image part + relationship
        name = f"media/gfx_{gfx}a-NWIRP.png"
        parts["word/" + name] = png.read_bytes()
        n = 1
        while f'Id="rIdGfx{n}"' in rels:
            n += 1
        new = f"rIdGfx{n}"
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="{new}" Type="{REL_IMAGE}" Target="{name}"/></Relationships>')

        # natural size: the PNG is rendered at 300 dpi from a deck built at its placed size, so placing
        # it at pixels / 300 keeps the chart text at exactly 8 pt; never wider than the text column
        w, h = Image.open(png).size
        cx = round(w / 300 * 914400)
        if cx > text_w:
            cx = text_w
        cy = round(cx * h / w)
        if is_chart:
            # a native Word chart: swap its chart graphic for a picture graphic (the chart part is
            # left in the package, unreferenced from the page)
            pic = ('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                   '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                   f'<pic:nvPicPr><pic:cNvPr id="0" name="{gfx}a-NWIRP.png"/><pic:cNvPicPr/></pic:nvPicPr>'
                   f'<pic:blipFill><a:blip r:embed="{new}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                   f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                   '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData>')
            p2 = re.sub(r'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">.*?</a:graphicData>',
                        pic, p, count=1, flags=re.S)
            p2 = re.sub(r'<wp:extent cx="\d+" cy="\d+"/>', f'<wp:extent cx="{cx}" cy="{cy}"/>', p2)
        else:
            p2 = fit_picture(p.replace(f'r:embed="{old}"', f'r:embed="{new}"'), cx, cy)
        doc = doc[:a] + p2 + doc[b:]
        delta = len(p2) - len(p)
        paras = [(s if s < a else s + delta, e if e < a else e + delta) for s, e in paras]
        paras[pic_i] = (a, b + delta)

        if is_chart:
            report.append(f"Figure {fig}: native chart -> {name}  ({cx / 914400:.2f} x {cy / 914400:.2f} in)")
            continue
        # drop the old image part if nothing else points at it
        target = re.search(rf'<Relationship Id="{old}"[^>]*Target="([^"]+)"', rels).group(1)
        if f'"{old}"' not in doc:
            rels = re.sub(rf'<Relationship Id="{old}"[^>]*/>', "", rels)
            if f'Target="{target}"' not in rels:
                parts.pop("word/" + target, None)
        report.append(f"Figure {fig}: {target} -> {name}  ({cx / 914400:.2f} x {cy / 914400:.2f} in)")

    if 'Extension="png"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1)
    parts["word/document.xml"] = doc.encode("utf-8")
    parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    tmp = docx.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for nm, data in parts.items():
            info = infos.get(nm)
            out.writestr(info if info else nm, data)
    tmp.replace(docx)
    print("\n".join(report))


if __name__ == "__main__":
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2], [tuple(a.split("=")) for a in sys.argv[3:]])
