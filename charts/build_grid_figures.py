"""GFX 1371 (Fig 6-3 heatmap) and 1381 (Fig 6-13 contour) as native PowerPoint grids.

Usage:  python build_grid_figures.py [--hm C:/Users/Artist/Documents/HM-261113]
Reads the grids from the chapter's charts-xlsx workbooks; writes ppt/<gfx>a-NWIRP.pptx (render the
PNG from PowerPoint, as Build-GfxChartDecks.ps1 does).

Excel has no heatmap/contour chart that survives restyling, so each data cell is a rectangle with
its value printed in it (Arial 8 pt), on one sequential Claret ramp shared by both figures (the same
quantity, expected coverage). Text colour per cell is white or black, whichever is higher contrast
(always >= 4.59:1).
"""
import argparse
from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Pt

ap = argparse.ArgumentParser()
ap.add_argument("--hm", default="C:/Users/Artist/Documents/HM-261113", help="folder holding the chapter folders")
args = ap.parse_args()
OUT = args.hm.rstrip("/\\") + "/NWIRP_2-1_Chapter-06_Modeling/"


def read_grid(name):
    """Workbook layout: x values across row 2 (from B), y values down column A (from row 3, top row =
    largest y), cell values in the body. Returns xs, ys ascending, and rows ordered to match ys."""
    ws = load_workbook(OUT + "charts-xlsx/" + name, data_only=True)["Data"]
    xs = [c for c in next(ws.iter_rows(min_row=2, max_row=2, min_col=2, values_only=True)) if c is not None]
    ys, rows = [], []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[0] is None or not isinstance(r[0], (int, float)):
            break
        ys.append(r[0]); rows.append(list(r[1:1 + len(xs)]))
    return xs, ys[::-1], rows[::-1]
W = 468.0
STRIP = 9.0
FONT, PT = "Arial", 8
INK, AXIS, LABEL = "000000", "000000", "000000"     # all text black
# legend house style (nw-2.3-Timeline-v2-ARTB.pptx)
LEG_INK, LEG_RULE, LEG_FONT, LEG_HEAD_FONT = "000000", "333333", "Arial", "Arial"     # Arial, like the chart text
SWATCH, SWATCH_RADIUS = 8.0, 0.32383
LOW, HIGH = (0xB8, 0x33, 0x5A), (0xF8, 0xEB, 0xEF)      # Claret Soft -> 90% tint


def ramp(v):
    v = max(0.0, min(1.0, v))
    return "".join(f"{round(a + (b - a) * v):02X}" for a, b in zip(LOW, HIGH))


def lum(h):
    c = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    c = [x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4 for x in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


def contrast(a, b):
    la, lb = sorted([lum(a), lum(b)], reverse=True)
    return (la + 0.05) / (lb + 0.05)


def ink_for(fill):
    # no colour on a continuous ramp gives 4.5:1 with both white and #33333B (mid-ramp tops out
    # near 3.6:1); white/black always reach >= 4.59:1, so cell values use black for the dark ink
    return "FFFFFF" if contrast("FFFFFF", fill) >= contrast("000000", fill) else "000000"


class Slide:
    def __init__(self, w, h):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Pt(w), Pt(h + STRIP)
        self.s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.w, self.h = w, h

    def rect(self, x, y, w, h, fill, text=None, color=INK, line=None):
        r = self.s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
        st = r._element.find(qn("p:style"))
        if st is not None:
            r._element.remove(st)
        r.fill.solid(); r.fill.fore_color.rgb = RGBColor.from_string(fill)
        if line:
            r.line.color.rgb = RGBColor.from_string(line); r.line.width = Pt(0.5)
        else:
            r.line.fill.background()
        if text is not None:
            self._txt(r.text_frame, text, color, PP_ALIGN.CENTER)
        return r

    def text(self, x, y, w, h, s, color=AXIS, align="c", rot=0, size=PT, bold=False, font=FONT, spacing=0):
        tb = self.s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
        self._txt(tb.text_frame, s, color, {"c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT, "l": PP_ALIGN.LEFT}[align],
                  size, bold, font, spacing)
        tb.rotation = rot
        return tb

    def line(self, x1, y1, x2, y2):
        c = self.s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
        st = c._element.find(qn("p:style"))
        if st is not None:
            c._element.remove(st)
        c.line.color.rgb = RGBColor.from_string(LEG_RULE); c.line.width = Pt(0.25)

    def swatch(self, x, y, fill):
        r = self.s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(SWATCH), Pt(SWATCH))
        st = r._element.find(qn("p:style"))
        if st is not None:
            r._element.remove(st)
        r.adjustments[0] = SWATCH_RADIUS
        r.fill.solid(); r.fill.fore_color.rgb = RGBColor.from_string(fill)
        r.line.fill.background()

    @staticmethod
    def _txt(tf, s, color, align, size=PT, bold=False, font=FONT, spacing=0):
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = align
        r = p.add_run(); r.text = s
        f = r.font; f.name = font; f.size = Pt(size); f.bold = bold; f.color.rgb = RGBColor.from_string(color)
        if spacing:
            r._r.get_or_add_rPr().set("spc", str(int(spacing * 100)))      # character spacing, 1/100 pt

    def save(self, name):
        self.text(self.w - 60, self.h + 1, 60, 8, name, color=LABEL, align="r", size=6)
        self.prs.save(OUT + f"ppt/{name}.pptx")


def grid(name, xs, ys, grid_vals, xtitle, ytitle, xfmt, yfmt, ylabel_every, width_in, height_in):
    W, H = width_in * 72, height_in * 72          # rows need >= ~11 pt for the 8 pt cell values
    sl = Slide(W, H)
    left, right, top, bottom = 46, 66, 4, 30
    pw, ph = W - left - right, H - top - bottom
    nx, ny = len(xs), len(ys)
    cw, ch = pw / nx, ph / ny
    for j, yv in enumerate(ys):                      # ys ascending; draw top row = largest y
        row = grid_vals[j]
        yy = top + (ny - 1 - j) * ch
        for i, v in enumerate(row):
            fill = ramp(v)
            sl.rect(left + i * cw, yy, cw, ch, fill, f"{v:.2f}", ink_for(fill), line="FFFFFF")
        if j % ylabel_every == 0:
            sl.text(left - 30, yy, 26, ch, yfmt.format(yv), align="r")
    for i, xv in enumerate(xs):
        sl.text(left + i * cw, top + ph + 2, cw, 10, xfmt.format(xv))
    sl.text(left, top + ph + 14, pw, 10, xtitle)
    sl.text(left - 46 - ph / 2 + 5, top + ph / 2 - 5, ph, 10, ytitle, rot=270)
    # legend in the house style (nw-2.3-Timeline-v2-ARTB.pptx slide 6): a 0.25 pt rule, a letter-spaced
    # LEGEND heading and title, then stepped rounded swatches with right-aligned values
    rx = W - right + 8                    # no vertical rule beside the legend
    lx = rx + 7
    sl.text(lx, top, right - 15, 10, "LEGEND", color=LEG_INK, align="l", font=LEG_HEAD_FONT, bold=True, spacing=1)
    sl.text(lx, top + 11, right - 15, 10, "Expected", color=LEG_INK, align="l", font=LEG_HEAD_FONT)
    sl.text(lx, top + 21, right - 15, 10, "coverage", color=LEG_INK, align="l", font=LEG_HEAD_FONT)
    for k, t in enumerate([i / 10 for i in range(11)]):
        cy = top + 40 + k * 12
        sl.text(lx, cy - 5, 16, 10, f"{t:.1f}", color=LEG_INK, align="r", font=LEG_FONT)
        sl.swatch(lx + 21, cy - 4, ramp(t))
    sl.save(name)
    return H + STRIP


xs, ys, rows = read_grid("image05_coverage-heatmap.xlsx")
H1 = grid("1371a-NWIRP", xs, ys, rows,
          "Road link disruption probability", "Hospital disruption probability", "{:g}", "{:.1f}", 2, 5.0, 3.5)
xs, ys, rows = read_grid("image19_coverage-contour.xlsx")
H2 = grid("1381a-NWIRP", xs, ys, rows,
          "Road disruption reduction factor", "NDWE recovery factor", "{:g}", "{:.2f}", 2, 5.5, 3.4)
print("1371a", H1, " 1381a", H2)
