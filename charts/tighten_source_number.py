"""Pull each figure's source number up under the art and trim the slide to it.

Usage:  python tighten_source_number.py PPT_DIR [PPT_DIR ...]

Excel's chart area keeps padding under its last axis title, and the legend reserves a full row
pitch, so a source number placed under the reserved space floats well below the visible art. For
each <gfx>a-NWIRP.pptx this reads the rendered PNG (../png/<same name>.png, 300 dpi), finds the lowest
row of ink outside the source number itself, moves the source number to GAP_PT below it, and sets
the slide height to end just under the source number. The slide size is edited in the file, not
through PowerPoint, so nothing is rescaled. Re-render the PNGs afterwards.
"""
import sys
from pathlib import Path

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Pt

GAP_PT = 2.0          # space between the art and the source number
LABEL_H_PT = 7.5      # 6 pt text, single spaced
INK = 245             # a pixel darker than this counts as art
PX_PER_PT = 300 / 72


def tighten(pptx_path):
    png = pptx_path.parent.parent / "png" / (pptx_path.stem + ".png")
    prs = Presentation(pptx_path)
    slide = prs.slides[0]
    label = next(s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip().endswith("-NWIRP"))

    a = np.asarray(Image.open(png).convert("L")) < INK
    sw_pt = prs.slide_width / 12700
    sx = a.shape[1] / sw_pt                                        # px per pt as rendered
    # blank out the source number so it does not count as art
    x0, y0 = int(label.left / 12700 * sx) - 2, int(label.top / 12700 * sx) - 2
    x1, y1 = int((label.left + label.width) / 12700 * sx) + 2, int((label.top + label.height) / 12700 * sx) + 2
    a[max(y0, 0):y1, max(x0, 0):x1] = False
    rows = np.nonzero(a.any(axis=1))[0]
    bottom_pt = (rows[-1] + 1) / sx

    label.top = Pt(bottom_pt + GAP_PT)
    label.height = Pt(LABEL_H_PT)
    new_h = bottom_pt + GAP_PT + LABEL_H_PT
    old_h = prs.slide_height / 12700
    prs.slide_height = Pt(new_h)
    prs.save(pptx_path)
    return old_h, new_h


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    for d in sys.argv[1:]:
        for p in sorted(Path(d).glob("*a-NWIRP.pptx")):
            old_h, new_h = tighten(p)
            print(f"{p.stem}: {old_h:.1f} -> {new_h:.1f} pt  (removed {old_h - new_h:.1f} pt)")
