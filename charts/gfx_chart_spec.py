"""Write the GFX chart spec read by Build-GfxChartDecks.ps1, and the workbook for Figure 4-24 (GFX 1336).

Usage:  python gfx_chart_spec.py [--hm C:/Users/Artist/Documents/HM-261113] [--out chart_spec.json]

Each entry names the chapter folder, the chart workbook (charts-xlsx/), the original image whose
proportions set the slide height, the chart kind and layout, and the series colours.

Colours follow charts/chart-styling-guide.md (the soft palette the .bas modules use), keyed by
series name; semantic choices follow the guide's own assignments (Claret = Hurricane Maria /
power / severe, Navy = primary / IT-comms, the damage-state ramp for ordered levels).
"""
import argparse
import json
from PIL import Image
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series

ap = argparse.ArgumentParser()
ap.add_argument("--hm", default="C:/Users/Artist/Documents/HM-261113", help="folder holding the chapter folders")
ap.add_argument("--hm5", default="C:/Users/Artist/Documents/HM-261119", help="folder holding the Chapter 5 folder")
ap.add_argument("--out", default="chart_spec.json")
args = ap.parse_args()
H = args.hm.rstrip("/\\") + "/"
C4 = H + "NWIRP_2-1_Chapter-04_Wireless/"
C6 = H + "NWIRP_2-1_Chapter-06_Modeling/"
C5 = args.hm5.rstrip("/\\") + "/NWIRP_2-1_Chapter-05_Vegetation/"

NAVY, TAWNY, SHAM, PICTON, CLARET, HARVEST, MULB = ("1A3C8F", "E87A30", "2AB58E", "7CC5EE",
                                                    "B8335A", "F0B840", "D07EAC")
INK, AXIS = "33333B", "6B6B76"


def tint(hexc, t):
    r, g, b = (int(hexc[i:i + 2], 16) for i in (0, 2, 4))
    return "".join(f"{round(c + (255 - c) * t):02X}" for c in (r, g, b))


def aspect(path):
    w, h = Image.open(path).size
    return round(h / w, 4)


UNC, COL = NAVY, CLARET
METHODS = {"Standard": NAVY, "MLE": NAVY, "Lasso Minimum": PICTON, "Lasso One-standard-error": CLARET}
CONDITION = {"Fully functional": NAVY, "Immediately following HM": CLARET, "One month after HM": PICTON}
PORTS = ["None removed", "Las Mareas", "Jobos Bay", "Yabucoa", "Fajardo", "San Juan", "Guánica", "Ponce",
         "Guayanilla", "Mayagüez", "Aguadilla", "Arecibo"]
port_cols = [NAVY, PICTON, HARVEST, TAWNY, CLARET, MULB, SHAM] + [tint(c, 0.5) for c in (NAVY, TAWNY, CLARET, SHAM, MULB)]

hist = dict(kind="hist", layout="stack", panel_titles=["Uncollapsed", "Collapsed"], chart_colors=[UNC, COL])
cat2 = dict(kind="bar", layout="stack", panel_titles=["Uncollapsed", "Collapsed"], chart_colors=[UNC, COL],
            labels=True)

spec = [
    # ---------------- Chapter 6
    dict(gfx="1370", fig="6-2", dir=C6, wb="charts-xlsx/image04_cumulative-inaccessibility.xlsx",
         img="media/image4.jpeg", kind="line", layout="single",
         colors={**CONDITION, "0.5 h threshold": AXIS}),
    dict(gfx="1372", fig="6-4", dir=C6, wb="charts-xlsx/image06_fragility-recovery-curves.xlsx",
         img="media/image6.jpeg", kind="line", layout="side",
         panel_titles=["(a) Hospital fragility curves", "(b) Hospital recovery curves"],
         colors={"DS1": PICTON, "DS2": HARVEST, "DS3": TAWNY, "DS4": CLARET}),
    dict(gfx="1373", fig="6-5", dir=C6, wb="charts-xlsx/image07_coverage-by-disruption.xlsx",
         img="media/image7.jpeg", kind="bar", layout="single",
         colors={"Hospital disruption only": SHAM, "Road disruption only": HARVEST,
                 "Road and hospital disruptions": CLARET}),
    dict(gfx="1377", fig="6-9", dir=C6, wb="charts-xlsx/image15_coverage-vs-temporary-facilities.xlsx",
         img="media/image15.jpeg", kind="line", layout="single", markers=True,
         colors={"Immediately following HM, γ = 0.5 h": CLARET, "One month after HM, γ = 0.5 h": PICTON}),
    dict(gfx="1379", fig="6-11", dir=C6, wb="charts-xlsx/image17_pump-station-coverage-by-port.xlsx",
         img="media/image17.jpeg", kind="bar", layout="single", colors=dict(zip(PORTS, port_cols))),
    dict(gfx="1380", fig="6-12", dir=C6, wb="charts-xlsx/image18_coverage-vs-port-closure.xlsx",
         img="media/image18.jpeg", kind="line", layout="side", markers=True,
         panel_titles=["(a) Fully functional road system", "(b) Road system 1 month after HM"],
         colors={f"P(lose power) = {k}": c for k, c in
                 zip(["0", "0.2", "0.4", "0.6", "0.8", "1"], [SHAM, PICTON, HARVEST, TAWNY, CLARET, NAVY])}),
    # ---------------- Chapter 5
    dict(gfx="1359", fig="5-5", dir=C5, wb="charts-xlsx/fig5-05_variable-importance.xlsx",
         img="media/image7.jpeg", kind="bar", layout="grid", grid=[3, 1], colors={"MDI": NAVY},
         value_title="Mean decrease in\nimpurity (MDI)", plot_left=70,
         panel_titles=["(a) Power system model", "(b) Transportation system model, immediately following HM",
                       "(c) Transportation system model, one month after HM"]),
    dict(gfx="1360", fig="5-6", dir=C5, wb="charts-xlsx/fig5-06_partial-dependence.xlsx",
         img="media/image8.jpeg", kind="line", layout="grid", grid=[3, 3], colors={"Average NDWE": NAVY}),
    dict(gfx="1363", fig="5-9", dir=C5, wb="charts-xlsx/fig5-09_ndwe-change-histogram.xlsx",
         img="media/image11.png", kind="hist", layout="single", colors={"Count": CLARET},
         point_colors=[CLARET] * 10 + [PICTON] * 5,
         legend_items=[["Decrease in NDWE", CLARET], ["Increase in NDWE", PICTON]]),
    dict(gfx="1365", fig="5-11", dir=C5, wb="charts-xlsx/fig5-11_canopy-vs-gust.xlsx",
         img="media/image13.jpeg", kind="line", layout="single", markers=True,
         colors={"Mean gust": NAVY, "Max gust": TAWNY}),
    dict(gfx="1366", fig="5-12", dir=C5, wb="charts-xlsx/fig5-12_cost-benefit.xlsx",
         img="media/image14.png", kind="line", layout="single", markers=True,
         colors={"Net benefit ($ billion)": NAVY}),
    dict(gfx="1368", fig="5-14", dir=C5, wb="charts-xlsx/fig5-14_vm-cost-sensitivity.xlsx",
         img="media/image16.png", kind="line", layout="single", markers=True,
         colors={"Proportion of applicable lines with VM": NAVY}),
    # ---------------- Chapter 4
    dict(gfx="1317", fig="4-5", dir=C4, wb="charts-xlsx/image13_cell-sites-out-by-municipality.xlsx",
         img="media/image13.png", kind="line", layout="single",
         colors={"CULEBRA": CLARET, "RINCON": PICTON, "SAN JUAN": TAWNY, "VIEQUES": MULB, "YABUCOA": SHAM,
                 "UTUADO": HARVEST, "PUERTO RICO (ALL)": INK}),
    dict(gfx="1318", fig="4-6", dir=C4, wb="charts-xlsx/image14_causes-of-outages-stacked.xlsx",
         img="media/image14.png", kind="area", layout="single",
         colors={"Out due to power outage": CLARET, "Out due to network connection": NAVY,
                 "Out due to damage": TAWNY}),
    dict(gfx="1329", fig="4-17", dir=C4, wb="charts-xlsx/image29_wind-speed-histogram.xlsx",
         img="media/image29.png", labels=True, **hist),
    dict(gfx="1330", fig="4-18", dir=C4, wb="charts-xlsx/image30_structure-types.xlsx",
         img="media/image30.png", **cat2),
    dict(gfx="1331", fig="4-19", dir=C4, wb="charts-xlsx/image31_construction-dates.xlsx",
         img="media/image31.png", **cat2),
    dict(gfx="1332", fig="4-20", dir=C4, wb="charts-xlsx/image32_tower-height-histogram.xlsx",
         img="media/image32.png", **hist),
    dict(gfx="1333", fig="4-21", dir=C4, wb="charts-xlsx/image33_ground-elevation-histogram.xlsx",
         img="media/image33.png", **hist),
    dict(gfx="1334", fig="4-22", dir=C4, wb="charts-xlsx/image34_tower-exposures.xlsx",
         img="media/image34.png", kind="bar", layout="single", colors={"Count": NAVY},
         point_colors=[NAVY, PICTON, HARVEST]),
    dict(gfx="1336", fig="4-24", dir=C4, wb="charts-xlsx/fig4-24_exposure-change.xlsx", img=None, aspect_override=0.62,
         **cat2),
    dict(gfx="1337", fig="4-25", dir=C4, wb="charts-xlsx/image39_regression-coefficients.xlsx",
         img="media/image39.png", kind="bar", layout="single", colors=METHODS),
    dict(gfx="1338", fig="4-26", dir=C4, wb="charts-xlsx/image40_roc-curves.xlsx",
         img="media/image40.png", kind="line", layout="single",
         colors={**METHODS, "Reference points": INK}, marker_only=["Reference points"]),
    dict(gfx="1339", fig="4-27", dir=C4, wb="charts-xlsx/image41_exposure-change-probability.xlsx",
         img="media/image41.png", kind="line", layout="single", x_format="0",
         colors={"Fitted probability of category Q": CLARET, "Lower confidence bound": CLARET,
                 "Upper confidence bound": CLARET}),
    dict(gfx="1354", fig="4-42", dir=C4, wb="charts-xlsx/image58_regression-coefficients-v2.xlsx",
         img="media/image58.png", kind="bar", layout="single", colors=METHODS),
]
# Placed size of each figure's chart area (inches). Chosen by eye so a single chart takes about a third
# of a page: narrow where the content allows, full column width only for wide multi-series or
# two-panel figures. The legend and source number are added below this.
SIZES = {'1370': (5.0, 3.0), '1372': (6.5, 2.4), '1373': (4.5, 2.8), '1377': (4.5, 2.8), '1379': (6.5, 3.0), '1380': (6.5, 2.8), '1317': (6.5, 3.0), '1318': (5.5, 2.8), '1329': (6.5, 4.2), '1330': (5.0, 3.6), '1331': (6.0, 3.6), '1332': (6.5, 3.8), '1333': (6.5, 3.8), '1334': (3.5, 2.4), '1336': (6.0, 3.8), '1337': (6.0, 3.0), '1338': (3.6, 3.4), '1339': (4.5, 3.2), '1354': (6.0, 3.0),
         '1359': (6.5, 3.2), '1360': (6.5, 4.8), '1363': (5.5, 3.0), '1365': (5.0, 3.0), '1366': (4.5, 3.0), '1368': (4.5, 3.0)}
for s in spec:
    s.pop("aspect_override", None)
    s["width_in"], s["height_in"] = SIZES[s["gfx"]]

# ---- Figure 4-24 (GFX 1336): native Word chart data, exact counts
cats = ["Started exposure C or D", "Started exposure B, changed to C", "Started exposure B, did not change",
        "Started exposure B, unable to assess post-HM exposure"]
unc, col = [287, 138, 117, 25], [8, 12, 6, 1]
wb = Workbook(); ws = wb.active; ws.title = "Data"
ws.append(["Exposure change", "Uncollapsed", "Collapsed"])
for row in zip(cats, unc, col):
    ws.append(list(row))
for i, (title, colidx, anchor) in enumerate([("Uncollapsed", 2, "E2"), ("Collapsed", 3, "E20")]):
    ch = BarChart(); ch.type = "col"; ch.title = title
    ch.add_data(Reference(ws, min_col=colidx, min_row=1, max_row=5), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5))
    ch.y_axis.title = "Count"; ch.legend = None
    if i == 1:
        ch.x_axis.title = "Exposure change (ASCE 7-22)"
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ws.add_chart(ch, anchor)
ws2 = wb.create_sheet("Notes")
ws2["A1"] = "Source: native chart in NWIRP_2-1_Chapter-04_Wireless.docx (HM-261115 master), Figure 4-24. Exact counts."
wb.save(C4 + "charts-xlsx/fig4-24_exposure-change.xlsx")

json.dump(spec, open(args.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(len(spec), "charts;", [(s["gfx"], s["width_in"], s["height_in"]) for s in spec])
