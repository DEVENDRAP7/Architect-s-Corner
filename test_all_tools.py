#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_all_tools.py -- run EVERY catalog tool against a generated test drawing.

Builds test_building.dxf (a synthetic plan that follows all the tagging
conventions the tools expect: C-tags, F-tags, floor titles, LVL marks,
door/window/sanitary/parking blocks, rooms, dims, hatches, defects), then
runs each of the 93 catalog tools in-process and prints a PASS/WARN/FAIL
table.

  PASS = produced real output
  WARN = ran fine but found nothing to report (often legitimate)
  FAIL = crashed / errored

Usage:  python test_all_tools.py            (add --rebuild to regenerate DXF)
"""

import os
import re
import sys

import ezdxf

import ArchTools

HERE = os.path.dirname(os.path.abspath(__file__))
DXF = os.path.join(HERE, "test_building.dxf")


# ----------------------------------------------------------------------
# 1. the test drawing -- every convention the tools read
# ----------------------------------------------------------------------
def build():
    doc = ezdxf.new(setup=True)   # setup=True -> default dimstyle/linetypes
    msp = doc.modelspace()

    def txt(s, x, y, layer="ANNOT", h=250):
        e = msp.add_mtext(s, dxfattribs={"layer": layer, "char_height": h})
        e.set_location((x, y))

    # --- levels (heights/colvol/slab/concrete read these) ---
    for s, y in [("ROAD LVL +0.00", 0), ("GROUND FLOOR LVL +1.20", -600),
                 ("FIRST FLOOR LVL +4.70", -1200), ("SECOND FLOOR LVL +8.20", -1800),
                 ("TERRACE FLOOR LVL +11.70", -2400), ("PARAPET LVL +12.90", -3000)]:
        txt(s, -20000, y, "LEVELS")

    # --- two floor plans: title + tagged columns on a grid ---
    cols = [("C1", "450x450"), ("C2", "450x450"), ("C3", "300x450"),
            ("C4", "450x450"), ("C5", "300x450"), ("C6", "300x300"),
            ("C7", "450x450"), ("C8", "300x450"), ("C9", "300x300")]
    for title, ox in [("GROUND FLOOR PLAN", 0), ("FIRST FLOOR PLAN", 60000)]:
        txt(title, ox + 9000, 16000, "TITLES", 400)
        for i, (tag, sz) in enumerate(cols):
            x, y = ox + (i % 3) * 6000, (i // 3) * 5000
            txt(f"{tag}\n{sz}", x, y, "COLUMNS")
            w, h = (int(v) for v in sz.split("x"))
            msp.add_lwpolyline([(x-w/2, y-h/2), (x+w/2, y-h/2),
                                (x+w/2, y+h/2), (x-w/2, y+h/2)],
                               close=True, dxfattribs={"layer": "COLUMNS"})

    # --- footings ---
    for i, (tag, sz) in enumerate([("F1", "1800x1800"), ("F2", "1500x1500"),
                                   ("F3", "1500x1500"), ("F4", "1200x1200")]):
        txt(f"{tag}\n{sz}", i * 6000, -9000, "FOOTINGS")

    # --- building boundary (footprint 18 x 12 m) ---
    msp.add_lwpolyline([(-3000, -2500), (15000, -2500), (15000, 11500),
                        (-3000, 11500)], close=True,
                       dxfattribs={"layer": "BOUNDARY"})

    # --- rooms: closed polylines + names (room-areas / room-count) ---
    rooms = [("OFFICE", 0, 0, 5000, 4000), ("LAB", 6000, 0, 5000, 4000),
             ("STORE", 0, 5000, 5000, 4000), ("TOILET", 6000, 5000, 3000, 4000)]
    for nm, x, y, w, h in rooms:
        msp.add_lwpolyline([(x, y), (x+w, y), (x+w, y+h), (x, y+h)],
                           close=True, dxfattribs={"layer": "ROOMS"})
        txt(nm, x + w/2, y + h/2, "ROOM-NAMES")

    # --- walls / plinth beams / chajja runs (length tools) ---
    for i in range(4):
        msp.add_line((0, 12000 + i*300), (18000, 12000 + i*300),
                     dxfattribs={"layer": "WALLS"})
    for i in range(3):
        msp.add_line((0, -4000 - i*300), (12000, -4000 - i*300),
                     dxfattribs={"layer": "PLINTH-BEAM"})
    msp.add_line((0, -6000), (9000, -6000), dxfattribs={"layer": "CHAJJA"})

    # --- blocks: doors, windows, sanitary, parking, machines ---
    for name in ["DOOR_900", "DOOR_1200", "WINDOW_1500", "WC_EWC",
                 "WASH_BASIN", "CAR_PARK", "MACHINE_CNC"]:
        blk = doc.blocks.new(name=name)
        blk.add_lwpolyline([(0, 0), (500, 0), (500, 500), (0, 500)], close=True)
    inserts = [("DOOR_900", 3), ("DOOR_1200", 2), ("WINDOW_1500", 4),
               ("WC_EWC", 2), ("WASH_BASIN", 2), ("CAR_PARK", 6),
               ("MACHINE_CNC", 3)]
    px = 20000
    for name, n in inserts:
        for k in range(n):
            msp.add_blockref(name, (px, k * 1500),
                             dxfattribs={"layer": "BLOCKS"})
        px += 2500

    # --- opening schedule texts (schedule / door-width-check) ---
    for s, x in [("D1\n1000x2100", 0), ("D2\n900x2100", 3000),
                 ("W1\n1500x1200", 6000), ("ED1\n1800x2400", 9000)]:
        txt(s, x, -12000, "SCHEDULE")

    # --- staircase / lift labels ---
    txt("STAIRCASE-1", 12000, 5000, "CORE")
    txt("LIFT-1", 13500, 5000, "CORE")

    # --- dimensions (dims / qa-report) ---
    for i in range(3):
        d = msp.add_linear_dim(base=(i*6000, -1500), p1=(i*6000, -2500),
                               p2=(i*6000 + 6000, -2500))
        d.render()

    # --- circles (circles tool) ---
    for i in range(5):
        msp.add_circle((30000 + i * 1200, 0), 400,
                       dxfattribs={"layer": "CIRCLES"})

    # --- hatch (hatch tool) ---
    h = msp.add_hatch(dxfattribs={"layer": "FLOOR-FINISH"})
    h.paths.add_polyline_path([(0, 0), (5000, 0), (5000, 4000), (0, 4000)],
                              is_closed=True)

    # --- deliberate defects (qa-report) ---
    msp.add_line((0, -7000), (5000, -7000), dxfattribs={"layer": "WALLS"})
    msp.add_line((0, -7000), (5000, -7000), dxfattribs={"layer": "WALLS"})
    msp.add_line((100, -7200), (104, -7200), dxfattribs={"layer": "WALLS"})
    msp.add_lwpolyline([(40000, 0), (44000, 0), (44000, 3000), (40010, 20)],
                       dxfattribs={"layer": "WALLS"})
    doc.layers.add("UNUSED-LAYER")
    msp.add_blockref("DOOR_900", (46000, 0))          # block on layer 0

    doc.saveas(DXF)
    print(f"built {DXF}")


# ----------------------------------------------------------------------
# 2. run every catalog tool
# ----------------------------------------------------------------------
# values for required params the drawing can't supply
OVERRIDES = {
    "--other": DXF, "--q": "LAB", "--unit": "m2-ft2",
    "--keywords": "LAB,STORE,OFFICE", "--layer": "ROOMS", "--type": "LINE",
}
# optional params still worth passing so the tool has something to find
PASS_OPTIONAL = {"--q", "--type", "--layer"}
# per-tool layer choices (so layer tools test the right content)
TOOL_LAYER = {
    "beams": "WALLS", "wall-area": "WALLS", "plinth-beams": "PLINTH-BEAM",
    "chajja": "CHAJJA", "footings": "FOOTINGS", "centroid": "COLUMNS",
    "layer-detail": "COLUMNS", "bbox-layer": "COLUMNS",
    "layer-area": "ROOMS", "layer-length": "WALLS",
    "room-areas": "ROOMS", "coving-length": "ROOMS", "fixtures": "BLOCKS",
}
NUM_DEFAULT = 50.0

def argv_for(tool):
    argv = [tool["command"]]
    if tool["needs_building"] or tool["needs_path"]:
        argv.append(DXF)
    for p in tool["params"]:
        name = p["name"]
        if p["type"] == "flag":
            continue
        if name == "--layer":
            argv += [name, TOOL_LAYER.get(tool["command"], OVERRIDES["--layer"])]
            continue
        if not p["required"]:
            if name in PASS_OPTIONAL:
                argv += [name, str(OVERRIDES[name])]
            continue
        if name in OVERRIDES:
            argv += [name, str(OVERRIDES[name])]
        elif p["type"] in ("float", "int"):
            argv += [name, str(int(NUM_DEFAULT) if p["type"] == "int" else NUM_DEFAULT)]
        else:
            argv += [name, "test"]
    return argv


WARN_PAT = re.compile(
    r"^(no |error|pass --|\(none|not found)|no .* found|ERROR", re.I)

def classify(out):
    if not out or out == "error":
        return "FAIL"
    first = out.strip().splitlines()[0]
    if WARN_PAT.search(first.strip()):
        return "WARN"
    if "Traceback" in out or out.strip().lower().startswith("error"):
        return "FAIL"
    return "PASS"


def main():
    if "--rebuild" in sys.argv or not os.path.isfile(DXF):
        build()
    tools = ArchTools.build_catalog()
    results = {"PASS": [], "WARN": [], "FAIL": []}
    for t in tools:
        argv = argv_for(t)
        try:
            out = ArchTools.run_command(argv)
            status = classify(out)
        except Exception as e:
            out, status = f"{type(e).__name__}: {e}", "FAIL"
        results[status].append((t["command"], out.strip().splitlines()[0][:90]
                                if out else "(empty)"))
        mark = {"PASS": " . ", "WARN": " ? ", "FAIL": " X "}[status]
        print(f"{mark}{t['command']:24} {out.strip().splitlines()[0][:80] if out else '(empty)'}")
    print("\n" + "=" * 60)
    print(f"PASS {len(results['PASS'])}   WARN {len(results['WARN'])}   "
          f"FAIL {len(results['FAIL'])}   (of {len(tools)})")
    if results["WARN"]:
        print("\nWARN (ran, found nothing -- check if expected):")
        for c, o in results["WARN"]:
            print(f"  {c:24} {o}")
    if results["FAIL"]:
        print("\nFAIL:")
        for c, o in results["FAIL"]:
            print(f"  {c:24} {o}")
        sys.exit(1)


if __name__ == "__main__":
    main()
