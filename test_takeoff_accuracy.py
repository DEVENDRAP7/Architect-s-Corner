#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_takeoff_accuracy.py -- ground-truth accuracy tests for the geometry engine.

Every drawing built here contains NO text, NO tags, NO dimensions and NO
meaningful layer names -- the case the tools are meant to survive. Each file
has a known correct answer, so accuracy is measured rather than assumed.

Two axes are covered:

  * HOW the column is drawn  -- block, loose lines, legacy polyline, hatch,
    circle, unflagged-closed polyline. All of these used to return zero.
  * WHAT the building is     -- small house through wide-span shed, in
    millimetres, metres and feet. Fixed thresholds used to reject most of them.

Plus areas (curved boundaries, nested plots), and walls/beams drawn as two
faces, which used to be counted twice.

Usage:  python test_takeoff_accuracy.py
"""

import math
import sys

import ezdxf

import dxfgeom as G

FAILS = []
PASSES = [0]


def check(name, got, want, tol=0):
    ok = abs(got - want) <= tol if isinstance(want, (int, float)) else got == want
    if ok:
        PASSES[0] += 1
    else:
        FAILS.append(f"{name}: got {got!r}, want {want!r}")
    flag = "ok  " if ok else "FAIL"
    print(f"  {flag}  {name:52} got {got!r:>12}  want {want!r}")


# ----------------------------------------------------------------------
# builders -- all textless
# ----------------------------------------------------------------------
def col_grid(nx, ny, spacing, w=450, h=600, style="lwpolyline", unit_div=1.0):
    """A bare column grid drawn in one of several styles."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()
    u = lambda v: v / unit_div

    def corners(cx, cy):
        return [(u(cx - w / 2), u(cy - h / 2)), (u(cx + w / 2), u(cy - h / 2)),
                (u(cx + w / 2), u(cy + h / 2)), (u(cx - w / 2), u(cy + h / 2))]

    if style == "block":
        blk = doc.blocks.new("COL")
        blk.add_lwpolyline([(u(-w/2), u(-h/2)), (u(w/2), u(-h/2)),
                            (u(w/2), u(h/2)), (u(-w/2), u(h/2))], close=True)

    for r in range(ny):
        for c in range(nx):
            cx, cy = c * spacing, r * spacing
            pts = corners(cx, cy)
            if style == "lwpolyline":
                msp.add_lwpolyline(pts, close=True)
            elif style == "block":
                msp.add_blockref("COL", (u(cx), u(cy)))
            elif style == "lines":
                for i in range(4):
                    msp.add_line(pts[i], pts[(i + 1) % 4])
            elif style == "polyline":
                msp.add_polyline2d(pts, close=True)
            elif style == "unflagged":
                msp.add_lwpolyline(pts + [pts[0]], close=False)
            elif style == "hatch":
                hh = msp.add_hatch()
                hh.paths.add_polyline_path(pts, is_closed=True)
            elif style == "solid":
                msp.add_solid([pts[0], pts[1], pts[3], pts[2]])
    return doc


def analyse(doc):
    ents = list(G.iter_entities(doc.modelspace()))
    regs = G.regions(ents)
    scale = G.detect_scale(doc, regs)
    return G.detect_columns(regs, scale), scale


# ----------------------------------------------------------------------
print("=" * 88)
print("1. HOW THE COLUMN IS DRAWN   (16 columns, 450x600, 4 m grid, no text)")
print("=" * 88)
for style in ["lwpolyline", "block", "lines", "polyline", "unflagged", "hatch", "solid"]:
    doc = col_grid(4, 4, 4000, style=style)
    res, _ = analyse(doc)
    check(f"style={style}", res.total, 16)

print()
print("=" * 88)
print("2. WHAT THE BUILDING IS   (drawn as blocks, no text -- the hard case)")
print("=" * 88)
BUILDINGS = [
    ("small house      3x2 @ 3.5 m", 3, 2, 3500, 6),
    ("bungalow         3x3 @ 3.5 m", 3, 3, 3500, 9),
    ("apartment        4x4 @ 4.0 m", 4, 4, 4000, 16),
    ("tight residential 4x4 @ 2.2 m", 4, 4, 2200, 16),
    ("office           5x4 @ 6.0 m", 5, 4, 6000, 20),
    ("warehouse        4x3 @ 12 m", 4, 3, 12000, 12),
    ("shed             5x3 @ 15 m", 5, 3, 15000, 15),
]
for label, nx, ny, sp, truth in BUILDINGS:
    doc = col_grid(nx, ny, sp, style="block")
    res, _ = analyse(doc)
    check(label, res.total, truth)

print()
print("=" * 88)
print("3. DRAWING UNITS   (same 4x4 building authored in different units)")
print("=" * 88)
for label, div, want_unit in [("millimetres", 1.0, "mm"),
                              ("metres", 1000.0, "m"),
                              ("centimetres", 10.0, "cm")]:
    doc = col_grid(4, 4, 4000, unit_div=div)
    res, scale = analyse(doc)
    check(f"units={label}: columns", res.total, 16)
    check(f"units={label}: detected", scale.unit, want_unit)

print()
print("=" * 88)
print("4. FIXTURES MUST NOT BECOME COLUMNS")
print("=" * 88)
# 16 real columns on a grid + a row of 6 parking bays + scattered fixtures,
# all the same size, all as blocks, no text anywhere.
doc = col_grid(4, 4, 4000, style="block")
msp = doc.modelspace()
fx = doc.blocks.new("FIXTURE")
fx.add_lwpolyline([(-225, -300), (225, -300), (225, 300), (-225, 300)], close=True)
for i in range(6):                       # a row of parking bays
    msp.add_blockref("FIXTURE", (i * 2500, -9000))
for (x, y) in [(1500, 2300), (7100, 900), (2600, 11200), (9900, 6400)]:
    msp.add_blockref("FIXTURE", (x, y))  # scattered sanitary fixtures
res, _ = analyse(doc)
check("16 columns kept, 10 fixtures rejected", res.total, 16)

print()
print("=" * 88)
print("5. AREAS")
print("=" * 88)
# 6 x 4 m room with a semicircular bay on one side
doc = ezdxf.new("R2010")
msp = doc.modelspace()
lw = msp.add_lwpolyline([(0, 0), (6000, 0), (6000, 4000), (0, 4000)], close=True)
lw.set_points([(6000, 0, 0, 0, 0), (6000, 4000, 0, 0, 1.0),
               (0, 4000, 0, 0, 0), (0, 0, 0, 0, 0)], format="xyseb")
regs = G.regions(list(G.iter_entities(msp)))
scale = G.detect_scale(doc, regs)
true_area = 6.0 * 4.0 + math.pi * 3.0 ** 2 / 2   # bay on the 6 m side
check("curved boundary area m2", round(scale.area_to_m2(regs[0].area), 2),
      round(true_area, 2), 0.05)

# plot 30x20 m with the building 18x12 m drawn inside it
doc = ezdxf.new("R2010")
msp = doc.modelspace()
msp.add_lwpolyline([(0, 0), (30000, 0), (30000, 20000), (0, 20000)], close=True)
msp.add_lwpolyline([(5000, 4000), (23000, 4000), (23000, 16000), (5000, 16000)],
                   close=True)
regs = G.regions(list(G.iter_entities(msp)))
scale = G.detect_scale(doc, regs)
net, outer, nested = G.net_areas(regs)
check("nested plot: net area m2", round(scale.area_to_m2(net)), 600 - 216)
check("nested plot: outer shapes", len(outer), 1)
check("nested plot: holes found", nested, 1)

# rooms drawn as things other than a closed lwpolyline
for style, maker in [
    ("legacy POLYLINE", lambda m: m.add_polyline2d(
        [(0, 0), (5000, 0), (5000, 4000), (0, 4000)], close=True)),
    ("HATCH", lambda m: m.add_hatch().paths.add_polyline_path(
        [(0, 0), (5000, 0), (5000, 4000), (0, 4000)], is_closed=True)),
    ("CIRCLE", lambda m: m.add_circle((0, 0), 2523.13)),
    ("unflagged close", lambda m: m.add_lwpolyline(
        [(0, 0), (5000, 0), (5000, 4000), (0, 4000), (0, 0)], close=False)),
]:
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()
    maker(msp)
    regs = G.regions(list(G.iter_entities(msp)))
    got = round(sum(r.area for r in regs) / 1e6, 1)
    check(f"room as {style}: area m2", got, 20.0, 0.2)

print()
print("=" * 88)
print("6. WALLS AND BEAMS DRAWN AS TWO FACES")
print("=" * 88)
doc = ezdxf.new("R2010")
msp = doc.modelspace()
for off in (0, 230):                       # one 10 m wall, two faces
    msp.add_line((0, off), (10000, off))
segs = G.segments(list(G.iter_entities(msp)))
centres, thicks = G.pair_faces(segs, max_sep=600)
check("double-line wall: runs", len(centres), 1)
check("double-line wall: length m", round(sum(G.seg_length(s) for s in centres) / 1000, 1), 10.0, 0.05)
check("double-line wall: thickness mm", round(thicks[0]), 230, 1)

doc = ezdxf.new("R2010")
msp = doc.modelspace()
for seg in range(3):                       # one 6 m beam split at 2 columns
    msp.add_line((seg * 2000, 0), ((seg + 1) * 2000, 0))
segs = G.segments(list(G.iter_entities(msp)))
merged = G.merge_collinear(segs)
check("beam split at columns: runs", len(merged), 1)
check("beam split at columns: length m",
      round(sum(G.seg_length(s) for s in merged) / 1000, 1), 6.0, 0.05)

print()
print("=" * 88)
print("7. OPENINGS NESTED INSIDE A PLAN BLOCK")
print("=" * 88)
doc = ezdxf.new("R2010")
door = doc.blocks.new("DOOR-900")
door.add_line((0, 0), (900, 0))
plan = doc.blocks.new("GF-PLAN")
for i in range(6):
    plan.add_blockref("DOOR-900", (i * 3000, 0))
msp = doc.modelspace()
msp.add_blockref("GF-PLAN", (0, 0))
found = sum(1 for e, _d in G.iter_inserts(msp) if "DOOR" in (e.dxf.name or ""))
check("doors inside a nested plan block", found, 6)

print()
print("=" * 88)
if FAILS:
    print(f"FAILED {len(FAILS)}  (passed {PASSES[0]})")
    for f in FAILS:
        print("  - " + f)
    sys.exit(1)
print(f"ALL {PASSES[0]} ACCURACY CHECKS PASSED")
