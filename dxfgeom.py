#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dxfgeom.py -- read a drawing from its SHAPES, for drawings that carry no
tags, no dimensions, no useful layers and no text at all.

The takeoff tools used to look for one thing in one place: a closed
LWPOLYLINE sitting directly in modelspace, sized in millimetres. Real
drawings put the same column inside a block, or draw it with four separate
lines, or as a legacy POLYLINE, or as a hatch with no outline -- and may be
drawn in metres or feet. Every one of those returned nothing.

This module is the shared answer to that. It has four jobs:

  1. FLATTEN   -- walk into block references so geometry inside a block is
                  seen exactly like geometry in modelspace.
  2. NORMALISE -- turn every kind of closed region (lwpolyline, polyline,
                  circle, ellipse, hatch, solid, four loose lines) into one
                  Region type with a bulge-accurate area.
  3. SCALE     -- work out what one drawing unit means, from the geometry
                  itself, because $INSUNITS is very often wrong.
  4. RECOGNISE -- find columns by the fact that one small shape REPEATS on a
                  regular lattice. That test needs no text and no layers, and
                  it is scale-free, so it works on a 2.2 m house grid and a
                  15 m shed grid alike.

Everything returns a confidence and a human-readable note, because a takeoff
number that feeds a BOQ must say how much to trust it.
"""

from __future__ import annotations

import math
from collections import Counter, defaultdict
from dataclasses import dataclass, field

# ======================================================================
# 1. FLATTEN -- see inside blocks
# ======================================================================
# A drawing that reuses one COL block 40 times has 40 INSERTs in modelspace
# and zero rectangles. Every tool that iterated modelspace directly saw
# nothing. virtual_entities() renders a block reference into real geometry,
# already transformed by the insert's position, rotation and scale.

MAX_BLOCK_DEPTH = 8


def iter_entities(container, max_depth: int = MAX_BLOCK_DEPTH, _depth: int = 0):
    """Yield every entity in `container`, descending into block references.

    Nested blocks are expanded recursively. `max_depth` stops a self-
    referencing block from looping forever; a malformed INSERT is skipped
    rather than aborting the whole read.
    """
    for e in container:
        if e.dxftype() == "INSERT":
            if _depth >= max_depth:
                continue
            try:
                sub = list(e.virtual_entities())
            except Exception:
                continue
            yield from iter_entities(sub, max_depth, _depth + 1)
        else:
            yield e


def iter_inserts(container, max_depth: int = MAX_BLOCK_DEPTH, _depth: int = 0):
    """Yield (insert_entity, depth) for every block reference, nested included.

    Counting doors or fixtures means counting the references themselves, not
    their exploded geometry, so this is kept separate from iter_entities.
    """
    for e in container:
        if e.dxftype() != "INSERT":
            continue
        yield e, _depth
        if _depth < max_depth:
            try:
                sub = list(e.virtual_entities())
            except Exception:
                continue
            yield from iter_inserts(sub, max_depth, _depth + 1)


# ======================================================================
# 2. NORMALISE -- one Region type for every kind of closed shape
# ======================================================================

@dataclass
class Region:
    """A closed area, whatever entity it came from. Units are raw drawing units."""
    points: list                      # [(x, y), ...] outline, arcs flattened
    area: float                       # bulge-accurate, always positive
    kind: str                         # originating dxftype, for reporting
    layer: str = ""
    cx: float = 0.0
    cy: float = 0.0
    w: float = 0.0
    h: float = 0.0

    def bbox_fill(self) -> float:
        """How much of its bounding box the shape fills: 1.0 for a rectangle,
        ~0.79 for a circle. Separates solid column sections from L-shapes."""
        box = self.w * self.h
        return (self.area / box) if box > 0 else 0.0


def _finish(points, area, kind, layer) -> "Region | None":
    if len(points) < 3 or area <= 0:
        return None
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    w, h = max(xs) - min(xs), max(ys) - min(ys)
    if w <= 0 or h <= 0:
        return None
    return Region(points, area, kind, layer,
                  (max(xs) + min(xs)) / 2, (max(ys) + min(ys)) / 2, w, h)


def shoelace(points) -> float:
    """Unsigned polygon area."""
    a = 0.0
    n = len(points)
    for i in range(n):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % n]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def bulge_area(points, bulges) -> float:
    """Polygon area including the circular segments a bulge adds or removes.

    A bulge is tan(sweep/4). Dropping it -- as a plain shoelace does -- can be
    20% out on a plan with a bay window or a curved boundary.
    """
    base = 0.0
    n = len(points)
    extra = 0.0
    for i in range(n):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % n]
        base += x1 * y2 - x2 * y1
        b = bulges[i] if i < len(bulges) else 0.0
        if not b:
            continue
        chord = math.hypot(x2 - x1, y2 - y1)
        if chord < 1e-9:
            continue
        theta = 4.0 * math.atan(b)
        sin_half = math.sin(theta / 2.0)
        if abs(sin_half) < 1e-12:
            continue
        r = chord / (2.0 * sin_half)
        extra += (r * r / 2.0) * (theta - math.sin(theta))
    return abs(base / 2.0 + extra)


def _arc_points(cx, cy, r, a0, a1, ccw=True, steps=24):
    if ccw:
        while a1 <= a0:
            a1 += 2 * math.pi
    else:
        while a1 >= a0:
            a1 -= 2 * math.pi
    return [(cx + r * math.cos(a0 + (a1 - a0) * i / steps),
             cy + r * math.sin(a0 + (a1 - a0) * i / steps))
            for i in range(steps + 1)]


def _lwpolyline_region(e):
    try:
        pts = list(e.get_points())
    except Exception:
        return None
    if len(pts) < 3:
        return None
    xy = [(p[0], p[1]) for p in pts]
    bulges = [p[4] if len(p) > 4 else 0.0 for p in pts]

    closed = _is_closed(e)
    if not closed and len(xy) > 3:
        # Very common: the draughtsman repeated the first point instead of
        # setting the closed flag. Geometrically it is shut, so treat it so.
        if math.dist(xy[0], xy[-1]) < 1e-6:
            xy = xy[:-1]
            bulges = bulges[:-1]
            closed = True
    if not closed:
        return None
    return _finish(xy, bulge_area(xy, bulges), "LWPOLYLINE", e.dxf.layer or "")


def _is_closed(e) -> bool:
    """Closed flag across entity types.

    LWPOLYLINE exposes `closed`, legacy POLYLINE exposes `is_closed`. Reading
    only the first is why R12-era drawings returned nothing at all.
    """
    for attr in ("is_closed", "closed"):
        v = getattr(e, attr, None)
        if isinstance(v, bool):
            return v
    return False


def _polyline_region(e):
    """Legacy POLYLINE -- what every R12-era drawing uses."""
    try:
        if getattr(e, "get_mode", lambda: "")() not in ("AcDb2dPolyline", ""):
            return None                     # 3d polylines and meshes are not areas
        verts = list(e.vertices)
    except Exception:
        return None
    xy, bulges = [], []
    for v in verts:
        try:
            xy.append((v.dxf.location.x, v.dxf.location.y))
            bulges.append(getattr(v.dxf, "bulge", 0.0) or 0.0)
        except Exception:
            return None
    if len(xy) < 3:
        return None
    if not _is_closed(e):
        if math.dist(xy[0], xy[-1]) > 1e-6:
            return None
        xy, bulges = xy[:-1], bulges[:-1]
    return _finish(xy, bulge_area(xy, bulges), "POLYLINE", e.dxf.layer or "")


def _circle_region(e):
    r = float(e.dxf.radius)
    if r <= 0:
        return None
    c = e.dxf.center
    pts = _arc_points(c.x, c.y, r, 0.0, 2 * math.pi, steps=32)[:-1]
    return _finish(pts, math.pi * r * r, "CIRCLE", e.dxf.layer or "")


def _ellipse_region(e):
    try:
        if abs(abs(e.dxf.end_param - e.dxf.start_param) - 2 * math.pi) > 1e-6:
            return None                     # an elliptical arc is not a region
        pts = [(p.x, p.y) for p in e.flattening(distance=0.5)]
    except Exception:
        return None
    return _finish(pts, shoelace(pts), "ELLIPSE", e.dxf.layer or "")


def _hatch_regions(e):
    """A hatch is often the ONLY thing drawn -- column poche or floor fill."""
    out = []
    layer = e.dxf.layer or ""
    try:
        paths = list(e.paths)
    except Exception:
        return out
    for p in paths:
        pts, bulges = [], []
        try:
            if p.path_type_flags & 2:                       # polyline path
                for v in p.vertices:
                    pts.append((v[0], v[1]))
                    bulges.append(v[2] if len(v) > 2 else 0.0)
            else:                                           # edge path
                for edge in p.edges:
                    t = type(edge).__name__
                    if t == "LineEdge":
                        pts.append((edge.start[0], edge.start[1]))
                    elif t == "ArcEdge":
                        pts += _arc_points(edge.center[0], edge.center[1],
                                           edge.radius,
                                           math.radians(edge.start_angle),
                                           math.radians(edge.end_angle),
                                           getattr(edge, "ccw", True), 12)
                    elif t == "EllipseEdge":
                        pts.append((edge.center[0], edge.center[1]))
                    elif t == "SplineEdge":
                        pts += [(c[0], c[1]) for c in edge.control_points]
        except Exception:
            continue
        if len(pts) < 3:
            continue
        area = bulge_area(pts, bulges) if any(bulges) else shoelace(pts)
        r = _finish(pts, area, "HATCH", layer)
        if r:
            out.append(r)
    return out


def _solid_region(e):
    try:
        corners = [e.dxf.vtx0, e.dxf.vtx1, e.dxf.vtx3, e.dxf.vtx2]
    except Exception:
        return None
    pts = []
    for c in corners:
        xy = (c.x, c.y)
        if not pts or math.dist(pts[-1], xy) > 1e-9:
            pts.append(xy)
    return _finish(pts, shoelace(pts), e.dxftype(), e.dxf.layer or "")


def _rects_from_lines(entities, snap=1.0, max_side=None):
    """Recover rectangles drawn as four separate LINE entities.

    Endpoints are snapped to a tolerance and axis-aligned segments are paired
    into boxes. Only small boxes are considered, so this cannot turn a whole
    floor outline into a candidate.
    """
    horiz, vert = defaultdict(list), defaultdict(list)
    for e in entities:
        if e.dxftype() != "LINE":
            continue
        a, b = e.dxf.start, e.dxf.end
        x1, y1, x2, y2 = a.x, a.y, b.x, b.y
        if max_side and max(abs(x2 - x1), abs(y2 - y1)) > max_side:
            continue
        if abs(y2 - y1) <= snap and abs(x2 - x1) > snap:
            key = round(y1 / snap)
            horiz[key].append((min(x1, x2), max(x1, x2), (y1 + y2) / 2,
                               e.dxf.layer or ""))
        elif abs(x2 - x1) <= snap and abs(y2 - y1) > snap:
            key = round(x1 / snap)
            vert[key].append((min(y1, y2), max(y1, y2), (x1 + x2) / 2,
                              e.dxf.layer or ""))

    out = []
    seen = set()
    for _, hs in horiz.items():
        for (hx0, hx1, hy, layer) in hs:
            for _, hs2 in horiz.items():
                for (gx0, gx1, gy, _l2) in hs2:
                    if gy <= hy or abs(gx0 - hx0) > snap or abs(gx1 - hx1) > snap:
                        continue
                    # need the two vertical sides to exist as well
                    if not (_has_vert(vert, hx0, hy, gy, snap)
                            and _has_vert(vert, hx1, hy, gy, snap)):
                        continue
                    key = (round(hx0 / snap), round(hy / snap),
                           round(hx1 / snap), round(gy / snap))
                    if key in seen:
                        continue
                    seen.add(key)
                    pts = [(hx0, hy), (hx1, hy), (hx1, gy), (hx0, gy)]
                    r = _finish(pts, shoelace(pts), "LINES", layer)
                    if r:
                        out.append(r)
    return out


def _has_vert(vert, x, y0, y1, snap):
    for (vy0, vy1, vx, _l) in vert.get(round(x / snap), []):
        if abs(vx - x) <= snap and vy0 <= y0 + snap and vy1 >= y1 - snap:
            return True
    return False


def regions(entities, recover_lines: bool = True, max_line_rect: float | None = None):
    """Every closed region in `entities`, whatever entity type drew it."""
    ents = list(entities)
    out = []
    for e in ents:
        t = e.dxftype()
        try:
            if t == "LWPOLYLINE":
                r = _lwpolyline_region(e)
                if r:
                    out.append(r)
            elif t == "POLYLINE":
                r = _polyline_region(e)
                if r:
                    out.append(r)
            elif t == "CIRCLE":
                r = _circle_region(e)
                if r:
                    out.append(r)
            elif t == "ELLIPSE":
                r = _ellipse_region(e)
                if r:
                    out.append(r)
            elif t == "HATCH":
                out += _hatch_regions(e)
            elif t in ("SOLID", "TRACE", "3DFACE"):
                r = _solid_region(e)
                if r:
                    out.append(r)
        except Exception:
            continue                       # one bad entity must not stop a takeoff
    if recover_lines:
        out += _rects_from_lines(ents, max_side=max_line_rect)
    return out


# ======================================================================
# 3. SCALE -- what does one drawing unit mean?
# ======================================================================
# $INSUNITS is unreliable: this repo's own test drawing declares 6 (metres)
# while being drawn in millimetres. So the header is treated as a hint and
# the geometry decides. Buildings are 3-500 m across and columns are
# 0.15-1.5 m, which is enough to pin the scale without reading any text.

_UNIT_MM = {1: 25.4, 2: 304.8, 4: 1.0, 5: 10.0, 6: 1000.0, 8: 2.54e-5,
            9: 0.0254, 10: 914.4, 13: 0.001, 14: 100.0, 15: 10000.0}
_UNIT_NAME = {1: "inches", 2: "feet", 4: "mm", 5: "cm", 6: "m", 10: "yards",
              13: "microns", 14: "dm", 15: "dam"}

# Candidate scales tried against the geometry, mm per drawing unit, with a
# small prior. Several units can be equally plausible for one building (12 m
# in centimetres and 30 m in inches both look like buildings), so the prior
# breaks the tie toward what architectural DXFs actually use.
_CANDIDATES = [("mm", 1.0, 0.30), ("m", 1000.0, 0.20), ("cm", 10.0, 0.10),
               ("ft", 304.8, 0.05), ("inch", 25.4, 0.02)]


@dataclass
class Scale:
    mm_per_unit: float = 1.0
    unit: str = "mm"
    source: str = "assumed"
    confidence: str = "low"
    note: str = ""

    def to_m(self, v: float) -> float:
        return v * self.mm_per_unit / 1000.0

    def area_to_m2(self, v: float) -> float:
        return v * (self.mm_per_unit ** 2) / 1_000_000.0


def _plausibility(extent_units, small_units, mm_per_unit):
    """Score a candidate scale: is the building a building, and are its small
    repeated shapes column-sized?"""
    score = 0.0
    span_m = extent_units * mm_per_unit / 1000.0
    if 3.0 <= span_m <= 500.0:
        score += 2.0
    elif 1.0 <= span_m <= 2000.0:
        score += 0.5
    if small_units:
        small_m = small_units * mm_per_unit / 1000.0
        if 0.15 <= small_m <= 1.5:
            score += 2.0
        elif 0.05 <= small_m <= 3.0:
            score += 0.5
    return score


def detect_scale(doc, regs=None, extents=None) -> Scale:
    """Work out the drawing's unit from its geometry, cross-checked against
    $INSUNITS. Geometry wins when the two disagree -- headers are often stale."""
    header_mm, header_name = None, None
    try:
        code = int(doc.header.get("$INSUNITS", 0) or 0)
        if code in _UNIT_MM:
            header_mm = _UNIT_MM[code]
            header_name = _UNIT_NAME.get(code, str(code))
    except Exception:
        pass

    if extents is None and regs:
        # Measured across the shapes' full bounding boxes, not their centres:
        # a drawing with a single region has zero centre-spread but is still
        # perfectly measurable.
        xs, ys = [], []
        for r in regs:
            xs += [r.cx - r.w / 2, r.cx + r.w / 2]
            ys += [r.cy - r.h / 2, r.cy + r.h / 2]
        extents = max(max(xs) - min(xs), max(ys) - min(ys)) if xs else 0.0
    extents = extents or 0.0

    # The modal small-shape size is the best scale probe available without text.
    small = None
    if regs:
        sizes = sorted(min(r.w, r.h) for r in regs if r.w > 0 and r.h > 0)
        if sizes:
            small = sizes[len(sizes) // 4]        # lower quartile

    if extents <= 0:
        if header_mm:
            return Scale(header_mm, header_name, "$INSUNITS", "low",
                         "empty drawing; used the header unit")
        return Scale(1.0, "mm", "assumed", "low", "no geometry to measure")

    scored = []
    for name, mm, prior in _CANDIDATES:
        raw = _plausibility(extents, small, mm)
        s = raw + prior
        # The header is a hint, not an authority: it nudges a tie but cannot
        # outvote geometry that clearly says otherwise.
        if header_mm and abs(header_mm - mm) < 1e-9:
            s += 0.5
        scored.append((s, raw, name, mm))
    scored.sort(key=lambda t: -t[0])
    best_score, best_raw, best_name, best_mm = scored[0]

    if best_raw <= 0:
        if header_mm:
            return Scale(header_mm, header_name, "$INSUNITS", "low",
                         "geometry inconclusive; used the header unit")
        return Scale(1.0, "mm", "assumed", "low",
                     "geometry inconclusive; assumed millimetres")

    if header_mm and abs(header_mm - best_mm) < 1e-9:
        return Scale(best_mm, best_name, "$INSUNITS + geometry", "high",
                     f"header and geometry agree on {best_name}")
    if header_mm:
        return Scale(best_mm, best_name, "geometry", "medium",
                     f"header says {header_name} but the geometry reads as "
                     f"{best_name}; used {best_name}")
    return Scale(best_mm, best_name, "geometry", "medium",
                 f"no $INSUNITS; geometry reads as {best_name}")


# ======================================================================
# 4. RECOGNISE -- columns, by repetition on a lattice
# ======================================================================
# With no tags and no layers the only reliable evidence that a small shape is
# a column is that THE SAME shape repeats on a regular lattice. That test is
# scale-free, so it replaces the old fixed 2.5-11 m grid window, the >=8
# column minimum and the >=100 m2 footprint gate -- all of which silently
# rejected small houses, tight grids and wide-span sheds.


def _axis_lines(values, tol):
    """Collapse coordinates into distinct grid lines within `tol`."""
    if not values:
        return []
    vs = sorted(values)
    lines, cur = [], [vs[0]]
    for v in vs[1:]:
        if v - cur[-1] <= tol:
            cur.append(v)
        else:
            lines.append(sum(cur) / len(cur))
            cur = [v]
    lines.append(sum(cur) / len(cur))
    return lines


def cluster_knn(centres, k=4):
    """Group points with a mutual k-nearest-neighbour graph. Returns index lists.

    Deliberately NOT a fixed distance: one drawing holds a structural grid at
    6 m and a legend block at 1.5 m, so any single threshold either shatters
    the grid or swallows the legend. Two points join only when each is among
    the other's k nearest, which adapts to whatever local spacing each region
    has and keeps a dense blob from chaining into a sparse one.
    """
    n = len(centres)
    if n <= 1:
        return [list(range(n))]
    k = max(1, min(k, n - 1))

    near = []
    for i, (x, y) in enumerate(centres):
        d = sorted(((math.hypot(x - centres[j][0], y - centres[j][1]), j)
                    for j in range(n) if j != i))[:k]
        near.append({j for _dist, j in d})

    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    for i in range(n):
        for j in near[i]:
            if i in near[j]:                 # mutual -- not merely one-way
                ra, rb = find(i), find(j)
                if ra != rb:
                    parent[rb] = ra

    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(i)
    return list(groups.values())


def cluster_points(centres, thresh):
    """Group points within `thresh` of a neighbour. Returns index lists."""
    n = len(centres)
    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    # Bucketed by a grid of side `thresh` so this stays near-linear.
    buckets = defaultdict(list)
    for i, (x, y) in enumerate(centres):
        buckets[(int(x // thresh), int(y // thresh))].append(i)
    for (bx, by), idx in buckets.items():
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for j in buckets.get((bx + dx, by + dy), ()):
                    for i in idx:
                        if i < j and math.dist(centres[i], centres[j]) <= thresh:
                            union(i, j)
    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(i)
    return list(groups.values())


def _line_members(values, tol):
    """Group coordinates into grid lines, returning (line_value, [indices])."""
    order = sorted(range(len(values)), key=lambda i: values[i])
    lines, cur = [], [order[0]]
    for i in order[1:]:
        if values[i] - values[cur[-1]] <= tol:
            cur.append(i)
        else:
            lines.append((sum(values[j] for j in cur) / len(cur), cur))
            cur = [i]
    lines.append((sum(values[j] for j in cur) / len(cur), cur))
    return lines


def refine_lattice(centres, tol, min_per_line=2):
    """Keep only the points that sit on a *shared* grid line on both axes.

    This is what separates a column grid from everything else the same size.
    A run of six parking bays shares one y line but each bay owns its own x
    line, so no bay survives; scattered sanitary fixtures share neither. Real
    columns sit on rows and gridlines that several of them use, so they all
    survive. Returns the surviving indices.
    """
    if len(centres) < 3:
        return list(range(len(centres)))
    keep = set(range(len(centres)))
    for _ in range(2):                      # one refinement pass, then confirm
        idx = sorted(keep)
        if len(idx) < 3:
            break
        xs = [centres[i][0] for i in idx]
        ys = [centres[i][1] for i in idx]
        good_x, good_y = set(), set()
        for _v, members in _line_members(xs, tol):
            if len(members) >= min_per_line:
                good_x.update(idx[m] for m in members)
        for _v, members in _line_members(ys, tol):
            if len(members) >= min_per_line:
                good_y.update(idx[m] for m in members)
        nxt = good_x & good_y
        if nxt == keep:
            break
        keep = nxt
    return sorted(keep)


def lattice_fit(centres, tol):
    """1.0 when the points form a full rectangular lattice, ~0 when scattered.

    A r x c grid of N points occupies exactly c distinct x lines and r
    distinct y lines, so N / (nx*ny) == 1. Scattered points need nearly N
    lines on each axis, driving the ratio toward 1/N.
    """
    if len(centres) < 3:
        return 0.0, 1, 1
    nx = len(_axis_lines([c[0] for c in centres], tol))
    ny = len(_axis_lines([c[1] for c in centres], tol))
    if nx == 0 or ny == 0:
        return 0.0, nx, ny
    return len(centres) / float(nx * ny), nx, ny


def median_spacing(centres, min_gap=1e-9):
    """Median nearest-neighbour distance -- the structural grid, in raw units."""
    if len(centres) < 2:
        return 0.0
    # Bucketed so a big drawing does not go quadratic.
    best = []
    pts = centres
    n = len(pts)
    step = max(1, n // 400)
    for i in range(0, n, step):
        x, y = pts[i]
        d_best = None
        for j in range(n):
            if i == j:
                continue
            d = math.hypot(x - pts[j][0], y - pts[j][1])
            if d > min_gap and (d_best is None or d < d_best):
                d_best = d
        if d_best is not None:
            best.append(d_best)
    if not best:
        return 0.0
    best.sort()
    return best[len(best) // 2]


@dataclass
class ColumnFamily:
    size: tuple                 # (w, h) rounded, raw units
    members: list = field(default_factory=list)
    fit: float = 0.0
    rows: int = 0
    cols: int = 0
    spacing: float = 0.0

    @property
    def n(self):
        return len(self.members)


@dataclass
class ColumnResult:
    families: list = field(default_factory=list)
    scale: Scale = field(default_factory=Scale)
    confidence: str = "none"
    note: str = ""
    rejected: list = field(default_factory=list)

    @property
    def total(self):
        return sum(f.n for f in self.families)

    def schedule(self, round_to_mm: int = 25):
        """Counter of (w_mm, h_mm) -> count over every accepted column.

        Sizes are read per member, because one grid legitimately carries
        several sections, and rounded so that a 449.7 and a 450.2 rectangle
        are one entry rather than two.
        """
        c = Counter()
        mm = self.scale.mm_per_unit
        step = max(1, round_to_mm)
        for f in self.families:
            for r in f.members:
                w = int(round(r.w * mm / step)) * step
                h = int(round(r.h * mm / step)) * step
                c[(w, h)] += 1
        return c

    def grid_m(self):
        sp = [f.spacing for f in self.families if f.spacing > 0]
        if not sp:
            return 0.0
        sp.sort()
        return self.scale.to_m(sp[len(sp) // 2])


# A column section is 150-1500 mm a side and roughly chunky, not a sliver.
COL_MIN_MM, COL_MAX_MM, COL_MAX_ASPECT = 150.0, 1500.0, 4.0
# Below this lattice fit a repeated shape is furniture, not structure.
LATTICE_MIN = 0.34
# A column grid is spaced several times the column's own width apart.
MIN_PITCH_RATIO = 3.0


def detect_columns(regs, scale: Scale, size_tol_mm: float = 25.0,
                   lattice_min: float = LATTICE_MIN) -> ColumnResult:
    """Find the column grid among the closed regions.

    Candidates are every column-proportioned shape, REGARDLESS of size: a real
    grid mixes 450x450, 300x450 and 300x300 sections on the same lines, so
    grouping by size before testing the lattice makes every section look
    scattered while a crowd of identical fixtures looks like a grid.

    So: take all candidates, split them into spatial groups, and ask of each
    group whether its members share gridlines. Sizes are reported afterwards.
    """
    mm = scale.mm_per_unit

    cands = []
    for r in regs:
        w_mm, h_mm = r.w * mm, r.h * mm
        if not (COL_MIN_MM <= w_mm <= COL_MAX_MM and COL_MIN_MM <= h_mm <= COL_MAX_MM):
            continue
        if max(w_mm, h_mm) / max(1e-9, min(w_mm, h_mm)) > COL_MAX_ASPECT:
            continue
        if r.bbox_fill() < 0.55:      # a hollow L or a ring is not a section
            continue
        cands.append(r)

    # Collapse shapes stacked at one spot: a legend or a detail draws many
    # boxes on top of each other, which would read as a dense column cluster.
    cell = max((size_tol_mm / mm) if mm else size_tol_mm, 1e-9) * 2
    cands = _dedupe(cands, cell)
    if len(cands) < 3:
        return ColumnResult([], scale, "none",
                            "fewer than three column-sized shapes in the drawing",
                            [("column-sized shapes", len(cands))])

    centres = [(r.cx, r.cy) for r in cands]
    clusters = cluster_knn(centres, k=4)

    families, rejected = [], []
    for cluster in clusters:
        members = [cands[i] for i in cluster]
        if len(members) < 3:
            rejected.append(("isolated shape", len(members)))
            continue
        _judge(members, mm, families, rejected, lattice_min)

    families.sort(key=lambda f: -f.n)
    if not families:
        return ColumnResult([], scale, "none",
                            "no repeated column-sized shape sits on a grid",
                            rejected)

    best_fit = max(f.fit for f in families)
    if best_fit >= 0.8 and scale.confidence in ("high", "medium"):
        conf = "high"
    elif best_fit >= 0.5:
        conf = "medium"
    else:
        conf = "low"
    note = (f"{sum(f.n for f in families)} columns on {len(families)} "
            f"grid(s); lattice fit {best_fit:.2f}; units {scale.unit} "
            f"({scale.source})")
    return ColumnResult(families, scale, conf, note, rejected)


def _judge(members, mm, families, rejected, lattice_min):
    """Decide whether one spatial group of same-sized shapes is a column grid."""
    centres = [(r.cx, r.cy) for r in members]
    spacing = median_spacing(centres)
    # Tolerance for "same gridline": generous enough for a column nudged off
    # centre, tight enough not to merge two adjacent bays into one line.
    tol = min(max(max(r.w for r in members) * 1.5, spacing * 0.12),
              max(spacing * 0.35, 1e-9))

    # Drop members that share no gridline with the rest of their group.
    keep = refine_lattice(centres, tol)
    if len(keep) != len(members):
        members = [members[i] for i in keep]
        centres = [centres[i] for i in keep]
        spacing = median_spacing(centres) or spacing
    if len(members) < 3:
        rejected.append(("scattered, not on a shared grid", len(keep)))
        return

    fit, nx, ny = lattice_fit(centres, tol)
    avg_w = sum(r.w for r in members) / len(members)
    avg_h = sum(r.h for r in members) / len(members)
    fam = ColumnFamily((avg_w, avg_h), members, fit, ny, nx, spacing)

    # Structure stands far apart relative to its own section; a legend or a
    # schedule table stacks same-sized boxes nearly shoulder to shoulder.
    biggest = max(max(r.w, r.h) for r in members)
    if spacing > 0 and biggest > 0 and spacing / biggest < MIN_PITCH_RATIO:
        rejected.append((f"boxes {int(round(biggest * mm))}mm at "
                         f"{int(round(spacing * mm))}mm pitch: packed too "
                         f"tightly to be structure", len(members)))
        return
    # A lattice needs depth on both axes. A single row is a queue of parking
    # bays far more often than it is a line of columns.
    if fit >= lattice_min and nx >= 2 and ny >= 2:
        families.append(fam)
    else:
        rejected.append((f"{int(round(avg_w * mm))}x{int(round(avg_h * mm))} "
                         f"x{len(members)}: lattice fit {fit:.2f}, "
                         f"{nx} x {ny} lines", len(members)))


def _dedupe(regs, cell):
    seen, out = set(), []
    for r in regs:
        key = (round(r.cx / cell), round(r.cy / cell))
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


# ======================================================================
# 5. LINEAR ELEMENTS -- count walls and beams, not line entities
# ======================================================================
# A beam or a wall is drawn as its two faces. Counting LINE entities counts
# each one twice and doubles every plaster, paint and concrete quantity that
# follows from it.


def segments(entities, layer=None):
    """(x1, y1, x2, y2, layer) for every straight run, polylines included."""
    out = []
    for e in entities:
        lay = e.dxf.layer or ""
        if layer is not None and lay != layer:
            continue
        t = e.dxftype()
        if t == "LINE":
            a, b = e.dxf.start, e.dxf.end
            out.append((a.x, a.y, b.x, b.y, lay))
        elif t == "LWPOLYLINE":
            try:
                pts = [(p[0], p[1]) for p in e.get_points()]
            except Exception:
                continue
            if _is_closed(e) and len(pts) > 2:
                pts = pts + [pts[0]]
            for i in range(len(pts) - 1):
                out.append((pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1], lay))
        elif t == "POLYLINE":
            try:
                pts = [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices]
            except Exception:
                continue
            if _is_closed(e) and len(pts) > 2:
                pts = pts + [pts[0]]
            for i in range(len(pts) - 1):
                out.append((pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1], lay))
    return out


def merge_collinear(segs, tol=1.0):
    """Join segments that continue each other into one run.

    A beam split at every column it crosses is still one beam; without this
    it is counted once per span.
    """
    axis = defaultdict(list)
    others = []
    for (x1, y1, x2, y2, lay) in segs:
        if abs(y2 - y1) <= tol and abs(x2 - x1) > tol:
            axis[("h", round(((y1 + y2) / 2) / tol), lay)].append(
                (min(x1, x2), max(x1, x2), (y1 + y2) / 2))
        elif abs(x2 - x1) <= tol and abs(y2 - y1) > tol:
            axis[("v", round(((x1 + x2) / 2) / tol), lay)].append(
                (min(y1, y2), max(y1, y2), (x1 + x2) / 2))
        else:
            others.append((x1, y1, x2, y2, lay))

    runs = []
    for (kind, _k, lay), spans in axis.items():
        spans.sort()
        c0, c1, fixed = spans[0]
        for (a, b, f) in spans[1:]:
            if a <= c1 + tol:
                c1 = max(c1, b)
            else:
                runs.append((kind, c0, c1, fixed, lay))
                c0, c1, fixed = a, b, f
        runs.append((kind, c0, c1, fixed, lay))

    out = []
    for (kind, a, b, f, lay) in runs:
        if kind == "h":
            out.append((a, f, b, f, lay))
        else:
            out.append((f, a, f, b, lay))
    return out + others


def pair_faces(segs, max_sep, tol=1.0):
    """Collapse the two faces of a wall or beam into one centre-line.

    Two parallel runs closer than `max_sep` that overlap along their length
    are the same element seen from both sides. Returns
    (centre_segments, thickness_by_segment).
    """
    runs = merge_collinear(segs, tol)
    used = set()
    centres, thicks = [], []

    def orient(s):
        x1, y1, x2, y2, _l = s
        if abs(y2 - y1) <= tol:
            return "h"
        if abs(x2 - x1) <= tol:
            return "v"
        return None

    for i, s in enumerate(runs):
        if i in used:
            continue
        o = orient(s)
        if o is None:
            continue
        x1, y1, x2, y2, lay = s
        partner, best = None, None
        for j, t in enumerate(runs):
            if j == i or j in used or orient(t) != o:
                continue
            ux1, uy1, ux2, uy2, tlay = t
            if tlay != lay:
                continue
            if o == "h":
                sep = abs(uy1 - y1)
                overlap = min(max(x1, x2), max(ux1, ux2)) - max(min(x1, x2), min(ux1, ux2))
            else:
                sep = abs(ux1 - x1)
                overlap = min(max(y1, y2), max(uy1, uy2)) - max(min(y1, y2), min(uy1, uy2))
            if sep <= tol or sep > max_sep or overlap <= tol:
                continue
            if best is None or sep < best:
                best, partner = sep, j
        if partner is None:
            centres.append(s)
            thicks.append(0.0)
            used.add(i)
        else:
            t = runs[partner]
            used.add(i)
            used.add(partner)
            if o == "h":
                y = (y1 + t[1]) / 2
                a = max(min(x1, x2), min(t[0], t[2]))
                b = min(max(x1, x2), max(t[0], t[2]))
                centres.append((a, y, b, y, lay))
            else:
                x = (x1 + t[0]) / 2
                a = max(min(y1, y2), min(t[1], t[3]))
                b = min(max(y1, y2), max(t[1], t[3]))
                centres.append((x, a, x, b, lay))
            thicks.append(best)
    return centres, thicks


def seg_length(s):
    return math.hypot(s[2] - s[0], s[3] - s[1])


# ======================================================================
# 6. CONTAINMENT -- do not add a building to the plot it stands on
# ======================================================================


def point_in_polygon(pt, poly) -> bool:
    x, y = pt
    inside = False
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]
        if (y1 > y) != (y2 > y):
            xin = (x2 - x1) * (y - y1) / (y2 - y1 + 1e-30) + x1
            if x < xin:
                inside = not inside
    return inside


def net_areas(regs):
    """Resolve nesting: a region inside another is a hole, not extra area.

    Returns (net_area, outer_regions, nested_count) in raw units squared.
    Shapes at even nesting depth add, odd depth subtracts, which is how a
    courtyard inside a footprint inside a plot comes out right.
    """
    ordered = sorted(regs, key=lambda r: -r.area)
    depth = {}
    for i, r in enumerate(ordered):
        d = 0
        for j, bigger in enumerate(ordered):
            if j >= i:
                break
            if bigger.area > r.area and point_in_polygon((r.cx, r.cy), bigger.points):
                d += 1
        depth[id(r)] = d
    net = 0.0
    outer, nested = [], 0
    for r in ordered:
        if depth[id(r)] % 2 == 0:
            net += r.area
            if depth[id(r)] == 0:
                outer.append(r)
        else:
            net -= r.area
            nested += 1
    return max(net, 0.0), outer, nested
