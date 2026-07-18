#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ArchTools.py  --  Architect's drawing & estimate toolkit
========================================================

A single-file command-line toolkit for reading AutoCAD DXF drawings and
Excel (.xls/.xlsx) BOQ / estimate sheets, and for doing the everyday
quantity-takeoff and estimating maths an architect needs from a plan.

Built for an architect who works mostly from working drawings and BOQs.
Every command prints a clean table you can paste into a report.

------------------------------------------------------------------------
QUICK START
------------------------------------------------------------------------
    python ArchTools.py --help                 # list every tool
    python ArchTools.py info        plan.dxf   # drawing summary
    python ArchTools.py layers      plan.dxf   # layers + entity counts
    python ArchTools.py text        plan.dxf   # dump all unique text
    python ArchTools.py levels      plan.dxf   # floor / level marks
    python ArchTools.py columns     plan.dxf   # column count by size
    python ArchTools.py colvol      plan.dxf   # column concrete volume
    python ArchTools.py floorcols   plan.dxf   # columns per floor plan
    python ArchTools.py areas       plan.dxf   # closed-polyline areas/layer
    python ArchTools.py lengths     plan.dxf   # line lengths per layer
    python ArchTools.py blocks      plan.dxf   # block (INSERT) counts
    python ArchTools.py schedule    plan.dxf   # door/window WxH schedule
    python ArchTools.py hatch       plan.dxf   # hatch (finish) area/layer
    python ArchTools.py dims        plan.dxf   # dimension values
    python ArchTools.py extents     plan.dxf   # drawing size / sheet box
    python ArchTools.py export      plan.dxf   # entities -> CSV
    python ArchTools.py purge       plan.dxf   # empty/junk layer report

    python ArchTools.py xls         boq.xls            # list sheets
    python ArchTools.py xlsdump     boq.xls            # dump a sheet
    python ArchTools.py find        boq.xls column     # search BOQ items
    python ArchTools.py compare     a.xls b.xls column # compare two BOQs

    # Estimating calculators (no file needed):
    python ArchTools.py calc-concrete  --l 5 --b 0.3 --h 3 --n 4
    python ArchTools.py calc-rebar     --dia 16 --len 12 --n 40
    python ArchTools.py calc-brick     --area 50 --thk 230
    python ArchTools.py calc-plaster   --area 120 --thk 12
    python ArchTools.py calc-paint     --area 200 --coats 2

------------------------------------------------------------------------
DEPENDENCIES
------------------------------------------------------------------------
    pip install ezdxf xlrd openpyxl

    ezdxf    -> DXF reading            (required for all DXF commands)
    xlrd     -> old .xls (BIFF/OLE2)   (required for legacy .xls)
    openpyxl -> .xlsx                  (required for modern .xlsx)

NOTE on .dwg files: AutoCAD .dwg is a closed binary and cannot be read
directly. Convert to DXF first with the free ODA File Converter
(https://www.opendesign.com/guestfiles/oda_file_converter), then use the
DXF commands here.  See:  python ArchTools.py dwghelp
"""

import argparse
import csv
import os
import re
import sys
import math
from collections import Counter, defaultdict

# ----------------------------------------------------------------------
# Soft imports -- only fail when a command actually needs the library.
# ----------------------------------------------------------------------
try:
    import ezdxf
except ImportError:
    ezdxf = None

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


# Default floor-to-floor level scheme (metres) used by column-volume tools.
# Override on the command line with --levels "0:1.0,1:8.0,..." if needed.
# These are the Meridian M1 production-block levels discovered in this project.
DEFAULT_LEVELS = {
    "ROAD":     0.00,
    "GROUND":   1.00,
    "FIRST":    8.00,
    "SECOND":  15.00,
    "THIRD":   19.50,
    "FOURTH":  24.00,
    "TERRACE": 27.50,
    "LIFTTOP": 30.50,
    "PARAPET": 30.95,
}

FLOOR_ORDER = ["GROUND", "FIRST", "SECOND", "THIRD", "FOURTH", "TERRACE"]


# Rate library: item -> {unit, rate in Rs}. Ships with rough Indian 2024-25
# ballpark rates the architect edits per city/year (rates drift -- this is a
# knob, never hardcoded into a number). First run writes rates.json into the
# data dir so edits persist and survive app upgrades, like project.json.
# Items whose keys match the auto-takeoff output get priced automatically;
# the rest sit in the file for manual/future lines.
DEFAULT_RATES = {
    "RCC concrete (columns)":  {"unit": "m3",    "rate": 7500},
    "RCC concrete (slab)":     {"unit": "m3",    "rate": 7000},
    "Footing concrete":        {"unit": "m3",    "rate": 6500},
    "Reinforcement steel":     {"unit": "tonne", "rate": 75000},
    "Excavation":              {"unit": "m3",    "rate": 250},
    "Doors":                   {"unit": "nos",   "rate": 9000},
    "Windows":                 {"unit": "nos",   "rate": 7500},
    "Brickwork 230mm":         {"unit": "m3",    "rate": 7200},
    "Internal plaster 12mm":   {"unit": "m2",    "rate": 280},
    "Flooring (vitrified)":    {"unit": "m2",    "rate": 1200},
    "Paint (2 coats)":         {"unit": "m2",    "rate": 90},
}


# ----------------------------------------------------------------------
# Catalog metadata  --  drives the searchable tool list in the web UI.
# command -> (category, friendly title). Commands not listed default to
# category "Other". System commands (project/tools/catalog/dwghelp) are
# hidden from the catalog.
# ----------------------------------------------------------------------
CATALOG = {
    # Structure
    "columns":     ("Structure", "Count columns by size"),
    "colvol":      ("Structure", "Column concrete volume"),
    "floorcols":   ("Structure", "Columns per floor"),
    "colschedule": ("Structure", "Full column schedule (each tag)"),
    "perimeter-columns": ("Structure", "Perimeter vs interior columns"),
    "plinth-columns":    ("Structure", "Plinth (ground-floor) columns"),
    # Plinth suite
    "plinth-colvol":     ("Plinth", "Plinth column concrete volume"),
    "plinth-beams":      ("Plinth", "Plinth-beam concrete + steel"),
    "plinth-area":       ("Plinth", "Plinth area + perimeter"),
    "calc-plinth-beam":  ("Plinth", "Plinth beam concrete (L×B×D×N)"),
    "calc-plinth-fill":  ("Plinth", "Plinth earth filling volume"),
    "calc-plinth-masonry": ("Plinth", "Plinth wall masonry + bricks"),
    "calc-dpc":          ("Plinth", "Damp-proof course (DPC)"),
    "calc-anti-termite": ("Plinth", "Anti-termite treatment area"),
    "calc-plinth-protection": ("Plinth", "Plinth protection apron (PCC)"),
    "calc-pcc-bed":      ("Plinth", "PCC leveling bed volume"),
    # Chajja / sunshade suite
    "chajja":            ("Chajja / Sunshade", "Chajja concrete + steel (from layer)"),
    "calc-chajja":       ("Chajja / Sunshade", "Chajja concrete + steel"),
    "calc-chajja-shutter": ("Chajja / Sunshade", "Chajja shuttering / formwork area"),
    "calc-chajja-plaster": ("Chajja / Sunshade", "Chajja plaster area + mortar"),
    # Concrete + Slab suite
    "concrete":          ("Structure", "Total RCC concrete (auto from drawing)"),
    "slab":              ("Slab", "Auto slab takeoff (from drawing)"),
    "calc-slab":         ("Slab", "Slab concrete + steel"),
    "calc-slab-steel":   ("Slab", "Slab reinforcement (kg/m² or %)"),
    "calc-slab-shutter": ("Slab", "Slab shuttering / formwork"),
    "calc-slab-plaster": ("Slab", "Slab ceiling plaster + mortar"),
    "beams":       ("Structure", "Beams: count + length"),
    "footings":    ("Foundation", "Footings count"),
    "foundation":  ("Foundation", "Foundation takeoff (auto from drawing)"),
    "footing-schedule": ("Foundation", "Footing schedule (each F-tag)"),
    "staircases":  ("Structure", "Staircases + lifts"),
    "circles":     ("Structure", "Circles + radius stats"),
    "centroid":    ("Structure", "Centre point of a layer"),
    # Levels & heights
    "levels":      ("Levels & heights", "Floor levels"),
    "heights":     ("Levels & heights", "Floor-to-floor + building height"),
    "column-spacing": ("Structure", "Column grid spacing"),
    "bbox-layer":  ("Structure", "Footprint / size of a layer"),
    # Areas
    "areas":       ("Areas", "Closed-shape areas per layer"),
    "builtup":     ("Areas", "Built-up footprint area"),
    "room-areas":  ("Areas", "Area of each room on a layer"),
    "layer-area":  ("Areas", "Closed area on one layer"),
    # Openings & fixtures
    "schedule":    ("Openings", "Door / window schedule"),
    "doors":       ("Openings", "Count door blocks"),
    "windows":     ("Openings", "Count window blocks"),
    "fixtures":    ("Openings", "Count fixtures on a layer"),
    "blocks":      ("Openings", "Block / fixture counts"),
    # Finishes
    "wall-area":   ("Finishes & quantities", "Wall area (plaster/paint)"),
    "layer-length": ("Finishes & quantities", "Length on one layer (pipe/chajja)"),
    # Drawing QA
    "scan":        ("Drawing QA", "One-tap scan (no tags/layers needed)"),
    "drawing-check": ("Drawing QA", "Which tools will work on this drawing"),
    "qa-report":   ("Drawing QA", "Drawing health report (find errors)"),
    "dxf-diff":    ("Drawing QA", "Compare two drawings / revisions"),
    "attr-audit":  ("Drawing QA", "Block attribute + layer audit"),
    "findtext":    ("Drawing QA", "Find text on the drawing"),
    "entity-count": ("Drawing QA", "Count entities by type"),
    # Openings extra counts
    "parking":     ("Openings", "Count parking spaces"),
    "sanitary":    ("Openings", "Count toilets / sanitary fixtures"),
    # Areas
    "area-statement": ("Areas", "Area statement (one tap)"),
    # Layer analysis (works across every layer)
    "layers":       ("Layer analysis", "Layers + entity counts"),
    "layer-report": ("Layer analysis", "Full breakdown of every layer"),
    "layer-detail": ("Layer analysis", "Entity breakdown for one layer"),
    "layer-blocks": ("Layer analysis", "Blocks per layer"),
    "layer-texts":  ("Layer analysis", "Text per layer"),
    "layer-bounds": ("Layer analysis", "Size per layer"),
    "layer-area":   ("Layer analysis", "Closed area on one layer"),
    "layer-length": ("Layer analysis", "Length on one layer"),
    "bbox-layer":   ("Layer analysis", "Footprint / size of a layer"),
    "which-layer":  ("Layer analysis", "Find which layers hold X"),
    "purge":        ("Layer analysis", "Empty / junk layers"),
    # Finishes & quantities
    "hatch":       ("Finishes & quantities", "Hatch / finish area per layer"),
    "lengths":     ("Finishes & quantities", "Line lengths per layer (walls/pipes)"),
    # Drawing QA
    "info":        ("Drawing QA", "Drawing summary"),
    "extents":     ("Drawing QA", "Drawing size / sheet box"),
    "dims":        ("Drawing QA", "Dimension values"),
    "textheights": ("Drawing QA", "Text-height histogram"),
    "text":        ("Drawing QA", "Dump all text to file"),
    "export":      ("Drawing QA", "Export entities to CSV"),
    # Calculators (no drawing needed)
    "calc-concrete": ("Calculators", "Concrete volume (L×B×H×N)"),
    "calc-rebar":    ("Calculators", "Rebar weight"),
    "calc-brick":    ("Calculators", "Brick count for a wall"),
    "calc-plaster":  ("Calculators", "Plaster mortar"),
    "calc-paint":    ("Calculators", "Paint litres"),
    "calc-tiles":    ("Calculators", "Tiles for a floor"),
    "calc-excavation": ("Calculators", "Excavation volume"),
    "calc-steel":    ("Calculators", "Rebar weight from %"),
    "calc-stair":    ("Calculators", "Stair steps"),
    "calc-watertank": ("Calculators", "Water-tank capacity"),
    "calc-cost":     ("Calculators", "Amount = qty × rate"),
    "convert":       ("Calculators", "Unit conversion"),
    "formulas":      ("Calculators", "Show every formula + basis"),
    # MEP / cleanroom / HVAC
    "calc-ach":        ("MEP & services", "HVAC air changes -> airflow"),
    "calc-water-demand": ("MEP & services", "Daily water demand"),
    "calc-hvac":       ("MEP & services", "Cooling load (TR)"),
    "cleanroom-ref":   ("MEP & services", "Cleanroom grade / ACH reference"),
    "coving-length":   ("MEP & services", "Coving / skirting length"),
    "calc-electrical": ("MEP & services", "Electrical load (kVA)"),
    "calc-lighting":   ("MEP & services", "Light fixtures for lux"),
    "calc-rainwater":  ("MEP & services", "Roof runoff + downpipe"),
    "calc-sewage":     ("MEP & services", "Sewage / STP load"),
    # Code & compliance
    "calc-fire-area":  ("Code & compliance", "Fire compartment check"),
    "calc-travel":     ("Code & compliance", "Travel distance to exit"),
    "calc-fsi":        ("Code & compliance", "FSI / FAR consumed"),
    "calc-coverage":   ("Code & compliance", "Ground coverage %"),
    "calc-occupancy":  ("Code & compliance", "Occupant load + exits"),
    "calc-parking":    ("Code & compliance", "Parking spaces (ECS)"),
    "calc-ramp":       ("Code & compliance", "Ramp run length"),
    "calc-exit-width": ("Code & compliance", "Required egress width"),
    "door-width-check": ("Code & compliance", "Flag narrow openings"),
    "room-count":      ("Industry & rooms", "Count rooms by keyword"),
    # Estimate
    "priced-boq":      ("Estimate", "Priced BOQ (takeoff x rate library)"),
    "rates":           ("Estimate", "View the rate library"),
}
# BOQ/spreadsheet tools are parked for now -- a proper quotation module comes
# later. The CLI commands still exist; they're just hidden from the product.
CATALOG_HIDE = {"project", "tools", "catalog", "dwghelp",
                "xls", "xlsdump", "find", "compare",
                # manual calculators superseded by auto drawing-takeoff tools
                # (concrete / slab / chajja / plinth-beams already read the
                # sizes from the drawing and output steel/shutter/plaster too)
                "calc-concrete", "calc-rebar", "calc-steel",
                "calc-slab", "calc-slab-steel", "calc-slab-shutter",
                "calc-slab-plaster",
                "calc-chajja", "calc-chajja-shutter", "calc-chajja-plaster",
                "calc-plinth-beam"}


# ======================================================================
# Helpers
# ======================================================================
def need_ezdxf():
    if ezdxf is None:
        sys.exit("ERROR: ezdxf not installed.  Run:  pip install ezdxf")


def load_dxf(path):
    """Open a DXF and return the document, or exit with a clear message."""
    need_ezdxf()
    if not os.path.isfile(path):
        sys.exit(f"ERROR: file not found: {path}")
    try:
        return ezdxf.readfile(path)
    except Exception as e:
        sys.exit(f"ERROR: could not read DXF '{path}': {e}")


def entity_text(e):
    """Return the plain text of a TEXT/MTEXT entity, else None."""
    t = e.dxftype()
    if t == "MTEXT":
        try:
            return e.plain_text()
        except Exception:
            return e.text
    if t == "TEXT":
        return e.dxf.text
    return None


def all_text_entities(msp):
    """Yield (text, x, y, layer) for every TEXT/MTEXT in a layout."""
    for e in msp:
        s = entity_text(e)
        if not s:
            continue
        p = e.dxf.insert
        yield s, p.x, p.y, (e.dxf.layer or "")


def fmt_table(rows, headers):
    """Render a list-of-rows as an aligned monospaced table."""
    cols = list(zip(*([headers] + rows))) if rows else [[h] for h in headers]
    widths = [max(len(str(c)) for c in col) for col in cols]
    line = lambda r: "  ".join(str(c).ljust(w) for c, w in zip(r, widths))
    out = [line(headers), "  ".join("-" * w for w in widths)]
    out += [line(r) for r in rows]
    return "\n".join(out)


def parse_levels(arg):
    """Parse --levels 'GROUND:1.0,FIRST:8.0' into a dict, else defaults."""
    if not arg:
        return dict(DEFAULT_LEVELS)
    d = dict(DEFAULT_LEVELS)
    for part in arg.split(","):
        k, v = part.split(":")
        d[k.strip().upper()] = float(v)
    return d


_LVLPAT = re.compile(
    r"(ROAD|GROUND|FIRST|SECOND|THIRD|FOURTH|FORTH|FIFTH|SIXTH|TERRACE|"
    r"PLINTH|LIFT\s*TOP|LIFT\s*PARAPET|PARAPET)\b[^\n]*?LVL\.?\s*\+?\s*(-?\d+\.?\d*)",
    re.I)


def extract_levels(msp):
    """Read floor levels straight from the drawing's own 'X FLOOR LVL +N' marks.

    Returns {GROUND: 1.0, FIRST: 8.0, ...} -- nothing hardcoded per project, so
    every uploaded drawing gives its own real levels.
    """
    out = {}
    for s, *_ in all_text_entities(msp):
        u = re.sub(r"\s+", " ", s.strip().upper())
        m = _LVLPAT.search(u)
        if not m:
            continue
        key = m.group(1).replace(" ", "").replace("FORTH", "FOURTH")
        try:
            out[key] = float(m.group(2))
        except ValueError:
            pass
    return out


# ----------------------------------------------------------------------
# Project / multi-building support
# ----------------------------------------------------------------------
# A "project" is one plot that may hold several buildings, each with its own
# drawing.  It is stored as project.json in the working directory:
#
#   {
#     "plot": "Survey 169/Part-2, Manhulli",
#     "buildings": {
#       "M1": {"dxf": "block_M1.dxf", "note": "Production Block"},
#       "M2": {"dxf": "block_M2.dxf", "note": "Utility Block"}
#     }
#   }
#
# Any DXF command then accepts  --building M1  instead of a file path, or
# --all to run across every building in the plot.
# Data dir: AC_DATA env (set by the desktop launcher) else the working dir.
# Keeps project.json + uploads in one writable place for the packaged .exe.
DATA_DIR = os.environ.get("AC_DATA") or os.getcwd()
PROJECT_FILE = os.path.join(DATA_DIR, "project.json")
RATES_FILE = os.path.join(DATA_DIR, "rates.json")


def load_rates():
    """Rate library from rates.json (data dir), else the built-in defaults."""
    if os.path.isfile(RATES_FILE):
        import json
        try:
            with open(RATES_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return dict(DEFAULT_RATES)


def _seed_rates():
    """Write the default rate library on first run so the architect can edit it."""
    if os.path.isfile(RATES_FILE):
        return
    import json
    try:
        with open(RATES_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_RATES, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def load_project(path=PROJECT_FILE):
    """Load the project manifest, or return an empty skeleton."""
    if os.path.isfile(path):
        import json
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {"plot": "", "buildings": {}}


def save_project(proj, path=PROJECT_FILE):
    import json
    with open(path, "w", encoding="utf-8") as f:
        json.dump(proj, f, indent=2, ensure_ascii=False)


def resolve_file(args):
    """Return the DXF path for a command, from --building or a literal path.

    Precedence: explicit file path > --building lookup in project.json.
    Exits with a helpful message if neither resolves.
    """
    path = getattr(args, "file", None)
    if path:
        return path
    bld = getattr(args, "building", None)
    if bld:
        proj = load_project()
        entry = proj.get("buildings", {}).get(bld)
        if not entry:
            avail = ", ".join(proj.get("buildings", {})) or "(none)"
            sys.exit(f"ERROR: no building '{bld}' in {PROJECT_FILE}. "
                     f"Known buildings: {avail}")
        return entry["dxf"]
    sys.exit("ERROR: give a .dxf path, or --building NAME "
             "(see 'project list'), or --all.")


def open_workbook(path):
    """Return a uniform sheet iterator for .xls or .xlsx."""
    ext = os.path.splitext(path)[1].lower()
    if not os.path.isfile(path):
        sys.exit(f"ERROR: file not found: {path}")
    if ext == ".xls":
        if xlrd is None:
            sys.exit("ERROR: xlrd not installed.  Run:  pip install xlrd")
        wb = xlrd.open_workbook(path)
        sheets = []
        for sh in wb.sheets():
            rows = [[sh.cell_value(r, c) for c in range(sh.ncols)]
                    for r in range(sh.nrows)]
            sheets.append((sh.name, rows))
        return sheets
    elif ext in (".xlsx", ".xlsm"):
        if openpyxl is None:
            sys.exit("ERROR: openpyxl not installed.  Run:  pip install openpyxl")
        wb = openpyxl.load_workbook(path, data_only=True)
        sheets = []
        for ws in wb.worksheets:
            rows = [[("" if v is None else v) for v in row]
                    for row in ws.iter_rows(values_only=True)]
            sheets.append((ws.title, rows))
        return sheets
    else:
        sys.exit(f"ERROR: unsupported spreadsheet type: {ext}")


# ======================================================================
# DXF commands
# ======================================================================
def cmd_info(args):
    """Drawing summary: version, layers, blocks, entity counts, extents."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    ec = Counter(e.dxftype() for e in msp)
    print(f"File      : {args.file}")
    print(f"DXF ver   : {doc.dxfversion}")
    print(f"Layers    : {len(doc.layers)}")
    print(f"Blocks    : {len(doc.blocks)}")
    print(f"Entities  : {sum(ec.values())} in modelspace")
    print()
    rows = [[n, t] for t, n in ec.most_common()]
    print(fmt_table(rows, ["count", "entity type"]))
    ext = (doc.header.get("$EXTMIN"), doc.header.get("$EXTMAX"))
    if ext[0] and ext[1]:
        dx = ext[1][0] - ext[0][0]
        dy = ext[1][1] - ext[0][1]
        print(f"\nExtents   : {dx:,.0f} x {dy:,.0f} drawing units")


def cmd_layers(args):
    """List every layer with its modelspace entity count."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    ec = Counter((e.dxf.layer or "") for e in msp)
    rows = []
    for i, l in enumerate(doc.layers, 1):
        nm = l.dxf.name
        rows.append([i, nm, ec.get(nm, 0)])
    print(fmt_table(rows, ["#", "layer", "entities"]))
    print(f"\nTotal layers: {len(doc.layers)}")


def cmd_text(args):
    """Dump every unique TEXT/MTEXT string (optionally to a file)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    seen, uniq = set(), []
    total = 0
    for s, *_ in all_text_entities(msp):
        s = s.strip()
        if not s:
            continue
        total += 1
        k = s.lower()
        if k not in seen:
            seen.add(k)
            uniq.append(s)
    out = args.out or (os.path.splitext(args.file)[0] + "_text.txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(uniq))
    print(f"text strings: {total} total, {len(uniq)} unique -> {out}")


def cmd_levels(args):
    """Extract floor / level annotations like 'FIRST FLOOR LVL. + 8.00'."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    pat = re.compile(r"(.+?LVL\.?)\s*\+?\s*(-?\d+\.\d+)", re.I)
    found = {}
    for s, *_ in all_text_entities(msp):
        for m in pat.finditer(s.replace("\n", " ")):
            name = re.sub(r"\s+", " ", m.group(1)).strip()
            found[name] = float(m.group(2))
    if not found:
        print("No level annotations found.")
        return
    rows = sorted(([k, f"+{v:.2f}"] for k, v in found.items()),
                  key=lambda r: float(r[1]))
    print(fmt_table(rows, ["level", "elev (m)"]))


def _column_schedule(msp):
    """Return {tag: 'WxH'} from MTEXT of form 'C12\\n450x600'."""
    sched = {}
    pat = re.compile(r"^(C\d{1,3})\s*\n\s*(\d{3,4})\s*[xX]\s*(\d{3,4})\s*$")
    for s, *_ in all_text_entities(msp):
        m = pat.match(s)
        if m:
            sched[m.group(1)] = f"{m.group(2)}x{m.group(3)}"
    return sched


def cmd_columns(args):
    """Count structural columns by tag (C1..Cn) and by size variant."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    sched = _column_schedule(msp)
    if not sched:
        print("No 'Cn / WxH' column markers found.")
        return
    by_size = Counter(sched.values())
    rows = [[sz, n, f"{_area(sz):.4f}"] for sz, n in
            sorted(by_size.items(), key=lambda kv: -_area(kv[0]))]
    print(fmt_table(rows, ["section (mm)", "count", "area (m2)"]))
    print(f"\nDistinct column tags : {len(sched)}")
    print(f"Total column types   : {sum(by_size.values())}")


def _area(size):
    """'750x750' -> area in m^2 (input mm)."""
    w, h = size.lower().split("x")
    return (int(w) / 1000.0) * (int(h) / 1000.0)


def cmd_floorcols(args):
    """Count columns per floor plan by assigning each marker to the
    nearest '... FLOOR PLAN' title.  Handles drawings where the whole
    sheet set is duplicated by restricting to one bounding box."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    want = re.compile(r"(GROUND|FIRST|SECOND|THIRD|FOURTH|FORTH|TERRACE)\s+FLOOR\s+PLAN")
    sizepat = re.compile(r"^(C\d{1,3})\s*\n\s*(\d{3,4})\s*[xX]\s*(\d{3,4})\s*$")

    box = None
    if args.box:
        x0, y0, x1, y1 = map(float, args.box.split(","))
        box = (x0, y0, x1, y1)

    def inbox(x, y):
        if not box:
            return True
        return box[0] < x < box[2] and box[1] < y < box[3]

    markers, titles = [], []
    for s, x, y, _ in all_text_entities(msp):
        if not inbox(x, y):
            continue
        m = sizepat.match(s)
        if m:
            markers.append((x, y, f"{m.group(2)}x{m.group(3)}"))
        t = want.search(s.strip().upper())
        if t:
            nm = t.group(1)
            nm = "FOURTH" if nm == "FORTH" else nm
            titles.append((x, y, nm))

    if not titles:
        print("No floor-plan titles found. Try --box x0,y0,x1,y1 for one sheet copy.")
        return

    floor = defaultdict(Counter)
    for x, y, sz in markers:
        best, bd = None, 1e30
        for tx, ty, nm in titles:
            d = (tx - x) ** 2 + (ty - y) ** 2
            if d < bd:
                bd, best = d, nm
        floor[best][sz] += 1

    sizes = sorted({sz for c in floor.values() for sz in c},
                   key=lambda s: -_area(s))
    rows = []
    for nm in FLOOR_ORDER:
        if nm not in floor:
            continue
        c = floor[nm]
        rows.append([nm, sum(c.values())] + [c.get(s, 0) for s in sizes])
    print(fmt_table(rows, ["floor", "total"] + sizes))
    print("\nNote: if counts look ~2-3x high, the sheet set is duplicated;")
    print("re-run with --box x0,y0,x1,y1 to isolate one copy (see 'extents').")


def cmd_colvol(args):
    """Estimate column concrete volume per size variant.

    Model: a column shown on floor-N plan rises through that storey to
    the floor above.  Volume = section area x storey height, summed.
    Interior columns stop at terrace slab; tune heights with --levels.
    """
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    # Levels come from the drawing itself; --levels overrides; baked scheme is
    # only a last-resort fallback so the tool still runs on an unlabelled DXF.
    if args.levels:
        levels = parse_levels(args.levels)
    else:
        levels = extract_levels(msp) or dict(DEFAULT_LEVELS)

    # storey heights from consecutive floor levels
    seq = [f for f in FLOOR_ORDER if f in levels]
    heights = {}
    for i in range(len(seq) - 1):
        heights[seq[i]] = levels[seq[i + 1]] - levels[seq[i]]
    # top floor -> terrace already covered; terrace handled separately

    # per-floor column counts (single sheet via --box recommended)
    args2 = argparse.Namespace(file=args.file, box=args.box)
    # reuse the floorcols clustering inline
    want = re.compile(r"(GROUND|FIRST|SECOND|THIRD|FOURTH|FORTH|TERRACE)\s+FLOOR\s+PLAN")
    sizepat = re.compile(r"^(C\d{1,3})\s*\n\s*(\d{3,4})\s*[xX]\s*(\d{3,4})\s*$")
    box = None
    if args.box:
        x0, y0, x1, y1 = map(float, args.box.split(","))
        box = (x0, y0, x1, y1)

    def inbox(x, y):
        return True if not box else (box[0] < x < box[2] and box[1] < y < box[3])

    markers, titles = [], []
    for s, x, y, _ in all_text_entities(msp):
        if not inbox(x, y):
            continue
        m = sizepat.match(s)
        if m:
            markers.append((x, y, f"{m.group(2)}x{m.group(3)}"))
        t = want.search(s.strip().upper())
        if t:
            nm = t.group(1)
            nm = "FOURTH" if nm == "FORTH" else nm
            titles.append((x, y, nm))

    floor = defaultdict(Counter)
    for x, y, sz in markers:
        best, bd = None, 1e30
        for tx, ty, nm in titles:
            d = (tx - x) ** 2 + (ty - y) ** 2
            if d < bd:
                bd, best = d, nm
        floor[best][sz] += 1

    by_variant = Counter()
    seg_rows = []
    grand = 0.0
    for nm in FLOOR_ORDER[:-1]:  # GROUND..FOURTH carry to the floor above
        h = heights.get(nm)
        if h is None or nm not in floor:
            continue
        c = floor[nm]
        vol = sum(n * _area(sz) * h for sz, n in c.items())
        for sz, n in c.items():
            by_variant[sz] += n * _area(sz) * h
        grand += vol
        seg_rows.append([nm, f"{h:.2f}", sum(c.values()), f"{vol:.2f}"])

    print("STOREY CONCRETE (columns rising to the floor above)")
    print(fmt_table(seg_rows, ["floor", "h (m)", "cols", "vol m3"]))
    print()
    vr = [[sz, f"{v:.2f}"] for sz, v in
          sorted(by_variant.items(), key=lambda kv: -kv[1])]
    print("BY SIZE VARIANT")
    print(fmt_table(vr, ["section (mm)", "concrete m3"]))
    print(f"\nSUPERSTRUCTURE column concrete (excl. terrace top & plinth): "
          f"{grand:.2f} m3")
    print("Add separately: plinth columns (below ground floor) and any")
    print("peripheral parapet / lift-cabin columns above terrace slab.")


def cmd_areas(args):
    """Sum closed LWPOLYLINE areas per layer (floor / room / plot areas)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    by_layer = defaultdict(lambda: [0, 0.0])  # count, area(m2)
    for e in msp:
        if e.dxftype() == "LWPOLYLINE" and e.closed:
            try:
                a = abs(e.get_area()) if hasattr(e, "get_area") else _poly_area(e)
            except Exception:
                a = _poly_area(e)
            lay = e.dxf.layer or ""
            by_layer[lay][0] += 1
            by_layer[lay][1] += a / 1_000_000.0  # mm^2 -> m^2
    rows = [[lay, cnt, f"{ar:.2f}"] for lay, (cnt, ar) in
            sorted(by_layer.items(), key=lambda kv: -kv[1][1])]
    print(fmt_table(rows, ["layer", "closed polys", "area m2"]))
    print("\n(Assumes drawing units = mm. Divide differently if units differ.)")


def _auto_layer(msp, kind="area"):
    """Best-guess layer when the user gives none, so a tool never dead-ends
    asking for '--layer'. kind: 'area' = most closed polylines, 'line' = most
    line/polyline entities, 'busy' = most entities of any kind."""
    c = Counter()
    for e in msp:
        lay = e.dxf.layer or ""
        if lay in ("Defpoints",):
            continue
        t = e.dxftype()
        if kind == "area":
            if t == "LWPOLYLINE" and getattr(e, "closed", False):
                c[lay] += 1
        elif kind == "line":
            if t in ("LINE", "LWPOLYLINE", "POLYLINE", "ARC"):
                c[lay] += 1
        else:
            c[lay] += 1
    return c.most_common(1)[0][0] if c else None


def _resolve_layer(msp, args, kind="area"):
    """Return the layer to use: the one the user chose, else an auto-pick.
    Prints a one-line note when it auto-picks so the result stays honest."""
    if getattr(args, "layer", None):
        return args.layer
    lay = _auto_layer(msp, kind)
    if lay:
        print(f"(auto-picked layer '{lay}' -- change it in Advanced if this "
              f"isn't the right one)\n")
    return lay


def _poly_area(e):
    """Shoelace area for an LWPOLYLINE (ignores bulges)."""
    pts = [(p[0], p[1]) for p in e.get_points()]
    a = 0.0
    for i in range(len(pts)):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % len(pts)]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def cmd_lengths(args):
    """Total LINE + polyline length per layer (wall runs, pipe runs)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    by_layer = defaultdict(float)
    for e in msp:
        t = e.dxftype()
        lay = e.dxf.layer or ""
        if t == "LINE":
            a, b = e.dxf.start, e.dxf.end
            by_layer[lay] += math.dist((a.x, a.y), (b.x, b.y))
        elif t == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            for i in range(len(pts) - 1):
                by_layer[lay] += math.dist(pts[i], pts[i + 1])
            if e.closed and len(pts) > 2:
                by_layer[lay] += math.dist(pts[-1], pts[0])
    rows = [[lay, f"{ln/1000.0:,.2f}"] for lay, ln in
            sorted(by_layer.items(), key=lambda kv: -kv[1])]
    print(fmt_table(rows, ["layer", "length (m)"]))
    print("\n(Assumes drawing units = mm.)")


def cmd_blocks(args):
    """Count block references (INSERT) by block name -- fixtures, symbols."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    c = Counter(e.dxf.name for e in msp if e.dxftype() == "INSERT")
    if not c:
        print("No block references found.")
        return
    rows = [[n, cnt] for n, cnt in c.most_common()]
    print(fmt_table(rows, ["block name", "count"]))
    print(f"\nDistinct blocks placed: {len(c)}   Total inserts: {sum(c.values())}")


def cmd_schedule(args):
    """Extract a door/window schedule: tags + WxH sizes from text."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    # tags like D1, D6, V-1, VP4, RS, ED1 followed by a WxH somewhere
    sizepat = re.compile(r"(\d{3,4})\s*[xX]\s*(\d{3,4})")
    tagpat = re.compile(r"^(D\d+|V-?\d+|VP\d+|FVP\d+|ED\d+|HD|HW|RS-?\d*|SW\d*|OW|LW)$")
    sizes = Counter()
    tags = set()
    for s, *_ in all_text_entities(msp):
        for line in s.split("\n"):
            line = line.strip()
            if tagpat.match(line):
                tags.add(line)
            m = sizepat.search(line)
            if m:
                sizes[f"{m.group(1)} x {m.group(2)}"] += 1
    print("DOOR / WINDOW / OPENING TAGS FOUND:")
    print("  " + ", ".join(sorted(tags)) if tags else "  (none)")
    print("\nSIZES MENTIONED (count of occurrences):")
    rows = [[sz, n] for sz, n in sizes.most_common(40)]
    print(fmt_table(rows, ["size (mm)", "occurrences"]))


def cmd_hatch(args):
    """Total HATCH area per layer -- flooring / finish quantities."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    by_layer = defaultdict(lambda: [0, 0.0])
    for e in msp:
        if e.dxftype() == "HATCH":
            lay = e.dxf.layer or ""
            area = 0.0
            try:
                for path in e.paths:
                    pts = []
                    for v in getattr(path, "vertices", []):
                        pts.append((v[0], v[1]))
                    if len(pts) >= 3:
                        a = 0.0
                        for i in range(len(pts)):
                            x1, y1 = pts[i]
                            x2, y2 = pts[(i + 1) % len(pts)]
                            a += x1 * y2 - x2 * y1
                        area += abs(a) / 2.0
            except Exception:
                pass
            by_layer[lay][0] += 1
            by_layer[lay][1] += area / 1_000_000.0
    rows = [[lay, cnt, f"{ar:.2f}"] for lay, (cnt, ar) in
            sorted(by_layer.items(), key=lambda kv: -kv[1][1])]
    print(fmt_table(rows, ["layer", "hatches", "area m2"]))
    print("\n(Polyline-boundary hatches only; spline/island edges approximate.)")


def cmd_dims(args):
    """List dimension measurements (the actual numbers on the drawing)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    vals = []
    for e in msp:
        if e.dxftype() == "DIMENSION":
            m = e.get_measurement()
            if isinstance(m, (int, float)):
                vals.append(m)
    if not vals:
        print("No dimensions found.")
        return
    print(f"Dimensions      : {len(vals)}")
    print(f"Min / Max       : {min(vals):,.1f} / {max(vals):,.1f}")
    print(f"Sum             : {sum(vals):,.1f}")
    if args.list:
        for v in sorted(vals):
            print(f"  {v:,.1f}")


def cmd_textheights(args):
    """Histogram of distinct TEXT/MTEXT heights -- spot scaled/imported text.

    Added from a live 'ask anything' query: counts every annotation height so
    the architect can see which sizes are deliberate (100/150/230/300 mm) vs
    odd values dragged in from scaled blocks.
    """
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    c = Counter()
    for e in msp:
        t = e.dxftype()
        if t == "TEXT":
            h = e.dxf.height
        elif t == "MTEXT":
            h = e.dxf.char_height
        else:
            continue
        if h:
            c[round(h, 3)] += 1
    if not c:
        print("No text found.")
        return
    print(f"Distinct text heights: {len(c)}")
    top = c.most_common(1)[0]
    print(f"Most common: {top[0]:g} mm ({top[1]} uses)\n")
    rows = [[f"{h:g}", n] for h, n in c.most_common(args.top)]
    print(fmt_table(rows, ["height (mm)", "count"]))


def cmd_circles(args):
    """Count CIRCLE entities and report average / min / max radius (mm).

    Promoted into the core toolkit from an auto-saved plugin.
    """
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    radii = [c.dxf.radius for c in msp.query("CIRCLE")]
    if not radii:
        print("No CIRCLE entities found.")
        return
    print(f"CIRCLE count   : {len(radii)}")
    print(f"Average radius : {sum(radii)/len(radii):.2f} mm")
    print(f"Min / Max      : {min(radii):.2f} / {max(radii):.2f} mm")


def cmd_centroid(args):
    """Average XY centre of all entities on a layer (metres).

    Promoted into the core toolkit from a plugin -- finds the geometric
    centre of, e.g., a column grid, a core, or any named layer.
    """
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "busy")
    xs, ys = [], []
    for e in msp:
        if (e.dxf.layer or "") != args.layer:
            continue
        try:
            p = e.dxf.insert
        except Exception:
            try:
                p = e.dxf.start
            except Exception:
                continue
        xs.append(p.x)
        ys.append(p.y)
    if not xs:
        print(f"No entities on layer '{args.layer}'.")
        return
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    print(f"Layer '{args.layer}': {len(xs)} located entities")
    print(f"Centroid: ({cx/1000:.2f}, {cy/1000:.2f}) m  (raw {cx:.0f}, {cy:.0f})")


# ----------------------------------------------------------------------
# Structure helpers (shared by the column-grid tools)
# ----------------------------------------------------------------------
_SIZEPAT = re.compile(r"^(C\d{1,3})\s*\n\s*(\d{3,4})\s*[xX]\s*(\d{3,4})\s*$")
_FLOORPAT = re.compile(r"(GROUND|FIRST|SECOND|THIRD|FOURTH|FORTH|TERRACE)\s+FLOOR\s+PLAN")


def _parse_box(s):
    return tuple(float(v) for v in s.split(",")) if s else None


def _markers_titles(msp, box):
    """Column markers (x,y,tag,size) and floor titles (x,y,floor) inside box."""
    def inbox(x, y):
        return True if not box else (box[0] < x < box[2] and box[1] < y < box[3])
    markers, titles = [], []
    for s, x, y, _ in all_text_entities(msp):
        if not inbox(x, y):
            continue
        m = _SIZEPAT.match(s)
        if m:
            markers.append((x, y, m.group(1), f"{m.group(2)}x{m.group(3)}"))
        t = _FLOORPAT.search(s.strip().upper())
        if t:
            nm = t.group(1)
            titles.append((x, y, "FOURTH" if nm == "FORTH" else nm))
    return markers, titles


def _floor_of(x, y, titles):
    best, bd = None, 1e30
    for tx, ty, nm in titles:
        d = (tx - x) ** 2 + (ty - y) ** 2
        if d < bd:
            bd, best = d, nm
    return best


def _convex_hull(pts):
    pts = sorted(set(pts))
    if len(pts) <= 2:
        return pts
    def cr(o, a, b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    lo = []
    for p in pts:
        while len(lo) >= 2 and cr(lo[-2], lo[-1], p) <= 0:
            lo.pop()
        lo.append(p)
    up = []
    for p in reversed(pts):
        while len(up) >= 2 and cr(up[-2], up[-1], p) <= 0:
            up.pop()
        up.append(p)
    return lo[:-1] + up[:-1]


def _on_edge(p, a, b, tol=350):
    seglen = math.hypot(b[0]-a[0], b[1]-a[1]) or 1
    dist = abs((b[0]-a[0])*(p[1]-a[1]) - (b[1]-a[1])*(p[0]-a[0])) / seglen
    if dist > tol:
        return False
    dot = (p[0]-a[0])*(b[0]-a[0]) + (p[1]-a[1])*(b[1]-a[1])
    return -tol <= dot <= seglen * seglen + tol


def cmd_perimeter_columns(args):
    """Split a floor's columns into PERIMETER (on the building edge) vs
    INTERIOR, by convex hull of the column grid. Terrace/parapet columns are
    typically the perimeter ones. Default floor: ground.
    """
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    box = _parse_box(getattr(args, "box", None))
    markers, titles = _markers_titles(msp, box)
    floor = (args.floor or "GROUND").upper()
    pos = {}
    for x, y, tag, sz in markers:
        if titles and _floor_of(x, y, titles) != floor:
            continue
        pos[(round(x, 1), round(y, 1))] = sz
    if not pos:
        print(f"No columns found for floor {floor}.")
        return
    pts = list(pos)
    hull = _convex_hull(pts)
    hv = set(hull)
    perim = set()
    for p in pts:
        if p in hv:
            perim.add(p)
            continue
        for i in range(len(hull)):
            if _on_edge(p, hull[i], hull[(i + 1) % len(hull)]):
                perim.add(p)
                break
    per_c = Counter(pos[p] for p in perim)
    int_c = Counter(pos[p] for p in pts if p not in perim)
    print(f"Floor: {floor}   total columns: {len(pts)}")
    print(f"PERIMETER (edge): {sum(per_c.values())}   INTERIOR: {sum(int_c.values())}\n")
    sizes = sorted(set(per_c) | set(int_c), key=lambda s: -_area(s))
    rows = [[s, per_c.get(s, 0), int_c.get(s, 0)] for s in sizes]
    print(fmt_table(rows, ["section (mm)", "perimeter", "interior"]))


def cmd_plinth_columns(args):
    """Columns at plinth = the ground-floor column set (they rise from the
    footing up to ground-floor level). Count by size."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    box = _parse_box(getattr(args, "box", None))
    markers, titles = _markers_titles(msp, box)
    c = Counter()
    for x, y, tag, sz in markers:
        if titles and _floor_of(x, y, titles) != "GROUND":
            continue
        c[sz] += 1
    if not c:
        print("No ground-floor columns found.")
        return
    print(f"Plinth columns (= ground-floor set): {sum(c.values())}\n")
    rows = [[s, n] for s, n in sorted(c.items(), key=lambda kv: -_area(kv[0]))]
    print(fmt_table(rows, ["section (mm)", "count"]))


# ======================================================================
# PLINTH SUITE  --  every plinth-level quantity in one place
# ----------------------------------------------------------------------
# The plinth is the band between the footing top and the ground-floor slab.
# colvol deliberately SKIPS this band ("add plinth separately"); these tools
# fill that gap plus every other plinth-level BOQ item (filling, DPC, apron,
# anti-termite, masonry, PCC bed). DXF tools read the drawing; the calc-*
# ones are pure formulas with adjustable, documented assumptions.
# ======================================================================
def _plinth_height(levels, override):
    """Plinth height (m): explicit override, else GROUND-ROAD from the
    drawing's level marks, else None (caller must ask for --plinth)."""
    if override:
        return float(override)
    g, r = levels.get("GROUND"), levels.get("ROAD")
    if g is not None and r is not None:
        return g - r
    return None


def _largest_closed(doc, msp, layer=None):
    """(area m2, perimeter m) of the building footprint.

    When no specific layer is asked for, prefer the GEOMETRY footprint (the
    hull the columns enclose) -- it's immune to a plot/sheet border or a stray
    entity far away, which is what used to inflate this to absurd numbers. Falls
    back to the largest plausible closed polyline (old behaviour)."""
    if not layer:
        g = _detect_geometry(msp)
        if g.get("floor_count") and g["footprint"] > 0 and g.get("perimeter"):
            return (g["footprint"], g["perimeter"])
    cands = []
    for e in msp:
        if e.dxftype() != "LWPOLYLINE" or not e.closed:
            continue
        if layer and (e.dxf.layer or "") != layer:
            continue
        try:
            a = abs(e.get_area()) / 1_000_000.0
        except Exception:
            a = _poly_area(e) / 1_000_000.0
        pts = [(p[0], p[1]) for p in e.get_points()]
        per = sum(math.dist(pts[i], pts[(i + 1) % len(pts)])
                  for i in range(len(pts))) / 1000.0
        cands.append((a, per))
    if not cands:
        return None
    cands.sort(reverse=True)
    # skip the sheet-border-sized shape (same guard as 'builtup')
    sheet = None
    ext = (doc.header.get("$EXTMIN"), doc.header.get("$EXTMAX"))
    if ext[0] and ext[1]:
        sheet = abs((ext[1][0]-ext[0][0]) * (ext[1][1]-ext[0][1])) / 1_000_000.0
    if sheet and cands[0][0] > 0.6 * sheet:
        smaller = [c for c in cands if c[0] < 0.6 * sheet]
        if smaller:
            return smaller[0]
    return cands[0]


def cmd_plinth_colvol(args):
    """Plinth column concrete = ground-floor columns rising through the plinth
    band (footing top -> ground-floor slab). Closes the gap that 'colvol'
    leaves open. Volume = section area x plinth height, summed by size."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    levels = parse_levels(args.levels) if args.levels else \
        (extract_levels(msp) or dict(DEFAULT_LEVELS))
    h = _plinth_height(levels, getattr(args, "plinth", None))
    if h is None:
        print("Plinth height unknown (no ROAD/GROUND level marks). "
              "Re-run with --plinth <metres> (e.g. --plinth 1.2).")
        return
    box = _parse_box(getattr(args, "box", None))
    markers, titles = _markers_titles(msp, box)
    c = Counter()
    for x, y, tag, sz in markers:
        if titles and _floor_of(x, y, titles) != "GROUND":
            continue
        c[sz] += 1
    if not c:
        print("No ground-floor columns found (try --box to isolate one sheet).")
        return
    rows, grand = [], 0.0
    for sz, n in sorted(c.items(), key=lambda kv: -_area(kv[0])):
        v = n * _area(sz) * h
        grand += v
        rows.append([sz, n, f"{v:.3f}"])
    print(f"PLINTH COLUMN CONCRETE   (plinth height {h:.2f} m)")
    print(fmt_table(rows, ["section (mm)", "count", "concrete m3"]))
    print(f"\nTotal plinth column concrete: {grand:.2f} m3")
    print("Add to colvol's superstructure figure for the full column total.")


def cmd_plinth_beams(args):
    """Plinth (tie) beam concrete + steel from the beam run on a layer.
    Pass --layer (plinth-beam layer); --width/--depth in mm; --steel %."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "line")
    if not args.layer:
        print("No line geometry found to measure.")
        return
    L = 0.0
    for e in msp:
        if (e.dxf.layer or "") != args.layer:
            continue
        if e.dxftype() == "LINE":
            L += math.dist((e.dxf.start.x, e.dxf.start.y),
                           (e.dxf.end.x, e.dxf.end.y))
        elif e.dxftype() == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            for i in range(len(pts) - 1):
                L += math.dist(pts[i], pts[i + 1])
    lm = L / 1000.0
    w, d = args.width / 1000.0, args.depth / 1000.0
    vol = lm * w * d
    steel = vol * 7850 * args.steel / 100.0
    print(f"Layer '{args.layer}': plinth-beam run {lm:,.1f} m")
    print(f"Section {args.width} x {args.depth} mm  ->  {w*d:.3f} m2")
    print(f"Concrete = {lm:,.1f} x {w:.3f} x {d:.3f} = {vol:.2f} m3")
    print(f"Steel ({args.steel}%) ~ {steel:,.0f} kg ({steel/1000:.3f} tonne)")


def cmd_plinth_area(args):
    """Plinth area + perimeter (built-up footprint at plinth level). Drives
    DPC, anti-termite, filling & protection takeoffs. --layer for a clean
    boundary; feed the numbers into the calc-* plinth tools."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    res = _largest_closed(doc, msp, args.layer)
    if not res:
        print("No closed boundary polyline found"
              + (f" on layer '{args.layer}'" if args.layer else "")
              + ". Try --layer <plinth/wall boundary>.")
        return
    area, per = res
    print(f"Plinth area (footprint) : {area:,.2f} m2")
    print(f"Plinth perimeter        : {per:,.1f} m")
    print("\nUse these with:")
    print(f"  calc-plinth-fill  --area {area:.1f} --depth <m>")
    print(f"  calc-dpc          --perimeter {per:.1f}")
    print(f"  calc-anti-termite --area {area:.1f} --perimeter {per:.1f}")
    print(f"  calc-plinth-protection --perimeter {per:.1f}")
    if not args.layer:
        print("Tip: pass --layer for a cleaner boundary.")


# ----- Plinth calculators (no drawing needed) -------------------------
def cmd_calc_plinth_beam(args):
    """Plinth-beam concrete + steel from L x B x D x N (metres)."""
    vol = args.l * args.b * args.d * args.n
    steel = vol * 7850 * args.steel / 100.0
    print(f"Plinth beam {args.l} x {args.b} x {args.d} x {args.n} "
          f"= {vol:.3f} m3 concrete")
    print(f"Steel ({args.steel}%) ~ {steel:,.0f} kg ({steel/1000:.3f} tonne)")


def cmd_calc_plinth_fill(args):
    """Earth / murrum filling inside plinth: area x depth, plus loose volume
    to order (compaction adds ~20%)."""
    v = args.area * args.depth
    loose = v * (1 + args.compact / 100.0)
    print(f"Plinth fill {args.area} m2 x {args.depth} m = {v:.2f} m3 (compacted)")
    print(f"Loose volume to order (+{args.compact}%): {loose:.2f} m3")


def cmd_calc_plinth_masonry(args):
    """Plinth wall masonry: perimeter x thickness x height -> volume + bricks."""
    vol = args.perimeter * (args.thk / 1000.0) * args.height
    bricks = vol * 500
    print(f"Plinth wall {args.perimeter} m x {args.thk} mm x {args.height} m "
          f"= {vol:.3f} m3")
    print(f"Bricks (~500/m3) ~ {bricks:,.0f} nos")
    print(f"Mortar (~0.30 m3/m3) ~ {vol*0.3:.3f} m3")


def cmd_calc_dpc(args):
    """Damp-proof course: perimeter x band width -> area; x thickness ->
    concrete (1:1.5:3, ~7.8 bag/m3 cement)."""
    area = args.perimeter * (args.width / 1000.0)
    vol = area * (args.thk / 1000.0)
    print(f"DPC band {args.perimeter} m x {args.width} mm = {area:.2f} m2")
    print(f"Concrete ({args.thk} mm thick) = {vol:.3f} m3")
    print(f"Cement (1:1.5:3, ~7.8 bag/m3) ~ {vol*7.8:.1f} bags")


def cmd_calc_anti_termite(args):
    """Anti-termite treatment area = plinth (top surface) area + the vertical
    foundation face along the perimeter (perimeter x trench depth). Chemical
    emulsion at ~5 L/m2 (pre-construction soil treatment, IS 6313)."""
    vert = args.perimeter * args.depth
    total = args.area + vert
    chem = total * args.rate
    print(f"Top surface {args.area} m2 + vertical face "
          f"({args.perimeter} m x {args.depth} m = {vert:.1f} m2)")
    print(f"Treatment area = {total:.1f} m2")
    print(f"Chemical (~{args.rate} L/m2) ~ {chem:,.1f} litres emulsion")


def cmd_calc_plinth_protection(args):
    """Plinth protection apron around the building: perimeter x width ->
    area; x thickness -> PCC volume (1:4:8, ~3.4 bag/m3)."""
    area = args.perimeter * args.width
    vol = area * (args.thk / 1000.0)
    print(f"Apron {args.perimeter} m x {args.width} m wide = {area:.1f} m2")
    print(f"PCC ({args.thk} mm) = {vol:.2f} m3")
    print(f"Cement (1:4:8, ~3.4 bag/m3) ~ {vol*3.4:.1f} bags")


def cmd_calc_pcc_bed(args):
    """PCC leveling bed below plinth/footing: area x thickness -> volume
    (1:4:8 nominal, ~3.4 cement bags/m3)."""
    vol = args.area * (args.thk / 1000.0)
    print(f"PCC bed {args.area} m2 x {args.thk} mm = {vol:.2f} m3")
    print(f"Cement (1:4:8, ~3.4 bag/m3) ~ {vol*3.4:.1f} bags")


# ======================================================================
# CHAJJA / SUNSHADE SUITE  --  RCC projection over openings
# ----------------------------------------------------------------------
# A chajja (sunshade) is the small cantilever slab over a window/door.
# Section is normally TAPERED: thicker at the wall (root, ~150 mm) thinning
# to the tip (~75 mm); volume uses the average thickness. These tools cover
# concrete, steel, shuttering (formwork) and plastering -- the DXF tool reads
# the chajja run off a layer; the calc-* ones are pure formulas.
# ======================================================================
def _avg_thk(root, tip):
    """Average chajja thickness (m) from root & tip in mm."""
    return ((root + tip) / 2.0) / 1000.0


def cmd_chajja(args):
    """Chajja/sunshade concrete + steel from its run on a layer. Pass --layer
    (the chajja/sunshade layer); --proj (projection m); tapered --root/--tip
    (mm); --steel %. Volume = run x projection x average thickness."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "line")
    if not args.layer:
        print("No line geometry found to measure.")
        return
    L = 0.0
    for e in msp:
        if (e.dxf.layer or "") != args.layer:
            continue
        if e.dxftype() == "LINE":
            L += math.dist((e.dxf.start.x, e.dxf.start.y),
                           (e.dxf.end.x, e.dxf.end.y))
        elif e.dxftype() == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            for i in range(len(pts) - 1):
                L += math.dist(pts[i], pts[i + 1])
    lm = L / 1000.0
    at = _avg_thk(args.root, args.tip)
    vol = lm * args.proj * at
    steel = vol * 7850 * args.steel / 100.0
    front = lm * (args.tip / 1000.0)
    shutter = lm * args.proj + front
    plaster = 2 * args.proj * lm + front
    mortar = plaster * 0.012
    print(f"Layer '{args.layer}': chajja run {lm:,.1f} m")
    print(f"Projection {args.proj} m | section {args.root}->{args.tip} mm "
          f"(avg {at*1000:.0f} mm)")
    print(f"Concrete   = {lm:,.1f} x {args.proj} x {at:.3f} = {vol:.2f} m3")
    print(f"Steel ({args.steel}%) ~ {steel:,.0f} kg ({steel/1000:.3f} tonne)")
    print(f"Shuttering = soffit + front edge = {shutter:,.1f} m2")
    print(f"Plaster (top+bottom+drip) = {plaster:,.1f} m2; "
          f"mortar @ 12mm = {mortar:.2f} m3")


def cmd_calc_chajja(args):
    """Chajja/sunshade concrete + steel. Volume = projection x length x
    average thickness x count; tapered root->tip (mm)."""
    at = _avg_thk(args.root, args.tip)
    vol = args.proj * args.length * at * args.n
    steel = vol * 7850 * args.steel / 100.0
    print(f"Chajja {args.proj} m proj x {args.length} m x avg {at*1000:.0f} mm "
          f"x {args.n} = {vol:.3f} m3 concrete")
    print(f"Section: root {args.root} mm -> tip {args.tip} mm")
    print(f"Steel ({args.steel}%) ~ {steel:,.0f} kg ({steel/1000:.3f} tonne)")


def cmd_calc_chajja_shutter(args):
    """Chajja shuttering / formwork (contact) area: soffit + front edge + two
    ends (back sits against the wall, top is finished). x count."""
    soffit = args.proj * args.length
    front = args.length * (args.tip / 1000.0)
    ends = 2 * args.proj * _avg_thk(args.root, args.tip)
    area = (soffit + front + ends) * args.n
    print(f"Per chajja: soffit {soffit:.2f} + front {front:.3f} + ends "
          f"{ends:.3f} m2")
    print(f"Formwork area ({args.n} nos) = {area:.2f} m2")


def cmd_calc_chajja_plaster(args):
    """Chajja plaster area (top + bottom faces + front drip) x count, with
    mortar volume at --thk mm (1:4 mix)."""
    faces = 2 * args.proj * args.length
    front = args.length * (args.tip / 1000.0)
    area = (faces + front) * args.n
    mortar = area * (args.thk / 1000.0)
    print(f"Plaster area ({args.n} nos) = {area:.2f} m2 "
          f"(top+bottom+drip)")
    print(f"Mortar @ {args.thk} mm = {mortar:.3f} m3")
    print(f"Cement (1:4, ~5.5 bag/m3) ~ {mortar*5.5:.1f} bags; "
          f"sand ~ {mortar*1.1:.2f} m3")


# ======================================================================
# SLAB SUITE  --  floor-slab takeoff, mostly AUTO from the drawing
# ----------------------------------------------------------------------
# `slab` reads everything it can from the DXF -- the slab plan area (footprint
# or a chosen layer) and the number of floors (from the drawing's own level
# marks) -- so one click gives concrete, steel, shuttering and ceiling plaster
# with NOTHING typed. Defaults (thickness, steel rate) are sensible and
# overridable. The calc-* slabs are the manual fallback when there's no DXF.
# ======================================================================
def cmd_slab(args):
    """One-click slab takeoff straight from the drawing: reads the slab area
    (footprint, or --layer) and the floor count (from the level marks), then
    auto-computes concrete + steel + shuttering + ceiling plaster. Override
    only if you want: --thk (mm), --steelrate (kg/m2), --floors, --layer."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    res = _largest_closed(doc, msp, args.layer)
    if not res:
        print("No slab boundary found"
              + (f" on layer '{args.layer}'" if args.layer else "")
              + ". Pass --layer <slab/boundary layer>.")
        return
    area, per = res
    if args.floors:
        floors, seq = int(args.floors), []
    else:
        lv = extract_levels(msp)
        seq = [f for f in FLOOR_ORDER if f in lv]
        floors = len(seq) or 1
    thk = args.thk / 1000.0
    conc = area * thk * floors
    steel = area * floors * args.steelrate
    shutter = (area + per * thk) * floors
    plaster = area * floors
    mortar = plaster * (args.ptk / 1000.0)
    print("SLAB TAKEOFF  (auto from drawing)")
    print(f"Slab area / floor : {area:,.2f} m2"
          + (f"  (layer '{args.layer}')" if args.layer else "  (footprint)"))
    print(f"Floors counted    : {floors}"
          + (f"  ({', '.join(seq)})" if seq else "  (--floors)"))
    print(f"Thickness         : {args.thk:.0f} mm")
    print()
    print(f"Concrete   = {area:,.1f} x {thk:.3f} x {floors} = {conc:,.2f} m3")
    print(f"Steel      ({args.steelrate:.0f} kg/m2) = {steel:,.0f} kg "
          f"({steel/1000:.2f} tonne)")
    print(f"Shuttering = (area + perim {per:,.0f} m x thk) x {floors} "
          f"= {shutter:,.1f} m2")
    print(f"Ceiling plaster = {plaster:,.1f} m2;  mortar @ {args.ptk:.0f} mm "
          f"= {mortar:.2f} m3")
    if not args.layer:
        print("Tip: pass --layer <slab/boundary layer> for an exact area.")


def cmd_calc_slab(args):
    """Slab concrete + steel (manual): area x thickness x floors, steel by
    kg/m2. Use the 'slab' tool to pull these from a drawing automatically."""
    thk = args.thk / 1000.0
    conc = args.area * thk * args.n
    steel = args.area * args.n * args.steelrate
    print(f"Slab {args.area} m2 x {args.thk:.0f} mm x {args.n} floors "
          f"= {conc:,.2f} m3 concrete")
    print(f"Steel ({args.steelrate:.0f} kg/m2) = {steel:,.0f} kg "
          f"({steel/1000:.2f} tonne)")


def cmd_calc_slab_steel(args):
    """Slab reinforcement two ways: area-based (kg/m2) always; plus the
    %-of-concrete method if --vol is given."""
    area_kg = args.area * args.floors * args.rate
    print(f"Area method : {args.area} m2 x {args.floors} x {args.rate:.0f} "
          f"kg/m2 = {area_kg:,.0f} kg ({area_kg/1000:.2f} t)")
    if args.vol:
        vol_kg = args.vol * 7850 * args.pct / 100.0
        print(f"% method    : {args.vol} m3 x {args.pct}% x 7850 "
              f"= {vol_kg:,.0f} kg ({vol_kg/1000:.2f} t)")


def cmd_calc_slab_shutter(args):
    """Slab shuttering / formwork: soffit (plan area) + edge band
    (perimeter x thickness), x floors."""
    edge = args.perimeter * (args.thk / 1000.0)
    shutter = (args.area + edge) * args.n
    print(f"Soffit {args.area} m2 + edge ({args.perimeter} m x {args.thk:.0f} "
          f"mm = {edge:.2f} m2)")
    print(f"Formwork ({args.n} floors) = {shutter:,.1f} m2")


def cmd_calc_slab_plaster(args):
    """Slab ceiling plaster area + mortar (x floors)."""
    plaster = args.area * args.n
    mortar = plaster * (args.thk / 1000.0)
    print(f"Ceiling plaster = {args.area} m2 x {args.n} = {plaster:,.1f} m2")
    print(f"Mortar @ {args.thk:.0f} mm = {mortar:.3f} m3")
    print(f"Cement (1:4, ~5.5 bag/m3) ~ {mortar*5.5:.1f} bags; "
          f"sand ~ {mortar*1.1:.2f} m3")


# ======================================================================
# FOUNDATION SUITE  --  footing takeoff, auto from the drawing
# ----------------------------------------------------------------------
# Footings are tagged like columns: 'F1' + '1800x1800' text markers (and
# often a footing schedule table). These tools read those markers and give
# the whole foundation package -- concrete, excavation, PCC, backfill,
# steel -- one tap. Depth/thickness are site decisions a plan doesn't
# carry, so they are documented defaults, overridable in Advanced.
# ======================================================================
_FOOTPAT = re.compile(r"^(F\d{1,3})\s*\n\s*(\d{3,4})\s*[xX]\s*(\d{3,4})\s*$")


def _footing_markers(msp, box=None):
    """(x, y, tag, 'WxL') for every footing marker inside the box."""
    def inbox(x, y):
        return True if not box else (box[0] < x < box[2] and box[1] < y < box[3])
    out = []
    for s, x, y, _ in all_text_entities(msp):
        if not inbox(x, y):
            continue
        m = _FOOTPAT.match(s)
        if m:
            out.append((x, y, m.group(1), f"{m.group(2)}x{m.group(3)}"))
    return out


def cmd_footing_schedule(args):
    """Footing schedule read from the drawing's F1/F2... size markers:
    every tag with its size, plus totals per size."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    box = _parse_box(getattr(args, "box", None))
    marks = _footing_markers(msp, box)
    if not marks:
        print("No 'Fn / WxL' footing markers found. If footings are drawn "
              "without F-tags, use 'footings --layer <name>' to count them.")
        return
    sched = {}
    for x, y, tag, sz in marks:
        sched[tag] = sz
    rows = [[t, sched[t]] for t in sorted(sched, key=lambda x: int(x[1:]))]
    print(fmt_table(rows, ["tag", "size (mm)"]))
    by = Counter(m[3] for m in marks)
    print()
    print(fmt_table([[s, n] for s, n in by.most_common()],
                    ["size (mm)", "markers"]))
    print(f"\nDistinct footing tags: {len(sched)}   markers on plan: {len(marks)}")
    print("(Markers can exceed tags if the sheet set is duplicated -- "
          "set a box on the building to isolate one copy.)")


def cmd_foundation(args):
    """One-tap foundation takeoff from the drawing's footing markers:
    concrete, excavation, PCC bed, backfill and steel per footing size.
    Defaults (overridable): footing thk 450mm, depth 1.5m, PCC 100mm with
    100mm offset, working space 300mm each side, steel 0.8%."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    box = _parse_box(getattr(args, "box", None))
    marks = _footing_markers(msp, box)
    if not marks:
        print("No 'Fn / WxL' footing markers found in this drawing.")
        print("Fallbacks: 'footings --layer <name>' to count shapes, then")
        print("check the structural drawing for footing sizes.")
        return
    by = Counter(m[3] for m in marks)
    thk = args.thk / 1000.0
    ws = args.ws / 1000.0
    off = 0.1  # PCC offset beyond footing edge (m)
    pcc_t = args.pcc / 1000.0
    D = args.depth

    rows = []
    conc = exc = pcc = 0.0
    for sz, n in sorted(by.items(), key=lambda kv: -kv[1]):
        w, l = (int(v) / 1000.0 for v in sz.lower().split("x"))
        v_c = w * l * thk * n
        v_p = (w + 2*off) * (l + 2*off) * pcc_t * n
        v_e = (w + 2*ws) * (l + 2*ws) * D * n
        conc += v_c; pcc += v_p; exc += v_e
        rows.append([sz, n, f"{v_c:.2f}", f"{v_e:.1f}"])
    steel = conc * 7850 * args.steel / 100.0
    backfill = exc - conc - pcc
    print(f"FOUNDATION TAKEOFF  (auto from drawing -- {len(by)} sizes, "
          f"{sum(by.values())} footings)\n")
    print(fmt_table(rows, ["size (mm)", "nos", "concrete m3", "excavation m3"]))
    print(f"\nFooting concrete ({args.thk:.0f} mm thick) : {conc:,.2f} m3")
    print(f"Steel ({args.steel}%)                : {steel:,.0f} kg "
          f"({steel/1000:.2f} t)")
    print(f"PCC bed ({args.pcc:.0f} mm, +100 offset): {pcc:,.2f} m3 "
          f"(~{pcc*3.4:.0f} cement bags, 1:4:8)")
    print(f"Excavation (depth {D} m, +{args.ws:.0f} mm working space): "
          f"{exc:,.1f} m3")
    print(f"Backfill (excavation - footing - PCC): {backfill:,.1f} m3")
    print("\nNote: if counts look 2-3x high the sheet set is duplicated --")
    print("set the building's box once to isolate one copy.")
    print("Footing thickness/depth are site values (not in the plan) -- "
          "defaults shown, change in Advanced.")


def cmd_concrete(args):
    """Total RCC concrete read straight from the drawing -- columns (per
    storey) + floor slabs -- in one click, nothing typed. Reads the level
    marks and column grid itself. Optional --box (isolate one sheet),
    --layer (slab boundary), --thk (slab mm)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    levels = (parse_levels(args.levels) if getattr(args, "levels", None)
              else (extract_levels(msp) or dict(DEFAULT_LEVELS)))
    seq = [f for f in FLOOR_ORDER if f in levels]
    heights = {seq[i]: levels[seq[i + 1]] - levels[seq[i]]
               for i in range(len(seq) - 1)}
    box = _parse_box(getattr(args, "box", None))
    markers, titles = _markers_titles(msp, box)
    floor = defaultdict(Counter)
    for x, y, tag, sz in markers:
        nm = _floor_of(x, y, titles) if titles else (seq[0] if seq else "GROUND")
        floor[nm][sz] += 1
    col = 0.0
    by_size, counts = Counter(), Counter()
    for nm in FLOOR_ORDER[:-1]:
        h = heights.get(nm)
        if h is None or nm not in floor:
            continue
        for sz, n in floor[nm].items():
            v = n * _area(sz) * h
            col += v
            by_size[sz] += v
            counts[sz] += n
    res = _largest_closed(doc, msp, args.layer)
    nfloors = len(seq) or 1
    area, slab = 0.0, 0.0
    if res:
        area, _ = res
        slab = area * (args.thk / 1000.0) * nfloors
    total = col + slab

    print("CONCRETE  (auto from drawing -- sizes & heights read from the drawing)\n")
    if by_size:
        print("COLUMN SECTIONS (read from the drawing)")
        rows = [[sz, counts[sz], f"{by_size[sz]:.2f}"]
                for sz in sorted(by_size, key=lambda s: -_area(s))]
        print(fmt_table(rows, ["section (mm)", "count", "concrete m3"]))
        hh = [f"{seq[i]}->{seq[i+1]} {heights[seq[i]]:.1f}m"
              for i in range(len(seq) - 1)]
        if hh:
            print("Storey heights: " + ", ".join(hh))
        print()
    print("SUMMARY")
    print(fmt_table([["Columns (superstructure)", f"{col:.2f}"],
                     ["Floor slabs", f"{slab:.2f}"]],
                    ["item", "concrete m3"]))
    print(f"\nTotal RCC concrete: {total:.2f} m3")
    if area:
        print(f"(slab: {area:,.0f} m2 x {args.thk:.0f} mm x {nfloors} floors)")
    print("Sections & storey heights are taken from the drawing -- nothing typed.")
    print("Footings / plinth / chajja: use colvol, slab, plinth-colvol, footings, chajja.")


def _takeoff_lines(doc, msp, args):
    """Auto-read the big-ticket takeoff quantities as (item, qty, unit) rows.

    Columns and the footprint come from the GEOMETRY engine (rectangles +
    clustering) -- no tags, no layers, no manual box. Spec values that a
    drawing can't hold (floor height, mix, footing depth, steel kg/m3) use
    documented defaults, overridable in Advanced. Kept isolated from the
    single-tool commands so pricing can never break them."""
    lines = []
    d = _detect_geometry(msp)
    nfloors = d["floor_count"] or 1

    # Floor-to-floor height: real level marks if the drawing has them, else a
    # documented default (the one value that truly isn't in an early drawing).
    levels = extract_levels(msp)
    fh = getattr(args, "floor_height", None) or 3.5
    if levels:
        seq = [f for f in FLOOR_ORDER if f in levels]
        hs = [levels[seq[i + 1]] - levels[seq[i]] for i in range(len(seq) - 1)]
        if hs:
            fh = sum(hs) / len(hs)

    # Column concrete = section area x height, summed over the detected
    # schedule, times the number of floors. Sizes are the rectangle sizes.
    col_area = sum((w / 1000.0) * (h / 1000.0) * n
                   for (w, h), n in d["schedule"].items())
    col = col_area * fh * nfloors
    if col:
        lines.append(("RCC concrete (columns)", round(col, 2), "m3"))

    slab = d["footprint"] * (getattr(args, "thk", 150) / 1000.0) * nfloors
    if slab:
        lines.append(("RCC concrete (slab)", round(slab, 2), "m3"))

    box = _parse_box(getattr(args, "box", None))
    marks = _footing_markers(msp, box)
    fconc = exc = 0.0
    if marks:
        by = Counter(m[3] for m in marks)
        thk = getattr(args, "foot_thk", 450) / 1000.0
        ws = getattr(args, "ws", 300) / 1000.0
        D = getattr(args, "depth", 1.5)
        for sz, n in by.items():
            w, l = (int(v) / 1000.0 for v in sz.lower().split("x"))
            fconc += w * l * thk * n
            exc += (w + 2 * ws) * (l + 2 * ws) * D * n
        lines.append(("Footing concrete", round(fconc, 2), "m3"))
        lines.append(("Excavation", round(exc, 1), "m3"))

    # ponytail: uniform steel ratio over ALL concrete -- a real design varies
    # by member (footings ~80, columns ~150 kg/m3); tune with --steel-kg.
    total_conc = col + slab + fconc
    if total_conc:
        skg = getattr(args, "steel_kg", 100)
        lines.append(("Reinforcement steel", round(total_conc * skg / 1000.0, 3), "tonne"))

    def bcount(pat):
        r = re.compile(pat, re.I)
        return sum(1 for e in msp
                   if e.dxftype() == "INSERT" and r.search(e.dxf.name or ""))
    nd, nw = bcount("door"), bcount("window|glass|vp")
    if nd:
        lines.append(("Doors", nd, "nos"))
    if nw:
        lines.append(("Windows", nw, "nos"))
    return lines


def cmd_priced_boq(args):
    """Priced BOQ in one tap: quantities auto-read from the drawing x your
    rate library = an estimate, not just a takeoff. Quantities come from the
    drawing; rates are your editable list (rates.json). Items with no rate
    are listed unpriced so nothing is silently dropped."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    _seed_rates()
    rates = load_rates()
    lines = _takeoff_lines(doc, msp, args)
    if not lines:
        print("No auto-readable quantities found in this drawing.")
        print("Run 'drawing-check' to see which conventions it uses.")
        return
    rows, total, unpriced = [], 0.0, []
    for item, qty, unit in lines:
        r = rates.get(item)
        if r and r.get("rate"):
            amt = qty * r["rate"]
            total += amt
            rows.append([item, f"{qty:g}", unit, f"{r['rate']:,.0f}", f"{amt:,.0f}"])
        else:
            unpriced.append((item, qty, unit))
            rows.append([item, f"{qty:g}", unit, "-", "-"])
    print("PRICED BOQ  (quantities auto-read from the drawing)\n")
    print(fmt_table(rows, ["item", "qty", "unit", "rate Rs", "amount Rs"]))
    print(f"\nGRAND TOTAL: Rs {total:,.0f}")
    # Sanity nudge: a slab reading in the thousands of m3 almost always means
    # _largest_closed grabbed a plot/sheet boundary on a multi-sheet drawing.
    slab_m3 = next((q for it, q, u in lines if it == "RCC concrete (slab)"), 0)
    if slab_m3 > 2000:
        print("\nCAUTION: slab volume looks very high -- the footprint was read "
              "from the largest closed shape, likely a plot/sheet border. "
              "Set --box to isolate one building, or --layer <slab layer>.")
    if unpriced:
        print("\nUnpriced (no rate in library -- add one in rates.json):")
        for item, qty, unit in unpriced:
            print(f"  {item}: {qty:g} {unit}")
    print(f"\nRates from: {RATES_FILE}")
    print("Quantities are from the drawing; rates are your editable list. "
          "Spec values (slab/footing thickness, steel kg/m3) are defaults -- "
          "change them in Advanced.")


def cmd_rates(args):
    """Show the rate library (edit rates.json to change rates or add items).
    Items whose names match the priced-boq takeoff get priced automatically."""
    _seed_rates()
    rates = load_rates()
    rows = [[item, r.get("unit", ""), f"{r.get('rate', 0):,.0f}"]
            for item, r in rates.items()]
    print("RATE LIBRARY  (Rs per unit)\n")
    print(fmt_table(rows, ["item", "unit", "rate Rs"]))
    print(f"\nEdit this file to change rates or add items: {RATES_FILE}")


def cmd_colschedule(args):
    """Full column schedule: every tag (C1..Cn) with its section, plus totals
    per size."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    sched = _column_schedule(msp)
    if not sched:
        print("No column schedule found.")
        return
    rows = [[t, sched[t]] for t in sorted(sched, key=lambda x: int(x[1:]))]
    print(fmt_table(rows, ["tag", "section (mm)"]))
    by = Counter(sched.values())
    print()
    print(fmt_table([[s, n] for s, n in sorted(by.items(), key=lambda kv: -_area(kv[0]))],
                    ["section (mm)", "count"]))
    print(f"\nTotal column types: {len(sched)}")


def cmd_staircases(args):
    """Count staircases and lifts from their labels."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    stairs, lifts = set(), set()
    for s, *_ in all_text_entities(msp):
        u = s.strip().upper()
        if "DETAIL" in u:            # skip drawing/sheet titles
            continue
        m = re.search(r"STAIRCASE[-\s]*([0-9]+)", u)
        if m:
            stairs.add(m.group(1))
        if re.search(r"\bLIFT\b", u) and not any(
                k in u for k in ("STAIRCASE", "PARAPET", "LVL", "TOP")):
            lifts.add(re.sub(r"\s+", " ", u)[:28])
    print(f"Staircases: {len(stairs)}")
    if stairs:
        print("  " + ", ".join("STAIRCASE-" + s for s in sorted(stairs)))
    print(f"Lifts: {len(lifts)}")
    for l in sorted(lifts):
        print("  " + l)


def cmd_beams(args):
    """Count and total length of beams on a beam layer.

    DXFs only show beams if they sit on a dedicated layer. Pass --layer NAME;
    if omitted, auto-detects a layer whose name contains 'beam'. If none is
    found, lists the layers so you can pick.
    """
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    layer = args.layer
    if not layer:
        cand = [l.dxf.name for l in doc.layers if "beam" in l.dxf.name.lower()]
        if not cand:
            print("No 'beam' layer found. Re-run with --layer NAME. Layers:")
            print("  " + ", ".join(l.dxf.name for l in doc.layers))
            return
        layer = cand[0]
    n, total = 0, 0.0
    for e in msp:
        if (e.dxf.layer or "") != layer:
            continue
        if e.dxftype() == "LINE":
            a, b = e.dxf.start, e.dxf.end
            total += math.dist((a.x, a.y), (b.x, b.y))
            n += 1
        elif e.dxftype() == "LWPOLYLINE":
            n += 1
    print(f"Layer '{layer}': {n} beam entities, total run {total/1000:,.1f} m")


def cmd_footings(args):
    """Count footings from a footing layer or 'FOOTING' annotations.

    Pass --layer NAME, else auto-detects a layer containing 'foot'. Falls back
    to counting 'FOOTING' text labels if no footing layer exists.
    """
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    layer = args.layer
    if not layer:
        cand = [l.dxf.name for l in doc.layers if "foot" in l.dxf.name.lower()]
        layer = cand[0] if cand else None
    if layer:
        n = sum(1 for e in msp if (e.dxf.layer or "") == layer
                and e.dxftype() in ("INSERT", "LWPOLYLINE", "CIRCLE"))
        print(f"Layer '{layer}': {n} footing entities")
        return
    labels = sum(1 for s, *_ in all_text_entities(msp) if "FOOTING" in s.upper())
    print(f"No footing layer found. 'FOOTING' text labels: {labels}")
    print("(Pass --layer NAME if footings sit on a specific layer.)")


def cmd_heights(args):
    """All heights from the drawing's level marks: floor-to-floor, building
    height, plinth, parapet. Reads the DXF -- nothing assumed."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    lv = extract_levels(msp)
    if not lv:
        print("No 'FLOOR LVL' marks found in this drawing.")
        return
    order = ["ROAD", "GROUND", "FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH",
             "SIXTH", "TERRACE", "LIFTTOP", "LIFTPARAPET", "PARAPET"]
    print("LEVELS (m):")
    for k in order:
        if k in lv:
            print(f"  {k:12} {lv[k]:+.2f}")
    floors = [k for k in ["GROUND", "FIRST", "SECOND", "THIRD", "FOURTH",
                          "FIFTH", "SIXTH", "TERRACE"] if k in lv]
    if len(floors) > 1:
        print("\nFLOOR-TO-FLOOR (m):")
        for i in range(len(floors) - 1):
            print(f"  {floors[i]:8} -> {floors[i+1]:8} : "
                  f"{lv[floors[i+1]] - lv[floors[i]]:.2f}")
    base = lv.get("ROAD", lv.get("GROUND"))
    top = max(lv.values())
    print(f"\nBuilding height (base -> top): {top - base:.2f} m")
    if "GROUND" in lv and "ROAD" in lv:
        print(f"Plinth height: {lv['GROUND'] - lv['ROAD']:.2f} m")
    if "TERRACE" in lv:
        para = lv.get("PARAPET", lv.get("LIFTPARAPET", lv.get("LIFTTOP")))
        if para:
            print(f"Above terrace (parapet/lift): {para - lv['TERRACE']:.2f} m")


def cmd_builtup(args):
    """Built-up footprint = the largest plausible closed boundary. Skips a
    shape that's ~the whole sheet (the drawing border), and can target a
    specific --layer for a clean result."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    areas = []
    for e in msp:
        if e.dxftype() != "LWPOLYLINE" or not e.closed:
            continue
        if args.layer and (e.dxf.layer or "") != args.layer:
            continue
        try:
            a = abs(e.get_area())
        except Exception:
            a = _poly_area(e)
        areas.append(a / 1_000_000.0)
    if not areas:
        print("No closed polylines"
              + (f" on layer '{args.layer}'" if args.layer else "") + ".")
        return
    areas.sort(reverse=True)
    # detect & skip the sheet-border-sized shape
    sheet = None
    ext = (doc.header.get("$EXTMIN"), doc.header.get("$EXTMAX"))
    if ext[0] and ext[1]:
        sheet = abs((ext[1][0]-ext[0][0]) * (ext[1][1]-ext[0][1])) / 1_000_000.0
    pick, note = areas[0], ""
    if sheet and areas[0] > 0.6 * sheet:
        smaller = [a for a in areas if a < 0.6 * sheet]
        if smaller:
            pick, note = smaller[0], "  (skipped sheet-border-sized shape)"
    print(f"Footprint (largest plausible closed area): {pick:,.2f} m2{note}")
    print("Top closed areas (m2): " + ", ".join(f"{a:,.1f}" for a in areas[:6]))
    if args.layer:
        print(f"(layer '{args.layer}')")
    else:
        print("Tip: pass --layer <wall/boundary layer> for a cleaner footprint.")


def cmd_doors(args):
    """Count door blocks. Matches INSERT block names containing 'door' (or
    --name PATTERN)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    pat = re.compile(args.name or "door", re.I)
    c = Counter(e.dxf.name for e in msp
                if e.dxftype() == "INSERT" and pat.search(e.dxf.name or ""))
    if not c:
        print(f"No blocks matching '{args.name or 'door'}'. "
              "Try the 'blocks' tool to see block names, then --name.")
        return
    for nm, n in c.most_common():
        print(f"  {nm}: {n}")
    print(f"Total: {sum(c.values())}")


def cmd_windows(args):
    """Count window blocks. Matches INSERT names containing window/glass/vp
    (or --name PATTERN)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    pat = re.compile(args.name or "window|glass|vp", re.I)
    c = Counter(e.dxf.name for e in msp
                if e.dxftype() == "INSERT" and pat.search(e.dxf.name or ""))
    if not c:
        print(f"No blocks matching '{args.name or 'window|glass|vp'}'. "
              "Use 'blocks' to see names, then --name.")
        return
    for nm, n in c.most_common():
        print(f"  {nm}: {n}")
    print(f"Total: {sum(c.values())}")


def cmd_wall_area(args):
    """Wall face area for plaster/paint = wall run on a layer x wall height.

    Pass --layer (the wall layer) and --height (metres). Reports one-side and
    both-sides area."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "line")
    if not args.layer:
        print("No wall/line geometry found to measure.")
        return
    L = 0.0
    for e in msp:
        if (e.dxf.layer or "") != args.layer:
            continue
        if e.dxftype() == "LINE":
            a, b = e.dxf.start, e.dxf.end
            L += math.dist((a.x, a.y), (b.x, b.y))
        elif e.dxftype() == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            for i in range(len(pts) - 1):
                L += math.dist(pts[i], pts[i + 1])
    lm, h = L / 1000.0, args.height
    print(f"Layer '{args.layer}': wall run {lm:,.1f} m  x height {h} m")
    print(f"One side : {lm*h:,.1f} m2")
    print(f"Both sides: {lm*h*2:,.1f} m2")


def cmd_findtext(args):
    """Find drawing text containing a keyword; show matches + positions (m)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    q = (args.q or "").lower()
    if not q:
        print("Pass --q KEYWORD (e.g. --q tank).")
        return
    hits, seen = [], set()
    for s, x, y, lay in all_text_entities(msp):
        if q in s.lower():
            label = re.sub(r"\s+", " ", s.strip())[:40]
            key = (label, round(x / 1000, 1), round(y / 1000, 1))
            if key in seen:
                continue
            seen.add(key)
            hits.append([label, f"{x/1000:.1f},{y/1000:.1f}", lay])
    if not hits:
        print(f"No text matching '{args.q}'.")
        return
    print(fmt_table(hits[:args.limit], ["text", "pos (m)", "layer"]))
    print(f"\n{len(hits)} unique matches"
          + (f" (showing {args.limit})" if len(hits) > args.limit else ""))


def cmd_layer_detail(args):
    """Entity-type breakdown for ONE layer."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "busy")
    if not args.layer:
        print("This drawing has no geometry.")
        return
    c = Counter(e.dxftype() for e in msp if (e.dxf.layer or "") == args.layer)
    if not c:
        print(f"Layer '{args.layer}' is empty or missing.")
        return
    print(fmt_table([[t, n] for t, n in c.most_common()], ["entity", "count"]))
    print(f"Total: {sum(c.values())}")


def _layer_points(e):
    t = e.dxftype()
    if t == "LINE":
        return [(e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)]
    if t == "LWPOLYLINE":
        return [(p[0], p[1]) for p in e.get_points()]
    if t == "CIRCLE":
        return [(e.dxf.center.x, e.dxf.center.y)]
    try:
        return [(e.dxf.insert.x, e.dxf.insert.y)]
    except Exception:
        return []


def cmd_bbox_layer(args):
    """Bounding box + size (m) of all geometry on a layer (e.g. footprint from
    the column or wall layer)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "area")
    if not args.layer:
        print("This drawing has no geometry.")
        return
    xs, ys = [], []
    for e in msp:
        if (e.dxf.layer or "") != args.layer:
            continue
        for px, py in _layer_points(e):
            xs.append(px)
            ys.append(py)
    if not xs:
        print(f"No geometry on layer '{args.layer}'.")
        return
    w, h = (max(xs) - min(xs)) / 1000, (max(ys) - min(ys)) / 1000
    print(f"Layer '{args.layer}' bounds: {w:,.2f} x {h:,.2f} m  "
          f"(box area {w*h:,.1f} m2)")


def cmd_column_spacing(args):
    """Typical column grid spacing (nearest-neighbour, in m) for a floor."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    box = _parse_box(getattr(args, "box", None))
    markers, titles = _markers_titles(msp, box)
    floor = (args.floor or "GROUND").upper()
    pts = []
    for x, y, tag, sz in markers:
        if titles and _floor_of(x, y, titles) != floor:
            continue
        pts.append((round(x, 1), round(y, 1)))
    pts = list(set(pts))
    if len(pts) < 2:
        print("Not enough columns to measure spacing.")
        return
    nn = []
    for i, p in enumerate(pts):
        nn.append(min(math.dist(p, q) for j, q in enumerate(pts) if j != i) / 1000)
    nn.sort()
    print(f"Floor {floor}: {len(pts)} columns")
    print(f"Spacing (m) -- min {min(nn):.2f}, median {nn[len(nn)//2]:.2f}, "
          f"max {max(nn):.2f}")


def cmd_room_areas(args):
    """Area of each closed room on a layer, labelled by the nearest text."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "area")
    if not args.layer:
        print("No closed room shapes found in this drawing.")
        return
    texts = [(s.strip(), x, y) for s, x, y, _ in all_text_entities(msp) if s.strip()]
    rooms = []
    for e in msp:
        if (e.dxf.layer or "") != args.layer or e.dxftype() != "LWPOLYLINE" \
                or not e.closed:
            continue
        try:
            a = abs(e.get_area())
        except Exception:
            a = _poly_area(e)
        pts = [(p[0], p[1]) for p in e.get_points()]
        cx = sum(p[0] for p in pts) / len(pts)
        cy = sum(p[1] for p in pts) / len(pts)
        am = a / 1_000_000.0
        if am < args.min or am > args.max:   # drop tiny noise / huge borders
            continue
        lbl, bd = "?", 1e30
        for t, tx, ty in texts:
            d = (tx - cx) ** 2 + (ty - cy) ** 2
            if d < bd:
                bd, lbl = d, t[:24]
        rooms.append((lbl, am))
    if not rooms:
        print(f"No closed rooms on layer '{args.layer}' "
              f"between {args.min}-{args.max} m2.")
        return
    rooms.sort(key=lambda r: -r[1])
    print(fmt_table([[r[0], f"{r[1]:,.2f}"] for r in rooms[:args.limit]],
                    ["room (nearest label)", "area m2"]))
    print(f"\n{len(rooms)} rooms, total {sum(r[1] for r in rooms):,.1f} m2")


def cmd_layer_area(args):
    """Total closed-polyline area on ONE layer (m2)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "area")
    if not args.layer:
        print("No closed shapes found in this drawing.")
        return
    tot, n = 0.0, 0
    for e in msp:
        if (e.dxf.layer or "") != args.layer or e.dxftype() != "LWPOLYLINE" \
                or not e.closed:
            continue
        try:
            tot += abs(e.get_area())
        except Exception:
            tot += _poly_area(e)
        n += 1
    print(f"Layer '{args.layer}': {n} closed shapes, total {tot/1_000_000.0:,.2f} m2")


def cmd_layer_length(args):
    """Total line/polyline length on ONE layer (m) -- pipes, chajja, kerbs."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "line")
    if not args.layer:
        print("No line geometry found to measure.")
        return
    L = 0.0
    for e in msp:
        if (e.dxf.layer or "") != args.layer:
            continue
        if e.dxftype() == "LINE":
            a, b = e.dxf.start, e.dxf.end
            L += math.dist((a.x, a.y), (b.x, b.y))
        elif e.dxftype() == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            for i in range(len(pts) - 1):
                L += math.dist(pts[i], pts[i + 1])
            if e.closed and len(pts) > 2:
                L += math.dist(pts[-1], pts[0])
    print(f"Layer '{args.layer}': total length {L/1000:,.2f} m")


def cmd_entity_count(args):
    """Count entities of a DXF type (LINE, CIRCLE, MTEXT, INSERT...); no type =
    full breakdown."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    if not args.type:
        c = Counter(e.dxftype() for e in msp)
        print(fmt_table([[t, n] for t, n in c.most_common()], ["type", "count"]))
        return
    t = args.type.upper()
    print(f"{t}: {sum(1 for e in msp if e.dxftype() == t)}")


def cmd_fixtures(args):
    """Count fixtures = blocks on a layer (--layer sanitary/furniture) or whose
    name matches --name."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    if args.layer:
        c = Counter(e.dxf.name for e in msp if e.dxftype() == "INSERT"
                    and (e.dxf.layer or "") == args.layer)
    elif args.name:
        pat = re.compile(args.name, re.I)
        c = Counter(e.dxf.name for e in msp if e.dxftype() == "INSERT"
                    and pat.search(e.dxf.name or ""))
    else:
        print("Pass --layer or --name.")
        return
    if not c:
        print("No matching blocks. Use 'blocks' to see names/layers.")
        return
    for nm, n in c.most_common():
        print(f"  {nm}: {n}")
    print(f"Total: {sum(c.values())}")


def _hatch_area(e):
    """Approximate a HATCH area from its polyline boundary paths (mm^2)."""
    area = 0.0
    try:
        for path in e.paths:
            pts = [(v[0], v[1]) for v in getattr(path, "vertices", [])]
            if len(pts) >= 3:
                a = 0.0
                for i in range(len(pts)):
                    x1, y1 = pts[i]
                    x2, y2 = pts[(i + 1) % len(pts)]
                    a += x1 * y2 - x2 * y1
                area += abs(a) / 2.0
    except Exception:
        pass
    return area


# ======================================================================
# DRAWING QA / HEALTH SUITE  --  one-tap error finding (read-only)
# ----------------------------------------------------------------------
# Finds the things architects hunt by zooming around for hours: duplicate
# lines, stray fragments, overlapping text, duplicate dimensions, unclosed
# boundaries, blocks on wrong layers, empty layers. NOTHING is modified --
# these tools only report, so the drawing is never at risk.
# ======================================================================
def cmd_qa_report(args):
    """One-tap drawing health report: duplicate lines, stray fragments,
    text overlaps, duplicate dimensions, unclosed polylines, blocks on
    layer 0, empty layers. Read-only -- nothing in the drawing changes."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()

    lines = {}
    dup_lines = 0
    stray = 0
    unclosed = 0
    zero_blk = 0
    texts = []
    dims = {}
    dup_dims = 0
    counts = Counter()
    for e in msp:
        t = e.dxftype()
        counts[t] += 1
        if t == "LINE":
            a, b = e.dxf.start, e.dxf.end
            L = math.dist((a.x, a.y), (b.x, b.y))
            if L < 10:  # under 10 drawing units (~1cm) = stray fragment
                stray += 1
            key = tuple(sorted([(round(a.x, 1), round(a.y, 1)),
                                (round(b.x, 1), round(b.y, 1))]))
            if key in lines:
                dup_lines += 1
            lines[key] = 1
        elif t == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            if not e.closed and len(pts) >= 3:
                # start & end nearly touch -> meant to be closed
                if math.dist(pts[0], pts[-1]) < 50:
                    unclosed += 1
        elif t in ("TEXT", "MTEXT"):
            s = entity_text(e) or ""
            p = e.dxf.insert
            h = float(getattr(e.dxf, "height", 0) or
                      getattr(e.dxf, "char_height", 0) or 250)
            texts.append((p.x, p.y, h, s))
        elif t == "DIMENSION":
            try:
                d = e.dxf.defpoint
                key = (round(d.x, 0), round(d.y, 0))
                if key in dims:
                    dup_dims += 1
                dims[key] = 1
            except Exception:
                pass
        elif t == "INSERT" and (e.dxf.layer or "") in ("0", ""):
            zero_blk += 1

    # text overlaps via spatial bucketing (avoid O(n^2) on big drawings)
    buckets = defaultdict(list)
    overlaps = 0
    samples = []
    for x, y, h, s in texts:
        k = (int(x // 1000), int(y // 1000))
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for ox, oy, oh, os_ in buckets.get((k[0]+dx, k[1]+dy), []):
                    if abs(x-ox) < max(h, oh)*1.2 and abs(y-oy) < max(h, oh)*0.9:
                        overlaps += 1
                        if len(samples) < 5 and s.strip() and os_.strip():
                            samples.append(f"'{s.strip()[:24]}' / '{os_.strip()[:24]}'")
                        break
        buckets[k].append((x, y, h, s))

    ec = Counter((e.dxf.layer or "") for e in msp)
    empty = [l.dxf.name for l in doc.layers
             if ec.get(l.dxf.name, 0) == 0 and l.dxf.name not in ("0", "Defpoints")]

    issues = [
        ["Duplicate lines (same start/end)", dup_lines],
        ["Stray fragments (lines < 10 units)", stray],
        ["Nearly-closed polylines (gap < 50)", unclosed],
        ["Overlapping text pairs", overlaps],
        ["Duplicate dimensions", dup_dims],
        ["Blocks on layer 0", zero_blk],
        ["Empty layers", len(empty)],
    ]
    total = sum(n for _, n in issues)
    print("DRAWING HEALTH REPORT  (read-only -- nothing was changed)\n")
    print(fmt_table([[k, v] for k, v in issues], ["check", "found"]))
    if samples:
        print("\nText overlap samples:")
        for s in samples:
            print("  " + s)
    if empty:
        print("\nEmpty layers: " + ", ".join(empty[:15])
              + (" ..." if len(empty) > 15 else ""))
    print(f"\nEntities scanned: {sum(counts.values()):,}   "
          f"Issues found: {total}")
    print("Verdict: " + ("CLEAN — no obvious drawing errors." if total == 0
          else "Attention — review the counts above in AutoCAD."))


# ======================================================================
# Geometry engine  --  read a drawing with NO tags and NO layers.
# ----------------------------------------------------------------------
# The takeoff tools historically needed tagged text (C1 + 750x750), floor
# titles and clean layers. Real early-stage drawings have none of that: a
# column is just a small rectangle drawn to scale, walls are lines, and one
# DXF often carries several floor-plan sheets side by side. This engine reads
# the SHAPES instead of the labels: a column's section is the size of its
# rectangle, the grid is the spacing of their centres, and the separate sheets
# are found by spatially clustering the columns. Nothing typed, nothing tagged.
# ======================================================================
def _all_rects(msp):
    """Every closed 4-5 point polyline as (w, h, cx, cy) in drawing units."""
    out = []
    for e in msp:
        if e.dxftype() == "LWPOLYLINE" and e.closed:
            pts = [(p[0], p[1]) for p in e.get_points()]
            if 4 <= len(pts) <= 5:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                w, h = max(xs) - min(xs), max(ys) - min(ys)
                if w > 0 and h > 0:
                    out.append((w, h, (max(xs)+min(xs))/2, (max(ys)+min(ys))/2))
    return out


def _col_rects(rects, lo=200, hi=1000, aspect=3.0):
    """Rectangles that look like column sections: plausible size, not a sliver."""
    return [r for r in rects
            if lo <= r[0] <= hi and lo <= r[1] <= hi
            and max(r[0], r[1]) / min(r[0], r[1]) < aspect]


def _cluster(points, thresh):
    """Union-find clustering: connect points within `thresh` units. Grid-bucketed
    so it stays roughly O(n) instead of O(n^2). Returns list of index lists."""
    n = len(points)
    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    buckets = defaultdict(list)
    for i, (x, y) in enumerate(points):
        buckets[(int(x // thresh), int(y // thresh))].append(i)
    for (bx, by), idxs in buckets.items():
        neigh = []
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                neigh += buckets.get((bx + dx, by + dy), [])
        for i in idxs:
            xi, yi = points[i]
            for j in neigh:
                if j <= i:
                    continue
                xj, yj = points[j]
                if math.hypot(xi - xj, yi - yj) <= thresh:
                    ra, rb = find(i), find(j)
                    if ra != rb:
                        parent[ra] = rb
    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(i)
    return list(groups.values())


def _hull_area(pts):
    """Convex-hull area (m2) of centre points (mm) -- the footprint the columns
    enclose, immune to a stray plot/sheet border far away."""
    pts = sorted(set(pts))
    if len(pts) < 3:
        return 0.0

    def cross(o, a, b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])

    lower, upper = [], []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    poly = lower[:-1] + upper[:-1]
    a = 0.0
    for i in range(len(poly)):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % len(poly)]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2 / 1e6


def _hull_perimeter(pts):
    """Convex-hull perimeter (m) of centre points (mm)."""
    pts = sorted(set(pts))
    if len(pts) < 3:
        return 0.0

    def cross(o, a, b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])

    lower, upper = [], []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    poly = lower[:-1] + upper[:-1]
    return sum(math.dist(poly[i], poly[(i + 1) % len(poly)])
               for i in range(len(poly))) / 1000.0


def _median_spacing(cents):
    """Median nearest-neighbour distance (m) -- the typical column grid."""
    sp = []
    for i, (x, y) in enumerate(cents):
        best = 1e18
        for j, (x2, y2) in enumerate(cents):
            if i == j:
                continue
            d = math.hypot(x - x2, y - y2)
            if 50 < d < best:
                best = d
        if best < 1e17:
            sp.append(best)
    sp.sort()
    return (sp[len(sp) // 2] / 1000) if sp else 0.0


def _dedupe_rects(cols, cell=300):
    """Collapse stacked / overlapping rectangles at the same spot to one.
    A schedule table or legend draws many boxes on top of each other (nearest-
    neighbour ~0); without this they'd be counted as dozens of 'columns'."""
    seen, out = set(), []
    for w, h, cx, cy in cols:
        key = (round(cx / cell), round(cy / cell))
        if key in seen:
            continue
        seen.add(key)
        out.append((w, h, cx, cy))
    return out


# ponytail: 12 m cluster gap + >=8 cols/sheet are heuristics tuned on a real
# admin-block drawing. Two buildings closer than 12 m would merge; a shed with
# <8 columns needs min_cols lowered. Expose both as knobs if a drawing misreads.
def _detect_geometry(msp, cluster_m=12.0, min_cols=8):
    """Read a building from raw geometry -- no tags, no layers, no manual box.

    Columns are rectangles; sheets/floors are spatial clusters of columns;
    the footprint is the hull the columns enclose. Returns a dict describing
    the building's typical floor, plus the raw clusters. Everything here comes
    from shapes and coordinates -- nothing typed on the drawing."""
    cols = _dedupe_rects(_col_rects(_all_rects(msp)))
    cents = [(c[2], c[3]) for c in cols]
    groups = [g for g in _cluster(cents, cluster_m * 1000) if len(g) >= min_cols]
    groups.sort(key=len, reverse=True)

    clusters = []
    for g in groups:
        gcols = [cols[i] for i in g]
        gcents = [cents[i] for i in g]
        clusters.append({
            "n": len(gcols),
            "sizes": Counter((round(w), round(h)) for w, h, _, _ in gcols),
            "footprint": _hull_area(gcents),
            "perimeter": _hull_perimeter(gcents),
            "grid": _median_spacing(gcents),
        })
    if not clusters:
        return {"clusters": [], "floor_count": 0, "columns_per_floor": 0,
                "footprint": 0.0, "grid": 0.0, "schedule": Counter(),
                "total_columns": len(cols)}

    # Building floors = clusters with a real footprint on a normal structural
    # grid (2.5-11 m). Tight-grid / small clusters are tables, legends or detail
    # blocks -> NOT columns. If none qualify we say so rather than pretend a
    # clump of table cells is a floor.
    floors = [c for c in clusters
              if c["footprint"] >= 100 and 2.5 <= c["grid"] <= 11]
    if not floors:
        return {"clusters": clusters, "floor_count": 0, "columns_per_floor": 0,
                "footprint": 0.0, "grid": 0.0, "schedule": Counter(),
                "total_columns": len(cols), "no_grid": True}
    # keep only those near the biggest floor footprint (repeated plans of the
    # SAME building), so detail/schedule blocks don't inflate the floor count.
    fmax = max(c["footprint"] for c in floors)
    floors = [c for c in floors if c["footprint"] >= 0.55 * fmax]
    rep = max(floors, key=lambda c: c["n"])   # richest plan = best schedule
    fpts = sorted(c["footprint"] for c in floors)
    return {
        "clusters": clusters,
        "floor_count": len(floors),
        "columns_per_floor": int(sum(c["n"] for c in floors) / len(floors)),
        "footprint": fpts[len(fpts) // 2],     # median floor footprint
        "perimeter": rep["perimeter"],
        "grid": rep["grid"],
        "schedule": rep["sizes"],
        "total_columns": len(cols),
    }


def cmd_scan(args):
    """ONE-TAP takeoff from raw geometry -- no tags, no layers, no setup.
    Reads columns (as rectangles), auto-separates the floor-plan sheets in the
    file, and reports the building: columns per floor, floor count, footprint,
    grid and column schedule. Works on an untagged, unlayered early drawing."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    d = _detect_geometry(msp,
                         cluster_m=getattr(args, "gap", None) or 12.0,
                         min_cols=getattr(args, "mincols", None) or 8)
    if not d["clusters"]:
        print("No column-like rectangles found. If columns are drawn as blocks "
              "or circles, tell me and I'll widen the detector.")
        return
    if d.get("no_grid"):
        print("ONE-TAP SCAN  (read from the drawing's geometry -- no tags, no layers)\n")
        print(f"Found {d['total_columns']} column-like rectangles, but they don't "
              "form a real structural grid --")
        print("they sit in tight clumps, which usually means a schedule table, a "
              "legend, or a")
        print("detail block rather than a to-scale floor plan.\n")
        print("So this looks like an EARLY / detail drawing, not a plan drawn to "
              "scale. Counts and")
        print("sizes here can't be trusted yet -- once the plan is drawn to scale "
              "(columns spaced")
        print("metres apart on a grid), scan will read it straight off the shapes.")
        return
    print("ONE-TAP SCAN  (read from the drawing's geometry -- no tags, no layers)\n")
    print(f"Building floor plans detected : {d['floor_count']}")
    print(f"Columns per floor (typical)   : {d['columns_per_floor']}")
    print(f"Floor footprint               : {d['footprint']:,.0f} m2")
    print(f"Column grid (typical spacing) : {d['grid']:.2f} m")
    print(f"Total column rectangles read  : {d['total_columns']}\n")
    rows = [[f"{w} x {h}", n] for (w, h), n in d["schedule"].most_common(12)]
    print("COLUMN SCHEDULE (sizes read from the rectangles themselves)")
    print(fmt_table(rows, ["section (mm)", "count on the plan"]))
    # honesty grade
    std = {(230, 230), (230, 300), (300, 300), (300, 450), (450, 450),
           (450, 600), (525, 525), (600, 600), (600, 750), (750, 750)}
    hit = sum(n for sz, n in d["schedule"].items() if sz in std)
    frac = hit / max(1, sum(d["schedule"].values()))
    grade = ("TO-SCALE -- sizes match standard sections, trust the numbers"
             if frac > 0.4 else
             "ROUGH -- column boxes may be placeholders; counts/grid are good, "
             "confirm sizes before pricing")
    print(f"\nQuality: {grade}")
    print("Other floors/detail blocks in the file were separated automatically; "
          "nothing was tagged or layered.")


def cmd_drawing_check(args):
    """Scan a DXF and report which tagging CONVENTIONS it uses, so you know
    up-front which auto takeoff tools will work on it -- a trust check before
    you rely on any number. Read-only. The layer/QA/calculator tools work on
    ANY drawing; only the structural takeoff tools below need the tags."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()

    col, foot, floors, lvls, opens = set(), set(), set(), set(), set()
    open_pat = re.compile(r"^(D\d+|V-?\d+|VP\d+|FVP\d+|ED\d+|HD|HW|RS-?\d*|SW\d*|OW|LW)$")
    for s, *_ in all_text_entities(msp):
        st = s.strip()
        m = _SIZEPAT.match(st)
        if m:
            col.add(f"{m.group(1)} {m.group(2)}x{m.group(3)}")
        m = _FOOTPAT.match(st)
        if m:
            foot.add(f"{m.group(1)} {m.group(2)}x{m.group(3)}")
        u = re.sub(r"\s+", " ", st.upper())
        fm = _FLOORPAT.search(u)
        if fm:
            floors.add(fm.group(0))
        if _LVLPAT.search(u):
            lvls.add(u[:30])
        for line in st.split("\n"):
            if open_pat.match(line.strip()):
                opens.add(line.strip())

    bnames = [(e.dxf.name or "") for e in msp if e.dxftype() == "INSERT"]
    def blk(pat):
        r = re.compile(pat, re.I)
        return sorted({b for b in bnames if b and r.search(b)})
    doors, wins = blk("door"), blk("window|glass|vp")
    sani, park = blk(r"wc|toilet|urinal|basin|sink|sanit|ewc|wash"), blk(r"park|car|ecs|two.?wheel")
    nlayers = len(list(doc.layers))

    conv = [
        ("Column tags (C1 + 750x750)", col,
         "columns, colvol, floorcols, colschedule, concrete, perimeter/plinth-columns, column-spacing"),
        ("Footing tags (F1 + 1800x1800)", foot, "foundation, footing-schedule"),
        ("Floor-plan titles (GROUND FLOOR PLAN)", floors, "floorcols, per-floor colvol"),
        ("Level marks (GROUND FLOOR LVL +1.20)", lvls, "levels, heights, slab, floor counts, colvol elevations"),
        ("Opening tags (D1 1000x2100)", opens, "schedule, door-width-check"),
        ("Named door blocks (DOOR_900)", doors, "doors"),
        ("Named window blocks", wins, "windows"),
        ("Named sanitary blocks (WC)", sani, "sanitary"),
        ("Named parking blocks (CAR_PARK)", park, "parking"),
    ]
    rows = [[name, (f"yes ({len(found)})" if found else "no"), tools]
            for name, found, tools in conv]
    present = sum(1 for _, f, _ in conv if f)

    print("DRAWING CONVENTION CHECK  (read-only -- nothing was changed)\n")
    print(fmt_table(rows, ["convention", "found", "tools it unlocks"]))
    samples = []
    for name, found, _ in conv:
        if found:
            ex = ", ".join(sorted(found)[:4])
            samples.append(f"  {name.split(' (')[0]}: {ex}"
                           + (" ..." if len(found) > 4 else ""))
    if samples:
        print("\nExamples read from the drawing:")
        print("\n".join(samples))
    print(f"\nLayers in drawing: {nlayers}. The ~45 layer-analysis, drawing-QA "
          "and calculator tools work on ANY drawing regardless of the above.")
    print(f"\nConventions present: {present}/9")
    if present == 0:
        print("Verdict: no standard tags found -- the structural takeoff tools "
              "won't auto-read this drawing. Layer/QA/calculator tools still work.")
    elif present >= 6:
        print("Verdict: STRONG -- this drawing follows the conventions; the "
              "structural takeoff tools will work well.")
    else:
        print("Verdict: PARTIAL -- the tools for the ticked conventions will "
              "work; the rest need their tags added or a manual calculator.")


def cmd_dxf_diff(args):
    """Compare two drawings / revisions: what was added, removed or changed.
    Give two buildings (--building A --other B) or two file paths. Reports
    per-layer entity changes, text added/removed, and block count changes."""
    other = args.other
    proj = load_project()
    ob = proj.get("buildings", {}).get(other)
    path_b = ob["dxf"] if ob else other
    doc_a, doc_b = load_dxf(args.file), load_dxf(path_b)
    msp_a, msp_b = doc_a.modelspace(), doc_b.modelspace()

    def stats(msp):
        lay = Counter((e.dxf.layer or "") for e in msp)
        blk = Counter(e.dxf.name for e in msp if e.dxftype() == "INSERT")
        txt = Counter((entity_text(e) or "").strip() for e in msp
                      if e.dxftype() in ("TEXT", "MTEXT"))
        txt.pop("", None)
        return lay, blk, txt

    la, ba, ta = stats(msp_a)
    lb, bb, tb = stats(msp_b)

    rows = []
    for lay in sorted(set(la) | set(lb)):
        d = lb.get(lay, 0) - la.get(lay, 0)
        if d:
            rows.append([lay[:38], la.get(lay, 0), lb.get(lay, 0),
                         f"{d:+d}"])
    print("REVISION COMPARE   A = current building, B = other\n")
    if rows:
        print("LAYER CHANGES (entity counts)")
        print(fmt_table(rows, ["layer", "A", "B", "change"]))
    else:
        print("No per-layer entity count changes.")
    brows = [[nm[:38], ba.get(nm, 0), bb.get(nm, 0),
              f"{bb.get(nm,0)-ba.get(nm,0):+d}"]
             for nm in sorted(set(ba) | set(bb))
             if ba.get(nm, 0) != bb.get(nm, 0)]
    if brows:
        print("\nBLOCK CHANGES (doors/equipment moved-in/out show here)")
        print(fmt_table(brows, ["block", "A", "B", "change"]))
    added = [s for s in tb if s not in ta][:15]
    removed = [s for s in ta if s not in tb][:15]
    if added:
        print("\nTEXT ADDED in B: " + "; ".join(s[:30] for s in added))
    if removed:
        print("TEXT REMOVED in B: " + "; ".join(s[:30] for s in removed))
    if not (rows or brows or added or removed):
        print("\nDrawings look identical by these measures.")


def cmd_attr_audit(args):
    """Block attribute audit: blocks whose attributes are missing or empty,
    plus blocks placed on layer 0 (usually wrong). Read-only."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    missing = Counter()
    empty_att = Counter()
    zero_layer = Counter()
    total = 0
    # which block definitions declare attributes
    has_attdef = set()
    for b in doc.blocks:
        for e in b:
            if e.dxftype() == "ATTDEF":
                has_attdef.add(b.name)
                break
    for e in msp:
        if e.dxftype() != "INSERT":
            continue
        total += 1
        nm = e.dxf.name or "?"
        attribs = list(getattr(e, "attribs", []) or [])
        if nm in has_attdef and not attribs:
            missing[nm] += 1
        for a in attribs:
            if not (a.dxf.text or "").strip():
                empty_att[nm] += 1
        if (e.dxf.layer or "") in ("0", ""):
            zero_layer[nm] += 1
    print(f"BLOCK ATTRIBUTE AUDIT   ({total} block inserts scanned)\n")
    if missing:
        print("Blocks missing their attributes:")
        print(fmt_table([[n, c] for n, c in missing.most_common(15)],
                        ["block", "count"]))
    if empty_att:
        print("\nBlocks with EMPTY attribute values:")
        print(fmt_table([[n, c] for n, c in empty_att.most_common(15)],
                        ["block", "empty values"]))
    if zero_layer:
        print("\nBlocks on layer 0 (probably wrong layer):")
        print(fmt_table([[n, c] for n, c in zero_layer.most_common(15)],
                        ["block", "count"]))
    if not (missing or empty_att or zero_layer):
        print("CLEAN — every block has filled attributes and a proper layer.")


def cmd_parking(args):
    """Count parking spaces from their blocks (names matching park/car/ecs,
    or --name PATTERN)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    pat = re.compile(args.name or r"park|car|ecs|two.?wheel", re.I)
    c = Counter(e.dxf.name for e in msp
                if e.dxftype() == "INSERT" and pat.search(e.dxf.name or ""))
    if not c:
        print("No parking blocks found. Use 'blocks' to see names, then --name.")
        return
    for nm, n in c.most_common():
        print(f"  {nm}: {n}")
    print(f"Total parking spaces: {sum(c.values())}")


def cmd_sanitary(args):
    """Count toilets / sanitary fixtures from their blocks (wc/toilet/
    urinal/basin/sink, or --name PATTERN)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    pat = re.compile(args.name or r"wc|toilet|urinal|basin|sink|sanit|ewc|wash",
                     re.I)
    c = Counter(e.dxf.name for e in msp
                if e.dxftype() == "INSERT" and pat.search(e.dxf.name or ""))
    if not c:
        print("No sanitary blocks found. Use 'blocks' to see names, then --name.")
        return
    for nm, n in c.most_common():
        print(f"  {nm}: {n}")
    print(f"Total sanitary fixtures: {sum(c.values())}")


def cmd_area_statement(args):
    """One-tap area statement: built-up footprint + perimeter + closed areas
    per layer -- the numbers an area statement needs, straight off the plan."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    res = _largest_closed(doc, msp, args.layer)
    print("AREA STATEMENT  (auto from drawing)\n")
    if res:
        area, per = res
        print(f"Built-up footprint : {area:,.2f} m2")
        print(f"Perimeter          : {per:,.1f} m")
    lv = extract_levels(msp)
    seq = [f for f in FLOOR_ORDER if f in lv]
    if seq and res:
        print(f"Floors (from levels): {len(seq)}  ({', '.join(seq)})")
        print(f"Total built-up (footprint x floors): {res[0]*len(seq):,.2f} m2")
    per_layer = defaultdict(float)
    for e in msp:
        if e.dxftype() == "LWPOLYLINE" and e.closed:
            try:
                a = abs(e.get_area()) / 1e6
            except Exception:
                a = _poly_area(e) / 1e6
            if 1 <= a <= 100000:
                per_layer[e.dxf.layer or ""] += a
    rows = sorted(per_layer.items(), key=lambda kv: -kv[1])[:12]
    if rows:
        print("\nCLOSED AREAS PER LAYER (rooms/zones live on their layers)")
        print(fmt_table([[l[:38], f"{a:,.1f}"] for l, a in rows],
                        ["layer", "area m2"]))
    print("\nTip (optional): choose a room-boundary layer in Advanced for an")
    print("exact room-by-room table (the 'room-areas' tool).")


def cmd_layer_report(args):
    """Master per-LAYER breakdown across EVERY layer in one table:
    entities, line length (m), closed area (m2), circles, blocks, text,
    hatch area (m2). The one-shot 'what's on every layer' view."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    rep = defaultdict(lambda: {"ent": 0, "len": 0.0, "area": 0.0,
                               "circ": 0, "blk": 0, "txt": 0, "hatch": 0.0})
    for e in msp:
        r = rep[e.dxf.layer or ""]
        r["ent"] += 1
        t = e.dxftype()
        if t == "LINE":
            a, b = e.dxf.start, e.dxf.end
            r["len"] += math.dist((a.x, a.y), (b.x, b.y))
        elif t == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            for i in range(len(pts) - 1):
                r["len"] += math.dist(pts[i], pts[i + 1])
            if e.closed:
                if len(pts) > 2:
                    r["len"] += math.dist(pts[-1], pts[0])
                    try:
                        r["area"] += abs(e.get_area())
                    except Exception:
                        r["area"] += _poly_area(e)
        elif t == "CIRCLE":
            r["circ"] += 1
        elif t == "INSERT":
            r["blk"] += 1
        elif t in ("TEXT", "MTEXT"):
            r["txt"] += 1
        elif t == "HATCH":
            r["hatch"] += _hatch_area(e)
    rows = []
    for lay in sorted(rep, key=lambda k: -rep[k]["ent"]):
        r = rep[lay]
        rows.append([lay, r["ent"], f"{r['len']/1000:,.1f}",
                     f"{r['area']/1e6:,.1f}", r["circ"], r["blk"], r["txt"],
                     f"{r['hatch']/1e6:,.1f}"])
    print(fmt_table(rows, ["layer", "ent", "len m", "area m2", "circ",
                           "blocks", "text", "hatch m2"]))
    print(f"\n{len(rep)} layers with content.")


def cmd_layer_blocks(args):
    """Block (INSERT) counts per layer, across all layers."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    per = defaultdict(Counter)
    for e in msp:
        if e.dxftype() == "INSERT":
            per[e.dxf.layer or ""][e.dxf.name] += 1
    if not per:
        print("No blocks in drawing.")
        return
    rows = []
    for lay in sorted(per, key=lambda k: -sum(per[k].values())):
        c = per[lay]
        top = ", ".join(f"{n}x{nm}" for nm, n in c.most_common(3))
        rows.append([lay, sum(c.values()), len(c), top])
    print(fmt_table(rows, ["layer", "blocks", "kinds", "top blocks"]))


def cmd_layer_texts(args):
    """Text (TEXT/MTEXT) counts per layer, across all layers."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    c = Counter()
    for e in msp:
        if e.dxftype() in ("TEXT", "MTEXT"):
            c[e.dxf.layer or ""] += 1
    if not c:
        print("No text in drawing.")
        return
    print(fmt_table([[l, n] for l, n in c.most_common()], ["layer", "text count"]))


def cmd_layer_bounds(args):
    """Bounding-box size (m) of the geometry on each layer."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    bnd = {}
    for e in msp:
        lay = e.dxf.layer or ""
        for px, py in _layer_points(e):
            if lay not in bnd:
                bnd[lay] = [px, py, px, py]
            b = bnd[lay]
            b[0], b[1] = min(b[0], px), min(b[1], py)
            b[2], b[3] = max(b[2], px), max(b[3], py)
    if not bnd:
        print("No geometry found.")
        return
    rows = []
    for lay in sorted(bnd):
        b = bnd[lay]
        rows.append([lay, f"{(b[2]-b[0])/1000:,.1f}", f"{(b[3]-b[1])/1000:,.1f}"])
    print(fmt_table(rows, ["layer", "width m", "height m"]))


def cmd_which_layer(args):
    """Find which layers hold a given entity --type (LINE/INSERT/...) or text
    matching --q. Helps locate where things live across many layers."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    if args.type:
        t = args.type.upper()
        c = Counter(e.dxf.layer or "" for e in msp if e.dxftype() == t)
        if not c:
            print(f"No {t} entities found.")
            return
        print(f"Layers containing {t}:")
        print(fmt_table([[l, n] for l, n in c.most_common()], ["layer", "count"]))
    elif args.q:
        q = args.q.lower()
        c = Counter()
        for s, x, y, lay in all_text_entities(msp):
            if q in s.lower():
                c[lay] += 1
        if not c:
            print(f"No text matching '{args.q}'.")
            return
        print(f"Layers with text '{args.q}':")
        print(fmt_table([[l, n] for l, n in c.most_common()], ["layer", "matches"]))
    else:
        print("Pass --type LINE  or  --q keyword.")


def cmd_door_width_check(args):
    """Flag openings narrower than a minimum clear width (hospitals need
    ~1200mm for stretchers, ~900mm accessible). Reads WxH sizes from text."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    pat = re.compile(r"(\d{3,4})\s*[xX]\s*(\d{3,4})")
    sizes = set()
    for s, *_ in all_text_entities(msp):
        for w, h in pat.findall(s):
            sizes.add((int(w), int(h)))
    if not sizes:
        print("No 'W x H' opening sizes found.")
        return
    # Only door-height openings (h >= minheight) -- excludes column sizes,
    # ventilators and other small WxH labels.
    doors = [(w, h) for w, h in sizes if h >= args.minheight]
    mn = args.min
    small = sorted((w, h) for w, h in doors if w < mn)
    print(f"Min clear width target: {mn} mm   (door height >= {args.minheight} mm)")
    print(f"Door-height openings: {len(doors)}")
    if small:
        print(f"\nDoors BELOW {mn} mm wide:")
        for w, h in small:
            print(f"  {w} x {h}")
    else:
        print("All door-height openings meet the target width.")


def cmd_room_count(args):
    """Count rooms/spaces by keyword. --keywords 'CHANGE,LAB,STORE,OT,WARD'
    (comma list); no keywords = common facility spaces."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    texts = [re.sub(r"\s+", " ", s.strip().upper())
             for s, *_ in all_text_entities(msp) if s.strip()]
    if args.keywords:
        kws = [k.strip().upper() for k in args.keywords.split(",") if k.strip()]
    else:
        kws = ["ROOM", "STORE", "AREA", "TOILET", "LAB", "OFFICE", "CHANGE",
               "AIRLOCK", "OT", "ICU", "WARD", "LIFT", "STAIR", "STORAGE",
               "FREEZER", "CONTROL", "GAS"]
    rows = []
    for k in kws:
        rx = re.compile(r"\b" + re.escape(k) + r"\b")
        n = sum(1 for t in texts if rx.search(t))
        if n or args.keywords:
            rows.append([k, n])
    if not rows:
        print("No matching room labels.")
        return
    print(fmt_table(rows, ["keyword", "labels found"]))
    print("\n(Counts text labels containing each word -- a guide, not exact "
          "room count.)")


def cmd_coving_length(args):
    """Total wall-to-floor coving / skirting length = sum of room perimeters
    on a layer (pharma epoxy coving, hospital skirting)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    args.layer = _resolve_layer(msp, args, "area")
    if not args.layer:
        print("No closed shapes found in this drawing.")
        return
    tot, n = 0.0, 0
    for e in msp:
        if (e.dxf.layer or "") != args.layer or e.dxftype() != "LWPOLYLINE" \
                or not e.closed:
            continue
        pts = [(p[0], p[1]) for p in e.get_points()]
        tot += sum(math.dist(pts[i], pts[(i + 1) % len(pts)])
                   for i in range(len(pts)))
        n += 1
    print(f"Layer '{args.layer}': {n} rooms, "
          f"total coving/skirting {tot/1000:,.1f} m")


def cmd_extents(args):
    """Drawing extents / overall sheet size in drawing units and metres."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    xs, ys = [], []
    for e in msp:
        try:
            b = e.bbox() if hasattr(e, "bbox") else None
        except Exception:
            b = None
    # fall back to header extents
    ext = (doc.header.get("$EXTMIN"), doc.header.get("$EXTMAX"))
    if ext[0] and ext[1]:
        x0, y0 = ext[0][0], ext[0][1]
        x1, y1 = ext[1][0], ext[1][1]
        print(f"EXTMIN : {x0:,.1f}, {y0:,.1f}")
        print(f"EXTMAX : {x1:,.1f}, {y1:,.1f}")
        print(f"Size   : {x1-x0:,.1f} x {y1-y0:,.1f} units "
              f"(= {(x1-x0)/1000:,.1f} x {(y1-y0)/1000:,.1f} m if mm)")
    else:
        print("No $EXTMIN/$EXTMAX in header. Open & ZOOM EXTENTS, save, retry.")


def cmd_export(args):
    """Export every modelspace entity to CSV (type, layer, x, y, text)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    out = args.out or (os.path.splitext(args.file)[0] + "_entities.csv")
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["type", "layer", "x", "y", "text"])
        for e in msp:
            t = e.dxftype()
            lay = e.dxf.layer or ""
            x = y = ""
            try:
                p = e.dxf.insert
                x, y = round(p.x, 2), round(p.y, 2)
            except Exception:
                try:
                    p = e.dxf.start
                    x, y = round(p.x, 2), round(p.y, 2)
                except Exception:
                    pass
            txt = (entity_text(e) or "").replace("\n", " ")
            w.writerow([t, lay, x, y, txt])
    print(f"Exported modelspace entities -> {out}")


def cmd_purge(args):
    """Report empty layers and obvious junk (e.g. cracked-CAD watermarks)."""
    doc = load_dxf(args.file)
    msp = doc.modelspace()
    ec = Counter((e.dxf.layer or "") for e in msp)
    empty, junk = [], []
    for l in doc.layers:
        nm = l.dxf.name
        if ec.get(nm, 0) == 0 and nm not in ("0", "Defpoints"):
            empty.append(nm)
        if re.search(r"[一-鿿]", nm):  # CJK = pirated-CAD marker
            junk.append(nm)
    print("EMPTY layers (safe to purge):")
    print("  " + (", ".join(empty) if empty else "(none)"))
    if junk:
        print("\nNON-LATIN / watermark layers (often from cracked AutoCAD):")
        print("  " + ", ".join(junk))


def cmd_dwghelp(args):
    """Print how to convert a .dwg to .dxf so these tools can read it."""
    print(__doc__.split("NOTE on .dwg")[1].strip()
          if "NOTE on .dwg" in __doc__ else "")
    print("""
Steps:
  1. Download ODA File Converter (free):
     https://www.opendesign.com/guestfiles/oda_file_converter
  2. Input folder  = folder with your .dwg
     Output folder  = anywhere
     Output version = ACAD 2018 DXF   (or 2013)
     Output type    = DXF
  3. Press Start. Then run ArchTools on the resulting .dxf.
""")


# ======================================================================
# Excel / BOQ commands
# ======================================================================
def cmd_xls(args):
    """List sheets in an .xls/.xlsx workbook with dimensions."""
    sheets = open_workbook(args.file)
    rows = [[nm, len(rows_), max((len(r) for r in rows_), default=0)]
            for nm, rows_ in sheets]
    print(fmt_table(rows, ["sheet", "rows", "cols"]))


def cmd_xlsdump(args):
    """Dump all non-empty rows of a sheet (pipe-separated)."""
    sheets = open_workbook(args.file)
    target = None
    for nm, rows_ in sheets:
        if args.sheet is None or nm.lower() == args.sheet.lower():
            target = (nm, rows_)
            break
    if target is None:
        print("Sheet not found. Available:",
              ", ".join(nm for nm, _ in sheets))
        return
    nm, rows_ = target
    print(f"=== {nm} ===")
    for i, row in enumerate(rows_):
        cells = [str(x).strip() for x in row]
        if any(cells):
            print(i, "|".join(cells))


def cmd_find(args):
    """Search a BOQ/estimate for items matching a keyword; show qty/rate/amt.

    Heuristic columns: assumes [Sr, Description, Qty, Unit, Rate, Amount].
    Use --cols to override, e.g. --cols 1,2,3,4,5  (desc,qty,unit,rate,amt).
    """
    sheets = open_workbook(args.file)
    kw = args.keyword.lower()
    if args.cols:
        dcol, qcol, ucol, rcol, acol = (int(x) for x in args.cols.split(","))
    else:
        dcol, qcol, ucol, rcol, acol = 1, 2, 3, 4, 5
    hits = []
    for nm, rows_ in sheets:
        for i, row in enumerate(rows_):
            if dcol >= len(row):
                continue
            desc = str(row[dcol])
            if kw in desc.lower():
                def g(c):
                    return row[c] if c < len(row) else ""
                hits.append([nm, i, str(g(qcol)), str(g(ucol)),
                             str(g(rcol)), str(g(acol)), desc[:60]])
    if not hits:
        print(f"No items matching '{args.keyword}'.")
        return
    print(fmt_table(hits, ["sheet", "row", "qty", "unit", "rate", "amount", "description"]))


def cmd_compare(args):
    """Compare quantities for a keyword across two BOQ/estimate files."""
    kw = args.keyword.lower()

    def grab(path):
        sheets = open_workbook(path)
        items = []
        for nm, rows_ in sheets:
            for i, row in enumerate(rows_):
                if len(row) > 1 and kw in str(row[1]).lower():
                    q = row[2] if len(row) > 2 else ""
                    try:
                        q = float(q)
                    except (TypeError, ValueError):
                        q = None
                    items.append((i, q, str(row[1])[:50]))
        return items

    a, b = grab(args.file_a), grab(args.file_b)
    ta = sum(q for _, q, _ in a if isinstance(q, float))
    tb = sum(q for _, q, _ in b if isinstance(q, float))
    print(f"File A: {args.file_a}")
    for i, q, d in a:
        print(f"  row{i}: {q}  {d}")
    print(f"  TOTAL A = {ta:.3f}")
    print(f"\nFile B: {args.file_b}")
    for i, q, d in b:
        print(f"  row{i}: {q}  {d}")
    print(f"  TOTAL B = {tb:.3f}")
    print(f"\nDIFFERENCE (A - B) = {ta - tb:+.3f}", end="")
    if tb:
        print(f"   ({100*(ta-tb)/tb:+.1f}%)")
    else:
        print()


# ======================================================================
# Estimating calculators (no drawing needed)
# ======================================================================
def cmd_calc_concrete(args):
    """Concrete volume:  L x B x H x N  (metres)."""
    v = args.l * args.b * args.h * args.n
    print(f"Concrete volume = {args.l} x {args.b} x {args.h} x {args.n} "
          f"= {v:.3f} m3")
    # Nominal 1:2:4, dry-volume factor 1.54 (textbook):
    #   cement 1/7*1.54 = 0.22 m3 = 6.34 bags ; sand 0.44 ; aggregate 0.88 per m3
    print(f"Cement (1:2:4, 6.34 bag/m3) ~ {v*6.34:.1f} bags")
    print(f"Sand   (0.44 m3/m3)         ~ {v*0.44:.2f} m3")
    print(f"Aggr.  (0.88 m3/m3)         ~ {v*0.88:.2f} m3")
    print("Formula: V = L x B x H x N   (mix 1:2:4 nominal)")


def cmd_calc_rebar(args):
    """Reinforcement weight from bar diameter, length, count.
    Weight per metre = d^2 / 162  (kg/m, d in mm)."""
    wpm = args.dia ** 2 / 162.0
    total = wpm * args.len * args.n
    print(f"Bar dia {args.dia} mm -> {wpm:.3f} kg/m")
    print(f"{args.n} bars x {args.len} m = {args.n*args.len:.1f} running m")
    print(f"Total weight = {total:.2f} kg  ({total/1000:.3f} tonne)")


def cmd_calc_brick(args):
    """Brick / block count for a wall area & thickness (standard 230x110x75
    brick with 10 mm joint ~ 500 bricks per m3)."""
    vol = args.area * (args.thk / 1000.0)
    bricks = vol * 500
    print(f"Wall area {args.area} m2 x {args.thk} mm = {vol:.3f} m3")
    print(f"Bricks (~500/m3)  ~ {bricks:,.0f} nos")
    print(f"Mortar (~0.3 m3/m3) ~ {vol*0.3:.3f} m3")


def cmd_calc_plaster(args):
    """Plaster mortar volume for an area & thickness."""
    vol = args.area * (args.thk / 1000.0)
    print(f"Plaster area {args.area} m2 x {args.thk} mm = {vol:.3f} m3 mortar")
    print(f"Cement (1:4, ~5.5 bag/m3) ~ {vol*5.5:.1f} bags")
    print(f"Sand   ~ {vol*1.1:.2f} m3")


def cmd_calc_paint(args):
    """Paint quantity for an area & number of coats (~0.1 L/m2/coat)."""
    litres = args.area * args.coats * 0.10
    print(f"Paint area {args.area} m2 x {args.coats} coats")
    print(f"Paint (~0.10 L/m2/coat) ~ {litres:.1f} litres")


# Every calculation's formula + basis, in one place (transparency/audit).
FORMULAS = [
    ("calc-concrete", "V = L x B x H x N",
     "1:2:4 nominal, dry factor 1.54 -> 6.34 cement bags, 0.44 sand, 0.88 aggregate per m3"),
    ("calc-rebar", "weight = d^2 / 162  (kg/m, d in mm); total = weight x len x nos",
     "d^2/162 is the standard bar-weight rule (steel 7850 kg/m3)"),
    ("calc-steel", "kg = concrete_vol x steel% x 7850",
     "steel density 7850 kg/m3; steel% typically 0.8-1.5% slabs, 1-3% columns"),
    ("calc-brick", "bricks = wall_vol x 500;  mortar = wall_vol x 0.30",
     "brick 230x110x75 + 10mm joint ~ 500/m3; ~30% mortar"),
    ("calc-plaster", "mortar = area x thk;  cement = mortar x 5.5 bags; sand x 1.1",
     "1:4 mix, 12mm typical"),
    ("calc-paint", "litres = area x coats x 0.10",
     "coverage ~10 m2 per litre per coat"),
    ("calc-tiles", "tiles = ceil(area / tile^2 x (1 + waste%))",
     "tile size in mm (square); waste default 10%"),
    ("calc-excavation", "V = L x B x D;  to cart = V x (1 + bulking%)",
     "soil bulking ~25%"),
    ("calc-stair", "risers = round(H/riser); treads = risers-1; going = treads x tread",
     "comfort: 2R + T ~ 600mm; riser <=150-190, tread >=250-300"),
    ("calc-watertank", "litres = L x B x H x 1000",
     "L,B,H in metres"),
    ("calc-cost", "amount = qty x rate", "—"),
    ("convert", "m2->ft2 x10.7639; m->ft x3.28084; m3->ft3 x35.3147", "exact factors"),
    ("calc-ach", "supply CMH = area x height x ACH;  CFM = CMH x 0.5886",
     "ACH per cleanroom grade -- see cleanroom-ref"),
    ("calc-water-demand", "daily L = persons x LPCD",
     "LPCD: industrial ~45, office ~45, hospital ~340/bed (local code varies)"),
    ("calc-hvac", "TR ~ area / (m2 per TR)",
     "rule of thumb (offices ~10, labs ~7, clean/server ~5); add equipment+fresh air"),
    ("calc-fsi", "FSI/FAR = built-up / plot area",
     "limit set by local development control rules"),
    ("calc-coverage", "coverage% = footprint / plot x 100",
     "limit set by local rules"),
    ("calc-occupancy", "occupants = ceil(area / load-factor)",
     "load-factor m2/person: factory ~10, office ~10, lab ~9, ward ~7.5, assembly ~1.5 (NBC)"),
    ("calc-parking", "ECS = ceil(built-up / norm)",
     "norm m2 per car varies by city/use"),
    ("calc-ramp", "run = rise x N   (slope 1:N)",
     "1:12 = accessible; steeper not wheelchair-friendly"),
    ("calc-exit-width", "width(mm) = occupants x mm/person",
     "NBC ~5 mm/person stairs, ~3.3 doors/ramps; min single exit 1000-1500mm"),
    ("calc-electrical", "connected = area x VA/m2;  demand = connected x demand%",
     "density VA/m2: office ~60, lab ~80-120, industrial ~100-200 (excl. process)"),
    ("calc-lighting", "fixtures = ceil(lux x area / (lumen x UF x MF))",
     "lumen method; UF utilisation ~0.5, MF maintenance ~0.8"),
    ("calc-rainwater", "Q(L/s) = roof_area x intensity / 3600",
     "intensity mm/hr (India ~150); pipe by flow"),
    ("calc-sewage", "sewage = water x return-factor (~0.8)",
     "STP rounded up ~20%"),
    ("calc-fire-area", "check: compartment_area <= max",
     "max by occupancy/hazard (industrial ~750-2000 m2)"),
    ("calc-travel", "check: travel_distance <= max",
     "NBC ~22.5 m unsprinklered, ~30 m sprinklered"),
    ("colvol", "Sum over storeys: cols x section_area x floor-to-floor height",
     "levels read from the drawing; interior cols stop at terrace slab"),
    ("plinth-colvol", "vol = GF columns x section_area x plinth height",
     "plinth height = GROUND - ROAD from the drawing (or --plinth); fills colvol's gap"),
    ("plinth-beams", "vol = beam_run x width x depth;  steel = vol x 7850 x steel%",
     "beam run measured on --layer; default section 230x450 mm, steel 1.5%"),
    ("calc-plinth-beam", "vol = L x B x D x N;  steel = vol x 7850 x steel%",
     "default section 0.23 x 0.45 m, steel 1.5%"),
    ("calc-plinth-fill", "compacted = area x depth;  loose = compacted x (1+compact%)",
     "earth/murrum fill up to plinth; compaction adds ~20%"),
    ("calc-plinth-masonry", "vol = perimeter x thk x height;  bricks = vol x 500",
     "plinth wall above footing; 230 mm default, ~500 brick/m3, 30% mortar"),
    ("calc-dpc", "area = perimeter x width;  conc = area x thk",
     "DPC 1:1.5:3, ~40 mm thick, band = wall thickness (~7.8 bag/m3)"),
    ("calc-anti-termite", "area = plinth_area + perimeter x trench_depth;  chem = area x rate",
     "IS 6313 pre-construction; emulsion ~5 L/m2"),
    ("calc-plinth-protection", "area = perimeter x width;  PCC = area x thk",
     "apron ~0.9 m wide, 75 mm PCC 1:4:8 (~3.4 bag/m3)"),
    ("calc-pcc-bed", "vol = area x thk",
     "leveling course below plinth/footing, 1:4:8 (~3.4 bag/m3)"),
    ("chajja", "vol = run x projection x avg_thk;  steel = vol x 7850 x steel%",
     "run measured on --layer; tapered section root->tip, avg thickness; default 150->75 mm"),
    ("calc-chajja", "vol = projection x length x avg_thk x N;  avg_thk = (root+tip)/2",
     "tapered RCC sunshade, default root 150 / tip 75 mm, steel ~1%"),
    ("calc-chajja-shutter", "area = (soffit + front_edge + 2 ends) x N",
     "formwork contact area; back against wall & top finished are excluded"),
    ("calc-chajja-plaster", "area = (2 x proj x length + drip) x N;  mortar = area x thk",
     "top + bottom faces + front drip; 1:4 plaster, 12 mm default"),
    ("foundation", "per size: conc = WxLxthk x n; exc = (W+2ws)(L+2ws)xD x n; "
     "PCC = (W+.2)(L+.2)x0.1 x n; backfill = exc - conc - PCC",
     "footing sizes read from the drawing's F-tags; thk 450mm, depth 1.5m, "
     "working space 300mm, steel 0.8% -- site defaults, overridable"),
    ("concrete", "auto total = column storey concrete + slab(area x thk x floors)",
     "columns from the grid markers + levels; slab from footprint; one click, nothing typed"),
    ("slab", "auto: conc = area x thk x floors;  steel = area x floors x kg/m2",
     "area = footprint or --layer; floors counted from the drawing's level marks; thk 150mm, steel 10 kg/m2 default"),
    ("calc-slab", "conc = area x thk x floors;  steel = area x floors x kg/m2",
     "manual slab; default 150 mm, 10 kg/m2"),
    ("calc-slab-steel", "area: area x floors x kg/m2;  pct: vol x steel% x 7850",
     "kg/m2 ~8-12 for slabs; %-of-concrete ~0.8-1%"),
    ("calc-slab-shutter", "area = (plan_area + perimeter x thk) x floors",
     "soffit + edge band; formwork contact area"),
    ("calc-slab-plaster", "area = plan_area x floors;  mortar = area x thk",
     "ceiling plaster, 1:4, 12 mm default"),
]


def cmd_formulas(args):
    """Show the formula and assumptions behind every calculation -- so an
    architect can verify exactly what each number is based on."""
    flt = (args.q or "").lower()
    shown = 0
    for tool, formula, basis in FORMULAS:
        if flt and flt not in (tool + " " + formula + " " + basis).lower():
            continue
        shown += 1
        print(f"{tool}")
        print(f"    formula : {formula}")
        if basis and basis != "—":
            print(f"    basis   : {basis}")
        print()
    if not shown:
        print(f"No formula matching '{args.q}'.")
    else:
        print("Norms (FSI, parking, occupancy load, LPCD) are jurisdiction-"
              "specific -- adjust to your local code.")


def cmd_calc_tiles(args):
    """Tiles needed for a floor area. --area m2, --tile mm (square), --waste %."""
    per = (args.tile / 1000.0) ** 2
    tiles = math.ceil(args.area / per * (1 + args.waste / 100.0))
    print(f"Floor {args.area} m2, {args.tile}mm tiles ({per:.3f} m2 each), "
          f"+{args.waste}% wastage")
    print(f"Tiles needed: {tiles} nos")


def cmd_calc_excavation(args):
    """Excavation volume L×B×D + carting (bulking %)."""
    v = args.l * args.b * args.d
    print(f"Excavation {args.l} x {args.b} x {args.d} = {v:.2f} m3")
    print(f"With {args.bulk}% bulking to cart: {v*(1+args.bulk/100.0):.2f} m3")


def cmd_calc_steel(args):
    """Reinforcement weight from concrete volume × steel percentage."""
    kg = args.vol * 7850 * args.pct / 100.0
    print(f"Concrete {args.vol} m3 x {args.pct}% steel "
          f"= {kg:,.1f} kg ({kg/1000:.3f} tonne)")


def cmd_calc_stair(args):
    """Staircase steps from floor height, target riser & tread (mm)."""
    h = args.height * 1000.0
    risers = max(1, round(h / args.riser))
    actual = h / risers
    treads = risers - 1
    going = treads * args.tread
    print(f"Floor height {args.height} m | target riser {args.riser}mm, "
          f"tread {args.tread}mm")
    print(f"Risers: {risers} @ {actual:.1f}mm   Treads: {treads}   "
          f"Going: {going/1000:.2f} m")


def cmd_calc_watertank(args):
    """Water tank capacity from L×B×H (m) -> litres."""
    v = args.l * args.b * args.h
    print(f"Tank {args.l} x {args.b} x {args.h} m = {v:.3f} m3 "
          f"= {v*1000:,.0f} litres")


def cmd_calc_cost(args):
    """Amount = quantity × rate."""
    print(f"{args.qty} x {args.rate} = {args.qty*args.rate:,.2f}")


def cmd_convert(args):
    """Unit conversion. --value N --unit  m2-ft2 | ft2-m2 | m-ft | ft-m |
    m3-ft3 | ft3-m3."""
    v, u = args.value, args.unit.lower()
    table = {
        "m2-ft2": (v * 10.7639, "ft2"), "ft2-m2": (v / 10.7639, "m2"),
        "m-ft": (v * 3.28084, "ft"), "ft-m": (v / 3.28084, "m"),
        "m3-ft3": (v * 35.3147, "ft3"), "ft3-m3": (v / 35.3147, "m3"),
    }
    if u not in table:
        print("units: " + ", ".join(table))
        return
    out, lbl = table[u]
    print(f"{v} {u.split('-')[0]} = {out:,.3f} {lbl}")


# ----- Sector calculators: cleanroom / HVAC / code compliance ----------
def cmd_calc_ach(args):
    """HVAC supply air from air changes per hour. --area m2, --height m,
    --ach (changes/hr). Cleanroom ACH guide: GMP-A/B ~ 60+, C ~ 30-40,
    D ~ 15-20; OT ~ 20-25; general lab ~ 6-12."""
    vol = args.area * args.height
    cmh = vol * args.ach
    print(f"Room {args.area} m2 x {args.height} m = {vol:.1f} m3")
    print(f"At {args.ach} ACH -> {cmh:,.0f} m3/h supply air "
          f"({cmh*0.5886:,.0f} CFM)")


def cmd_calc_water_demand(args):
    """Daily water demand. --persons, --lpcd (litres/person/day; industrial
    ~45, hospital ~340/bed, office ~45). Suggests a 1-day storage tank."""
    total = args.persons * args.lpcd
    print(f"{args.persons} persons x {args.lpcd} L/day = {total:,.0f} L/day")
    print(f"1-day storage tank: {total/1000:,.1f} m3 "
          f"(e.g. split 50/50 OH+UG)")


def cmd_calc_hvac(args):
    """Rough cooling load (TR). --area m2, --persqm (m2 per TR; offices ~10,
    labs ~7, server/clean ~5). Quick sizing only."""
    tr = args.area / args.persqm
    print(f"Area {args.area} m2 / {args.persqm} m2 per TR "
          f"= {tr:,.1f} TR (rough). Add equipment/fresh-air load.")


def cmd_calc_fsi(args):
    """FSI / FAR consumed = built-up / plot area (same units)."""
    fsi = args.builtup / args.plot
    print(f"Built-up {args.builtup} / Plot {args.plot} = FSI/FAR {fsi:.2f}")


def cmd_calc_coverage(args):
    """Ground coverage % = footprint / plot area x 100."""
    pct = args.footprint / args.plot * 100
    print(f"Footprint {args.footprint} / Plot {args.plot} "
          f"= {pct:.1f}% ground coverage")


def cmd_calc_occupancy(args):
    """Occupant load = area / load-factor (m2 per person). Factory ~10,
    office ~10, lab ~9, ward ~7.5, assembly ~1.5. Flags min 2 exits."""
    occ = math.ceil(args.area / args.factor)
    print(f"Area {args.area} m2 / {args.factor} m2 per person = {occ} occupants")
    print(f"Minimum exits (code: >= 2): "
          f"{2 if occ <= 500 else 3 if occ <= 1000 else 4}")


def cmd_calc_parking(args):
    """Parking (ECS) = built-up / norm (m2 per car; varies by city, ~100)."""
    ecs = math.ceil(args.builtup / args.per)
    print(f"Built-up {args.builtup} / {args.per} m2 per car = {ecs} car spaces (ECS)")


def cmd_calc_ramp(args):
    """Ramp run length for a height at slope 1:N (accessibility = 1:12)."""
    length = args.height * args.slope
    print(f"Rise {args.height} m at 1:{args.slope} -> ramp run {length:.2f} m")
    if args.slope < 12:
        print("  (steeper than 1:12 -- not wheelchair-accessible)")


def cmd_calc_exit_width(args):
    """Required egress width = occupants x mm/person (stairs ~5, doors/ramps
    ~3.3 per NBC). Reports total width to provide across exits."""
    w = args.occupants * args.permm
    print(f"{args.occupants} occupants x {args.permm} mm/person "
          f"= {w:,.0f} mm total egress width")
    print(f"= {w/1000:.2f} m to provide across all exits "
          f"(min single exit 1000-1500 mm)")


def cmd_cleanroom_ref(args):
    """Reference: cleanroom grades vs ISO class and typical air changes."""
    rows = [
        ["GMP A (ISO 5)", "ISO 5", "laminar / unidirectional"],
        ["GMP B (ISO 5/7)", "ISO 7 at rest", "~60+ ACH"],
        ["GMP C (ISO 7/8)", "ISO 8 at rest", "~30-40 ACH"],
        ["GMP D (ISO 8)", "ISO 8", "~15-20 ACH"],
        ["Operating theatre", "ISO 5-7", "~20-25 ACH"],
        ["General lab", "-", "~6-12 ACH"],
        ["Ward / patient", "-", "~6 ACH"],
    ]
    print(fmt_table(rows, ["grade", "ISO class", "typical air changes"]))
    print("\nUse 'calc-ach' with the room's area/height to get supply CMH.")


def cmd_calc_electrical(args):
    """Electrical load. --area m2, --va (VA/m2 load density), --demand %.
    Density guide: office ~60, lab ~80-120, industrial ~100-200 (excl. process)."""
    conn = args.area * args.va / 1000.0
    dem = conn * args.demand / 100.0
    print(f"Area {args.area} m2 x {args.va} VA/m2 = {conn:,.1f} kVA connected")
    print(f"Demand load ({args.demand}%): {dem:,.1f} kVA")
    print(f"Suggested supply/DG: >= {math.ceil(dem*1.1)} kVA")


def cmd_calc_lighting(args):
    """Light fixtures for a target lux (lumen method). --area, --lux,
    --lumen/fixture, --uf utilisation, --mf maintenance."""
    fx = math.ceil(args.lux * args.area / (args.lumen * args.uf * args.mf))
    print(f"Room {args.area} m2, target {args.lux} lux")
    print(f"Fixtures: {fx} nos @ {args.lumen} lm "
          f"(UF {args.uf}, MF {args.mf})")
    print("Lux guide: office 300-500, lab 500-750, OT 1000, ward 100-300, "
          "warehouse 150-200")


def cmd_calc_rainwater(args):
    """Rainwater runoff from a roof. --area m2, --intensity mm/hr (India ~150).
    Gives flow + downpipe size guide."""
    q = args.area * args.intensity / 3600.0
    pipe = ("100 mm" if q <= 11 else "150 mm" if q <= 33
            else "200 mm" if q <= 70 else "250 mm+")
    print(f"Roof {args.area} m2 x {args.intensity} mm/hr = {q:,.1f} L/s runoff")
    print(f"Downpipe guide: {pipe} (one); split large roofs into several")


def cmd_calc_sewage(args):
    """Sewage / STP load = water demand x return factor. --water L/day,
    --factor (~0.8)."""
    s = args.water * args.factor
    print(f"Water {args.water:,.0f} L/day x {args.factor} = {s:,.0f} L/day sewage")
    print(f"STP capacity (round up ~20%): {math.ceil(s*1.2/1000)*1000:,.0f} L/day")


def cmd_calc_fire_area(args):
    """Check a fire compartment against its max area. --area, --limit m2
    (industrial ~750-2000 by hazard; malls/assembly differ)."""
    ok = args.area <= args.limit
    print(f"Compartment {args.area} m2 vs max {args.limit} m2: "
          + ("OK" if ok else "EXCEEDS -- add fire wall / sprinklers"))


def cmd_calc_travel(args):
    """Check travel distance to exit. --distance m, --max m (NBC: ~22.5 un-
    sprinklered, ~30 sprinklered; factory varies)."""
    ok = args.distance <= args.max
    print(f"Travel distance {args.distance} m vs max {args.max} m: "
          + ("OK" if ok else "EXCEEDS -- add an exit / reduce dead-end"))


# ======================================================================
# Project & tool-registry commands
# ======================================================================
def cmd_project(args):
    """Manage the plot/building manifest (project.json).

    Sub-actions:
      init   --plot "name"          create/initialise the project
      add    NAME FILE  [--note ..] register a building's drawing
      remove NAME                   drop a building
      list                          show all buildings
    """
    proj = load_project()
    act = args.action

    if act == "init":
        proj["plot"] = args.plot or proj.get("plot", "")
        proj.setdefault("buildings", {})
        save_project(proj)
        print(f"Project ready (plot: {proj['plot'] or '-'}) -> {PROJECT_FILE}")

    elif act == "add":
        if not args.name or not args.file:
            sys.exit("Usage: project add NAME FILE.dxf [--note ...]")
        if not os.path.isfile(args.file):
            sys.exit(f"ERROR: file not found: {args.file}")
        proj.setdefault("buildings", {})[args.name] = {
            "dxf": args.file, "note": args.note or ""}
        save_project(proj)
        print(f"Added building '{args.name}' -> {args.file}")

    elif act == "box":
        b = proj.get("buildings", {}).get(args.name)
        if b is None:
            sys.exit(f"No building '{args.name}'.")
        b["box"] = args.file  # the value reuses the file positional
        save_project(proj)
        print(f"Set box for '{args.name}': {args.file}")

    elif act == "remove":
        if proj.get("buildings", {}).pop(args.name, None) is None:
            sys.exit(f"No building '{args.name}'.")
        save_project(proj)
        print(f"Removed building '{args.name}'.")

    elif act == "list":
        b = proj.get("buildings", {})
        if not b:
            print("No buildings yet. Add one:  project add M1 file.dxf")
            return
        print(f"Plot: {proj.get('plot') or '-'}")
        rows = [[nm, e.get("dxf", ""), e.get("note", "")]
                for nm, e in b.items()]
        print(fmt_table(rows, ["building", "drawing", "note"]))
    else:
        sys.exit("Unknown project action. Use: init | add | remove | list")


def cmd_tools(args):
    """Print every available command as a machine-readable registry.

    The plain-English front-end (ask.py / a web backend) feeds this list to
    the model so it knows exactly which tools exist and what each one does.
    """
    import json
    parser = build_parser()
    # walk the subparsers to extract name, help, and options
    registry = []
    subactions = [a for a in parser._actions
                  if isinstance(a, argparse._SubParsersAction)]
    if subactions:
        choices = subactions[0].choices
        # help strings live on the _ChoicesPseudoAction list
        helps = {ca.dest: (ca.help or "")
                 for ca in subactions[0]._choices_actions}
        for name, sp in choices.items():
            opts = []
            for a in sp._actions:
                if isinstance(a, argparse._HelpAction):
                    continue
                flag = a.option_strings[0] if a.option_strings else a.dest
                opts.append({"name": flag,
                             "required": getattr(a, "required", False),
                             "help": a.help or ""})
            registry.append({"command": name,
                             "summary": helps.get(name, ""),
                             "args": opts})
    if args.json:
        print(json.dumps(registry, indent=2))
    else:
        rows = [[r["command"],
                 " ".join(o["name"] for o in r["args"]),
                 r["summary"]] for r in registry]
        print(fmt_table(rows, ["command", "args", "what it does"]))


def _arg_type(a):
    if isinstance(a, argparse._StoreTrueAction):
        return "flag"
    if a.type is int:
        return "int"
    if a.type is float:
        return "float"
    return "str"


def build_catalog():
    """Return the catalog: one dict per user-facing command, with category,
    title, and form-renderable parameters. Powers the web search UI."""
    parser = build_parser()
    sub = next(a for a in parser._actions
               if isinstance(a, argparse._SubParsersAction))
    helps = {ca.dest: (ca.help or "") for ca in sub._choices_actions}
    out = []
    for name, sp in sub.choices.items():
        if name in CATALOG_HIDE:
            continue
        params, needs_building, needs_path = [], False, False
        for a in sp._actions:
            if isinstance(a, argparse._HelpAction):
                continue
            if a.dest == "building":
                needs_building = True
                continue
            if a.dest == "all":
                continue
            if a.dest in ("file", "file_a", "file_b") and not a.option_strings:
                needs_path = True
                continue
            params.append({
                "name": a.option_strings[0] if a.option_strings else a.dest,
                "flag": a.option_strings[0] if a.option_strings else None,
                "type": _arg_type(a),
                "required": bool(getattr(a, "required", False)
                                 or (not a.option_strings and a.nargs is None)),
                "default": a.default if a.default is not None else "",
                "help": a.help or "",
            })
        cat, title = CATALOG.get(name, ("Other", helps.get(name, name)))
        out.append({
            "command": name, "category": cat, "title": title,
            "summary": helps.get(name, ""),
            "needs_building": needs_building, "needs_path": needs_path,
            "params": params,
        })
    out.sort(key=lambda t: (t["category"], t["title"]))
    return out


def auto_metrics(path):
    """Numbers that can be read straight from a drawing, used to auto-fill
    calculator forms so the user just presses Run. Geometry only -- spec
    values (thickness, mix, rates, persons) are NOT in a DXF and stay as
    defaults. Returns {} if the file can't be read."""
    out = {}
    try:
        doc = load_dxf(path)
        msp = doc.modelspace()
    except SystemExit:
        return out
    except Exception:
        return out
    res = _largest_closed(doc, msp, None)
    if res:
        area, per = res
        out["area"] = round(area, 2)
        out["footprint"] = round(area, 2)
        out["builtup"] = round(area, 2)
        out["perimeter"] = round(per, 1)
    lv = extract_levels(msp)
    seq = [f for f in FLOOR_ORDER if f in lv]
    if seq:
        out["floors"] = len(seq)
        if len(seq) > 1:
            out["floor_height"] = round(lv[seq[1]] - lv[seq[0]], 2)
    if "GROUND" in lv and "ROAD" in lv:
        out["plinth_height"] = round(lv["GROUND"] - lv["ROAD"], 2)
    if lv:
        base = lv.get("ROAD", lv.get("GROUND"))
        if base is not None:
            out["building_height"] = round(max(lv.values()) - base, 2)
    if "area" in out and "floors" in out:
        # occupant estimate: NBC-style ~10 m2/person on total built-up
        out["occupants"] = int(out["area"] * out["floors"] / 10)
    return out


def cmd_catalog(args):
    """Print the full tool catalog as JSON (used by the web UI)."""
    import json
    print(json.dumps(build_catalog(), indent=2))


# ======================================================================
# Plugin system  --  auto-discovered tools added at runtime
# ======================================================================
# New capabilities the AI invents for papa are saved as small files in the
# plugins/ folder, NOT by editing this file.  That way a bad auto-write can
# only break its own plugin, never the core toolkit.  Each plugin file looks
# like:
#
#     META = {"command": "centroid",
#             "summary": "centre point of a layer's geometry",
#             "args": [{"name": "--layer", "type": "str", "required": True,
#                       "help": "layer name"}]}
#
#     def run(args, ctx):
#         doc = ctx.load_dxf(args.file)        # ctx exposes the core helpers
#         ...
#         print(...)
#
# Commands accept --building / --all automatically, same as built-ins.
import types as _types

PLUGIN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plugins")

# Helpers handed to every plugin via the `ctx` argument.
CTX = _types.SimpleNamespace(
    load_dxf=load_dxf, open_workbook=open_workbook, fmt_table=fmt_table,
    all_text_entities=all_text_entities, entity_text=entity_text,
    parse_levels=parse_levels, Counter=Counter, defaultdict=defaultdict,
    ezdxf=ezdxf, re=re, math=math,
)


def load_plugins():
    """Discover plugins/*.py; return [(meta, module), ...]. Skips broken ones."""
    out = []
    if not os.path.isdir(PLUGIN_DIR):
        return out
    import importlib.util
    for fn in sorted(os.listdir(PLUGIN_DIR)):
        if not fn.endswith(".py") or fn.startswith("_"):
            continue
        path = os.path.join(PLUGIN_DIR, fn)
        try:
            spec = importlib.util.spec_from_file_location("plugin_" + fn[:-3], path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            if hasattr(mod, "META") and hasattr(mod, "run"):
                out.append((mod.META, mod))
        except Exception as e:  # a broken plugin must not sink the whole CLI
            sys.stderr.write(f"[plugin skipped] {fn}: {e}\n")
    return out


# ======================================================================
# CLI wiring
# ======================================================================
def build_parser():
    p = argparse.ArgumentParser(
        prog="ArchTools.py",
        description="Architect's DXF drawing & BOQ/estimate toolkit.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Run 'python ArchTools.py <command> -h' for command help.")
    sub = p.add_subparsers(dest="cmd", required=True)

    def add(name, func, help_):
        sp = sub.add_parser(name, help=help_)
        sp.set_defaults(func=func)
        return sp

    def file_arg(sp):
        """Give a DXF command an optional file path + --building / --all."""
        sp.add_argument("file", nargs="?",
                        help="path to .dxf (or use --building)")
        sp.add_argument("--building",
                        help="building name from project.json")
        sp.add_argument("--all", action="store_true",
                        help="run on every building in the plot")

    # --- DXF ---
    for name, func, h in [
        ("info", cmd_info, "drawing summary (version/layers/entities/extents)"),
        ("layers", cmd_layers, "list layers + entity counts"),
        ("levels", cmd_levels, "extract floor/level marks"),
        ("columns", cmd_columns, "count columns by size variant"),
        ("areas", cmd_areas, "closed-polyline areas per layer"),
        ("lengths", cmd_lengths, "line/polyline length per layer"),
        ("blocks", cmd_blocks, "block (INSERT) counts"),
        ("schedule", cmd_schedule, "door/window/opening schedule"),
        ("hatch", cmd_hatch, "hatch (finish) area per layer"),
        ("drawing-check", cmd_drawing_check, "which conventions/tools work on this drawing"),
        ("extents", cmd_extents, "drawing extents / sheet size"),
    ]:
        sp = add(name, func, h)
        file_arg(sp)

    sp = add("scan", cmd_scan, "one-tap geometry scan (no tags/layers needed)")
    file_arg(sp)
    sp.add_argument("--gap", type=float, help="sheet separation gap (m, default 12)")
    sp.add_argument("--mincols", type=int, help="min columns to count a sheet (default 8)")

    for name, func, h in [
        ("purge", cmd_purge, "report empty/junk layers"),
    ]:
        sp = add(name, func, h)
        file_arg(sp)

    sp = add("text", cmd_text, "dump unique text to a file")
    file_arg(sp); sp.add_argument("--out", help="output .txt path")

    sp = add("dims", cmd_dims, "dimension measurements")
    file_arg(sp); sp.add_argument("--list", action="store_true",
                                  help="print every value")

    sp = add("textheights", cmd_textheights, "histogram of text heights")
    file_arg(sp); sp.add_argument("--top", type=int, default=15,
                                  help="show top N heights")

    sp = add("centroid", cmd_centroid, "centre point of a layer's geometry")
    file_arg(sp); sp.add_argument("--layer", help="layer name (auto if omitted)")

    sp = add("circles", cmd_circles, "count circles + radius stats")
    file_arg(sp)

    sp = add("perimeter-columns", cmd_perimeter_columns,
             "perimeter vs interior columns")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 (auto from building)")
    sp.add_argument("--floor", help="floor name (default GROUND)")

    sp = add("plinth-columns", cmd_plinth_columns, "plinth (ground-floor) columns")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 (auto from building)")

    # --- Plinth suite ---
    sp = add("plinth-colvol", cmd_plinth_colvol, "plinth column concrete volume")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 (auto from building)")
    sp.add_argument("--levels", help="FLOOR:elev,... override (metres)")
    sp.add_argument("--plinth", type=float, help="plinth height (m) if no level marks")

    sp = add("plinth-beams", cmd_plinth_beams, "plinth-beam concrete + steel")
    file_arg(sp); sp.add_argument("--layer", help="plinth-beam layer name")
    sp.add_argument("--width", type=float, default=230, help="beam width (mm)")
    sp.add_argument("--depth", type=float, default=450, help="beam depth (mm)")
    sp.add_argument("--steel", type=float, default=1.5, help="steel percent")

    sp = add("plinth-area", cmd_plinth_area, "plinth area + perimeter")
    file_arg(sp); sp.add_argument("--layer", help="boundary layer (cleaner)")

    sp = add("calc-plinth-beam", cmd_calc_plinth_beam, "plinth beam concrete+steel L*B*D*N")
    sp.add_argument("--l", type=float, required=True)
    sp.add_argument("--b", type=float, default=0.23)
    sp.add_argument("--d", type=float, default=0.45)
    sp.add_argument("--n", type=float, default=1)
    sp.add_argument("--steel", type=float, default=1.5)

    sp = add("calc-plinth-fill", cmd_calc_plinth_fill, "plinth earth filling volume")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--depth", type=float, required=True)
    sp.add_argument("--compact", type=float, default=20)

    sp = add("calc-plinth-masonry", cmd_calc_plinth_masonry, "plinth wall masonry + bricks")
    sp.add_argument("--perimeter", type=float, required=True)
    sp.add_argument("--thk", type=float, default=230)
    sp.add_argument("--height", type=float, required=True)

    sp = add("calc-dpc", cmd_calc_dpc, "damp-proof course area + concrete")
    sp.add_argument("--perimeter", type=float, required=True)
    sp.add_argument("--width", type=float, default=230)
    sp.add_argument("--thk", type=float, default=40)

    sp = add("calc-anti-termite", cmd_calc_anti_termite, "anti-termite treatment area")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--perimeter", type=float, required=True)
    sp.add_argument("--depth", type=float, default=0.3)
    sp.add_argument("--rate", type=float, default=5)

    sp = add("calc-plinth-protection", cmd_calc_plinth_protection, "plinth protection apron PCC")
    sp.add_argument("--perimeter", type=float, required=True)
    sp.add_argument("--width", type=float, default=0.9)
    sp.add_argument("--thk", type=float, default=75)

    sp = add("calc-pcc-bed", cmd_calc_pcc_bed, "PCC leveling bed volume")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--thk", type=float, default=100)

    # --- Chajja / sunshade suite ---
    sp = add("chajja", cmd_chajja, "chajja/sunshade concrete + steel (from layer)")
    file_arg(sp); sp.add_argument("--layer", help="chajja/sunshade layer name")
    sp.add_argument("--proj", type=float, default=0.6, help="projection (m)")
    sp.add_argument("--root", type=float, default=150, help="root thickness (mm)")
    sp.add_argument("--tip", type=float, default=75, help="tip thickness (mm)")
    sp.add_argument("--steel", type=float, default=1.0, help="steel percent")

    sp = add("calc-chajja", cmd_calc_chajja, "chajja concrete + steel (proj*len*thk)")
    sp.add_argument("--proj", type=float, required=True)
    sp.add_argument("--length", type=float, required=True)
    sp.add_argument("--root", type=float, default=150)
    sp.add_argument("--tip", type=float, default=75)
    sp.add_argument("--n", type=float, default=1)
    sp.add_argument("--steel", type=float, default=1.0)

    sp = add("calc-chajja-shutter", cmd_calc_chajja_shutter, "chajja shuttering/formwork area")
    sp.add_argument("--proj", type=float, required=True)
    sp.add_argument("--length", type=float, required=True)
    sp.add_argument("--root", type=float, default=150)
    sp.add_argument("--tip", type=float, default=75)
    sp.add_argument("--n", type=float, default=1)

    sp = add("calc-chajja-plaster", cmd_calc_chajja_plaster, "chajja plaster area + mortar")
    sp.add_argument("--proj", type=float, required=True)
    sp.add_argument("--length", type=float, required=True)
    sp.add_argument("--tip", type=float, default=75)
    sp.add_argument("--n", type=float, default=1)
    sp.add_argument("--thk", type=float, default=12)

    # --- Foundation suite ---
    sp = add("foundation", cmd_foundation, "foundation takeoff (auto from drawing)")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 to isolate one sheet copy")
    sp.add_argument("--thk", type=float, default=450, help="footing thickness (mm)")
    sp.add_argument("--depth", type=float, default=1.5, help="foundation depth (m)")
    sp.add_argument("--pcc", type=float, default=100, help="PCC bed thickness (mm)")
    sp.add_argument("--ws", type=float, default=300, help="working space each side (mm)")
    sp.add_argument("--steel", type=float, default=0.8, help="steel percent")

    sp = add("footing-schedule", cmd_footing_schedule, "footing schedule (each F-tag)")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 to isolate one sheet copy")

    sp = add("concrete", cmd_concrete, "total RCC concrete (auto from drawing)")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 to isolate one sheet copy")
    sp.add_argument("--layer", help="slab boundary layer (else footprint)")
    sp.add_argument("--thk", type=float, default=150, help="slab thickness (mm)")
    sp.add_argument("--levels", help="FLOOR:elev,... override (metres)")

    sp = add("priced-boq", cmd_priced_boq, "priced BOQ: drawing takeoff x rate library")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 to isolate one sheet copy")
    sp.add_argument("--layer", help="slab boundary layer (else footprint)")
    sp.add_argument("--thk", type=float, default=150, help="slab thickness (mm)")
    sp.add_argument("--foot-thk", dest="foot_thk", type=float, default=450,
                    help="footing thickness (mm)")
    sp.add_argument("--depth", type=float, default=1.5, help="excavation depth (m)")
    sp.add_argument("--ws", type=float, default=300, help="working space each side (mm)")
    sp.add_argument("--steel-kg", dest="steel_kg", type=float, default=100,
                    help="reinforcement steel (kg per m3 concrete)")
    sp.add_argument("--floor-height", dest="floor_height", type=float,
                    help="floor-to-floor height (m) if the drawing has no level marks")

    sp = add("rates", cmd_rates, "view the rate library (edit rates.json to change)")

    # --- Slab suite (slab = auto from drawing) ---
    sp = add("slab", cmd_slab, "auto slab takeoff from the drawing")
    file_arg(sp); sp.add_argument("--layer", help="slab/boundary layer (else footprint)")
    sp.add_argument("--thk", type=float, default=150, help="slab thickness (mm)")
    sp.add_argument("--steelrate", type=float, default=10, help="steel kg/m2")
    sp.add_argument("--ptk", type=float, default=12, help="ceiling plaster thk (mm)")
    sp.add_argument("--floors", type=int, help="override floor count")

    sp = add("calc-slab", cmd_calc_slab, "slab concrete + steel (manual)")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--thk", type=float, default=150)
    sp.add_argument("--n", type=float, default=1)
    sp.add_argument("--steelrate", type=float, default=10)

    sp = add("calc-slab-steel", cmd_calc_slab_steel, "slab reinforcement (kg/m2 or percent)")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--floors", type=float, default=1)
    sp.add_argument("--rate", type=float, default=10)
    sp.add_argument("--vol", type=float, default=0, help="concrete m3 for percent method")
    sp.add_argument("--pct", type=float, default=0.9)

    sp = add("calc-slab-shutter", cmd_calc_slab_shutter, "slab shuttering / formwork")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--perimeter", type=float, default=0)
    sp.add_argument("--thk", type=float, default=150)
    sp.add_argument("--n", type=float, default=1)

    sp = add("calc-slab-plaster", cmd_calc_slab_plaster, "slab ceiling plaster + mortar")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--n", type=float, default=1)
    sp.add_argument("--thk", type=float, default=12)

    sp = add("colschedule", cmd_colschedule, "full column schedule (each tag)")
    file_arg(sp)

    sp = add("staircases", cmd_staircases, "count staircases + lifts")
    file_arg(sp)

    sp = add("beams", cmd_beams, "count + length of beams on a layer")
    file_arg(sp); sp.add_argument("--layer", help="beam layer name")

    sp = add("footings", cmd_footings, "count footings")
    file_arg(sp); sp.add_argument("--layer", help="footing layer name")

    sp = add("heights", cmd_heights, "floor-to-floor + building heights")
    file_arg(sp)

    sp = add("builtup", cmd_builtup, "built-up footprint area")
    file_arg(sp); sp.add_argument("--layer", help="boundary layer (cleaner)")

    sp = add("doors", cmd_doors, "count door blocks")
    file_arg(sp); sp.add_argument("--name", help="block-name pattern")

    sp = add("windows", cmd_windows, "count window blocks")
    file_arg(sp); sp.add_argument("--name", help="block-name pattern")

    sp = add("wall-area", cmd_wall_area, "wall area for plaster/paint")
    file_arg(sp); sp.add_argument("--layer", help="wall layer name")
    sp.add_argument("--height", type=float, default=3.0, help="wall height (m)")

    sp = add("findtext", cmd_findtext, "find text on the drawing")
    file_arg(sp); sp.add_argument("--q", help="keyword to find")
    sp.add_argument("--limit", type=int, default=30, help="max matches")

    sp = add("layer-detail", cmd_layer_detail, "entity breakdown for a layer")
    file_arg(sp); sp.add_argument("--layer", help="layer name")

    sp = add("bbox-layer", cmd_bbox_layer, "size/footprint of a layer")
    file_arg(sp); sp.add_argument("--layer", help="layer name")

    sp = add("column-spacing", cmd_column_spacing, "column grid spacing")
    file_arg(sp); sp.add_argument("--box", help="x0,y0,x1,y1 (auto from building)")
    sp.add_argument("--floor", help="floor name (default GROUND)")

    sp = add("room-areas", cmd_room_areas, "area of each room on a layer")
    file_arg(sp); sp.add_argument("--layer", help="room-boundary layer")
    sp.add_argument("--limit", type=int, default=40, help="max rooms")
    sp.add_argument("--min", type=float, default=1.0, help="min room m2")
    sp.add_argument("--max", type=float, default=100000.0, help="max room m2")

    sp = add("layer-area", cmd_layer_area, "closed area on one layer")
    file_arg(sp); sp.add_argument("--layer", help="layer name")

    sp = add("layer-length", cmd_layer_length, "length on one layer")
    file_arg(sp); sp.add_argument("--layer", help="layer name")

    sp = add("entity-count", cmd_entity_count, "count entities by type")
    file_arg(sp); sp.add_argument("--type", help="DXF type e.g. LINE, INSERT")

    sp = add("fixtures", cmd_fixtures, "count fixtures (blocks) on a layer")
    file_arg(sp); sp.add_argument("--layer", help="layer name")
    sp.add_argument("--name", help="block-name pattern")

    # --- Drawing QA / health suite (read-only) ---
    sp = add("qa-report", cmd_qa_report, "one-tap drawing health / error report")
    file_arg(sp)

    sp = add("dxf-diff", cmd_dxf_diff, "compare two drawings / revisions")
    file_arg(sp)
    sp.add_argument("--other", required=True,
                    help="other building name (or .dxf path) to compare with")

    sp = add("attr-audit", cmd_attr_audit, "block attribute + wrong-layer audit")
    file_arg(sp)

    sp = add("parking", cmd_parking, "count parking spaces")
    file_arg(sp); sp.add_argument("--name", help="block-name pattern")

    sp = add("sanitary", cmd_sanitary, "count toilets / sanitary fixtures")
    file_arg(sp); sp.add_argument("--name", help="block-name pattern")

    sp = add("area-statement", cmd_area_statement, "one-tap area statement")
    file_arg(sp); sp.add_argument("--layer", help="boundary layer (cleaner)")

    # --- all-layer analysis ---
    sp = add("layer-report", cmd_layer_report, "full breakdown of EVERY layer")
    file_arg(sp)

    sp = add("layer-blocks", cmd_layer_blocks, "blocks per layer (all layers)")
    file_arg(sp)

    sp = add("layer-texts", cmd_layer_texts, "text count per layer (all layers)")
    file_arg(sp)

    sp = add("layer-bounds", cmd_layer_bounds, "size per layer (all layers)")
    file_arg(sp)

    sp = add("which-layer", cmd_which_layer, "find which layers hold X")
    file_arg(sp); sp.add_argument("--type", help="DXF type e.g. INSERT")
    sp.add_argument("--q", help="text keyword")

    sp = add("calc-tiles", cmd_calc_tiles, "tiles for a floor area")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--tile", type=float, default=600)
    sp.add_argument("--waste", type=float, default=10)

    sp = add("calc-excavation", cmd_calc_excavation, "excavation L*B*D")
    sp.add_argument("--l", type=float, required=True)
    sp.add_argument("--b", type=float, required=True)
    sp.add_argument("--d", type=float, required=True)
    sp.add_argument("--bulk", type=float, default=25)

    sp = add("calc-steel", cmd_calc_steel, "rebar weight from concrete percentage")
    sp.add_argument("--vol", type=float, required=True)
    sp.add_argument("--pct", type=float, default=1.0)

    sp = add("calc-stair", cmd_calc_stair, "stair steps from floor height")
    sp.add_argument("--height", type=float, required=True)
    sp.add_argument("--riser", type=float, default=150)
    sp.add_argument("--tread", type=float, default=300)

    sp = add("calc-watertank", cmd_calc_watertank, "tank capacity (litres)")
    sp.add_argument("--l", type=float, required=True)
    sp.add_argument("--b", type=float, required=True)
    sp.add_argument("--h", type=float, required=True)

    sp = add("calc-cost", cmd_calc_cost, "amount = qty x rate")
    sp.add_argument("--qty", type=float, required=True)
    sp.add_argument("--rate", type=float, required=True)

    sp = add("convert", cmd_convert, "unit conversion (m2<->ft2 etc)")
    sp.add_argument("--value", type=float, required=True)
    sp.add_argument("--unit", required=True, help="m2-ft2 | m-ft | m3-ft3 ...")

    sp = add("formulas", cmd_formulas, "show every formula + its basis")
    sp.add_argument("--q", help="filter, e.g. concrete")

    # --- Sector: MEP / cleanroom / HVAC ---
    sp = add("calc-ach", cmd_calc_ach, "HVAC supply air from air changes")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--height", type=float, default=3.0)
    sp.add_argument("--ach", type=float, required=True)

    sp = add("calc-water-demand", cmd_calc_water_demand, "daily water demand")
    sp.add_argument("--persons", type=float, required=True)
    sp.add_argument("--lpcd", type=float, default=45)

    sp = add("calc-hvac", cmd_calc_hvac, "rough cooling load (TR)")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--persqm", type=float, default=10)

    sp = add("cleanroom-ref", cmd_cleanroom_ref, "cleanroom grade/ACH reference")

    sp = add("calc-electrical", cmd_calc_electrical, "electrical load (kVA)")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--va", type=float, default=60)
    sp.add_argument("--demand", type=float, default=80)

    sp = add("calc-lighting", cmd_calc_lighting, "light fixtures for lux")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--lux", type=float, default=300, help="target lux (office ~300)")
    sp.add_argument("--lumen", type=float, default=4000)
    sp.add_argument("--uf", type=float, default=0.5)
    sp.add_argument("--mf", type=float, default=0.8)

    sp = add("calc-rainwater", cmd_calc_rainwater, "roof runoff + downpipe")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--intensity", type=float, default=150)

    sp = add("calc-sewage", cmd_calc_sewage, "sewage / STP load")
    sp.add_argument("--water", type=float, required=True)
    sp.add_argument("--factor", type=float, default=0.8)

    sp = add("calc-fire-area", cmd_calc_fire_area, "fire compartment check")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--limit", type=float, default=750)

    sp = add("calc-travel", cmd_calc_travel, "travel-distance to exit check")
    sp.add_argument("--distance", type=float, required=True)
    sp.add_argument("--max", type=float, default=30)

    # --- Sector: code & compliance ---
    sp = add("calc-fsi", cmd_calc_fsi, "FSI/FAR consumed")
    sp.add_argument("--builtup", type=float, required=True)
    sp.add_argument("--plot", type=float, required=True)

    sp = add("calc-coverage", cmd_calc_coverage, "ground coverage percent")
    sp.add_argument("--footprint", type=float, required=True)
    sp.add_argument("--plot", type=float, required=True)

    sp = add("calc-occupancy", cmd_calc_occupancy, "occupant load + exits")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--factor", type=float, default=10)

    sp = add("calc-parking", cmd_calc_parking, "parking spaces (ECS)")
    sp.add_argument("--builtup", type=float, required=True)
    sp.add_argument("--per", type=float, default=100)

    sp = add("calc-ramp", cmd_calc_ramp, "ramp run length at a slope")
    sp.add_argument("--height", type=float, required=True)
    sp.add_argument("--slope", type=float, default=12)

    sp = add("calc-exit-width", cmd_calc_exit_width, "required egress width")
    sp.add_argument("--occupants", type=float, required=True)
    sp.add_argument("--permm", type=float, default=5)

    sp = add("door-width-check", cmd_door_width_check, "flag narrow openings")
    file_arg(sp); sp.add_argument("--min", type=int, default=1000,
                                  help="min clear width (mm)")
    sp.add_argument("--minheight", type=int, default=1800,
                    help="min height to treat as a door (mm)")

    sp = add("room-count", cmd_room_count, "count rooms by keyword")
    file_arg(sp); sp.add_argument("--keywords", help="comma list, e.g. LAB,OT,WARD")

    sp = add("coving-length", cmd_coving_length, "coving / skirting length")
    file_arg(sp); sp.add_argument("--layer", help="room-boundary layer")

    sp = add("export", cmd_export, "export entities to CSV")
    file_arg(sp); sp.add_argument("--out", help="output .csv path")

    sp = add("floorcols", cmd_floorcols, "columns per floor plan")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 to isolate one sheet copy")

    sp = add("colvol", cmd_colvol, "column concrete volume per variant")
    file_arg(sp)
    sp.add_argument("--box", help="x0,y0,x1,y1 to isolate one sheet copy")
    sp.add_argument("--levels", help="FLOOR:elev,... override (metres)")

    # --- project & registry ---
    sp = add("project", cmd_project, "manage plot/building manifest")
    sp.add_argument("action", choices=["init", "add", "remove", "list", "box"])
    sp.add_argument("name", nargs="?", help="building name")
    sp.add_argument("file", nargs="?",
                    help="building's .dxf (for add) / box value (for box)")
    sp.add_argument("--plot", help="plot name (for init)")
    sp.add_argument("--note", help="note for the building (for add)")

    sp = add("tools", cmd_tools, "list every command (registry for the AI)")
    sp.add_argument("--json", action="store_true", help="machine-readable")

    add("catalog", cmd_catalog, "tool catalog as JSON (for the web UI)")

    add("dwghelp", cmd_dwghelp, "how to convert .dwg -> .dxf")

    # --- Excel / BOQ ---
    sp = add("xls", cmd_xls, "list spreadsheet sheets")
    sp.add_argument("file")

    sp = add("xlsdump", cmd_xlsdump, "dump a sheet's rows")
    sp.add_argument("file"); sp.add_argument("--sheet", help="sheet name")

    sp = add("find", cmd_find, "search a BOQ for an item keyword")
    sp.add_argument("file"); sp.add_argument("keyword")
    sp.add_argument("--cols", help="desc,qty,unit,rate,amt column indexes")

    sp = add("compare", cmd_compare, "compare a keyword's qty across 2 BOQs")
    sp.add_argument("file_a"); sp.add_argument("file_b")
    sp.add_argument("keyword")

    # --- calculators ---
    sp = add("calc-concrete", cmd_calc_concrete, "concrete volume L*B*H*N")
    sp.add_argument("--l", type=float, required=True)
    sp.add_argument("--b", type=float, required=True)
    sp.add_argument("--h", type=float, required=True)
    sp.add_argument("--n", type=float, default=1)

    sp = add("calc-rebar", cmd_calc_rebar, "rebar weight (d^2/162)")
    sp.add_argument("--dia", type=float, required=True)
    sp.add_argument("--len", type=float, required=True)
    sp.add_argument("--n", type=float, default=1)

    sp = add("calc-brick", cmd_calc_brick, "brick count for wall area")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--thk", type=float, default=230)

    sp = add("calc-plaster", cmd_calc_plaster, "plaster mortar volume")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--thk", type=float, default=12)

    sp = add("calc-paint", cmd_calc_paint, "paint litres for area")
    sp.add_argument("--area", type=float, required=True)
    sp.add_argument("--coats", type=int, default=2)

    # --- auto-discovered plugins (AI-saved tools) ---
    _TYPES = {"int": int, "float": float, "str": str}
    for meta, mod in load_plugins():
        name = meta["command"]
        sp = add(name, _make_plugin_func(mod),
                 (meta.get("summary", "") + "  [plugin]"))
        file_arg(sp)
        for a in meta.get("args", []):
            kw = {"help": a.get("help", "")}
            if a.get("type") == "flag":
                kw["action"] = "store_true"
            else:
                kw["type"] = _TYPES.get(a.get("type", "str"), str)
                if "default" in a:
                    kw["default"] = a["default"]
                if a.get("required"):
                    kw["required"] = True
            sp.add_argument(a["name"], **kw)

    return p


def _make_plugin_func(mod):
    """Wrap a plugin's run(args, ctx) so argparse can call it as func(args)."""
    def _run(args):
        return mod.run(args, CTX)
    return _run


def _dispatch(args):
    """Resolve --building / --all (+ stored box), then run the command.

    Shared by the CLI (main) and the in-process API (run_command), so a
    front-end never needs to shell out to the command line.
    """
    def apply_box(entry):
        if "box" in entry and hasattr(args, "box") and not getattr(args, "box", None):
            args.box = entry["box"]

    if hasattr(args, "file") and args.func not in (cmd_project,):
        if getattr(args, "all", False):
            buildings = load_project().get("buildings", {})
            if not buildings:
                sys.exit("No buildings in project.json. Add some first.")
            for nm, entry in buildings.items():
                print(f"\n===== {nm}  ({entry['dxf']}) =====")
                args.file = entry["dxf"]
                apply_box(entry)
                args.func(args)
            return
        if not args.file and getattr(args, "building", None):
            entry = load_project().get("buildings", {}).get(args.building)
            if not entry:
                args.file = resolve_file(args)  # triggers the helpful error
            else:
                args.file = entry["dxf"]
                apply_box(entry)

    args.func(args)


def run_command(argv):
    """Run a tool IN-PROCESS and return its text output (no CLI, no subprocess).

    `argv` is the same token list the CLI would take, e.g.
        run_command(["columns", "--building", "AC1"])
    This is the doorway a web/desktop app should use: it imports ArchTools and
    calls the function directly, capturing what the tool prints.
    """
    import io
    import contextlib
    parser = build_parser()
    buf = io.StringIO()
    try:
        args = parser.parse_args(argv)
        with contextlib.redirect_stdout(buf):
            _dispatch(args)
    except SystemExit as e:           # argparse / sys.exit -> friendly text
        msg = str(e)
        out = buf.getvalue()
        if msg and not msg.isdigit():
            out += ("\n" if out else "") + msg
        return out.strip() or "error"
    return buf.getvalue().strip()


def main(argv=None):
    _dispatch(build_parser().parse_args(argv))


if __name__ == "__main__":
    main()
