#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
boq_engine.py -- parse BOQs from Excel / PDF / images and compare them
=======================================================================

parse_boq(path)  -> {"source", "kind", "items":[{desc,unit,qty,rate,amount}], "warnings":[...]}
compare(owner, contractors) -> report dict (matched rate table, missing items,
                               lowball/high flags, totals, ranking)

Column positions are DETECTED from the header row (desc/unit/qty/rate/amount
keywords), not assumed. Item matching across files is exact-normalized first,
then fuzzy (difflib) -- so contractors' slightly reworded descriptions still
match. Images need tesseract OCR installed; without it they fail with a clear
message instead of a wrong table.
"""

import difflib
import os
import re
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from ArchTools import open_workbook  # uniform xls/xlsx reader (reused)

# ----------------------------------------------------------------------
# parsing
# ----------------------------------------------------------------------
HDR = {
    "desc":   re.compile(r"desc|particular|item of work|work item", re.I),
    "unit":   re.compile(r"^unit|uom", re.I),
    "qty":    re.compile(r"qty|quantity", re.I),
    "rate":   re.compile(r"rate", re.I),
    "amount": re.compile(r"amount|total|value", re.I),
}


def _num(v):
    """Cell -> float or None. Handles '1,234.50', '₹ 850', '850/-'."""
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def _detect_header(rows):
    """(row_index, {field: col}) for the first row matching >=3 headers."""
    for i, row in enumerate(rows[:40]):
        cols = {}
        for c, cell in enumerate(row):
            s = str(cell).strip()
            if not s:
                continue
            for field, pat in HDR.items():
                if field not in cols and pat.search(s):
                    cols[field] = c
        if "desc" in cols and "rate" in cols and len(cols) >= 3:
            return i, cols
    return None, None


def _rows_to_items(rows, warnings):
    hi, cols = _detect_header(rows)
    if hi is None:
        warnings.append("no header row found (need desc/qty/rate columns)")
        return []
    items = []
    for row in rows[hi + 1:]:
        # OCR/PDF sometimes glues the serial number onto the description,
        # shifting every cell left by one -- split it back out.
        if row and isinstance(row[0], str):
            m = re.match(r"^(\d{1,4})\s+(\D.{4,})$", row[0].strip())
            if m:
                row = [m.group(1), m.group(2)] + list(row[1:])
        if not row:
            continue
        # row lost its leading serial cell (common in OCR) -> everything
        # sits one column left of the header. Detect: the desc column holds
        # a short unit-ish token while the cell before it is real text.
        off = 0
        dc = cols["desc"]
        if 0 < dc < len(row) + 1:
            cand = str(row[dc]).strip() if dc < len(row) else ""
            prev = str(row[dc - 1]).strip()
            if (len(prev) >= 8 and any(ch.isalpha() for ch in prev)
                    and (not cand or (len(cand) <= 6 and " " not in cand))):
                off = -1
        if dc + off >= len(row):
            continue
        desc = str(row[dc + off]).strip()
        rc = cols["rate"] + off
        rate = _num(row[rc]) if 0 <= rc < len(row) else None
        if not desc or rate is None:
            continue  # section headings, totals, blanks
        def g(f):
            c = cols.get(f)
            c = c + off if c is not None else None
            return row[c] if c is not None and 0 <= c < len(row) else None
        items.append({
            "desc": desc,
            "unit": str(g("unit") or "").strip(),
            "qty": _num(g("qty")),
            "rate": rate,
            "amount": _num(g("amount")),
        })
    return items


def _parse_pdf(path, warnings):
    import pdfplumber
    rows = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                rows += [[("" if c is None else c) for c in r] for r in table]
    if not rows:
        # scanned PDF (pages are images) -> render each page and OCR it
        rows = _ocr_pdf(path, warnings)
    return rows


def _ocr_pdf(path, warnings):
    try:
        import pypdfium2 as pdfium
    except ImportError:
        warnings.append("scanned PDF needs: pip install pypdfium2")
        return []
    rows = []
    doc = pdfium.PdfDocument(path)
    for page in doc:
        img = page.render(scale=2.5).to_pil()
        rows += _ocr_rows(img, warnings)
    doc.close()
    if rows:
        warnings.append("scanned PDF parsed via OCR -- verify numbers "
                        "before trusting")
    else:
        warnings.append("no tables found in PDF (and OCR found nothing)")
    return rows


def _parse_image(path, warnings):
    try:
        from PIL import Image
    except ImportError:
        warnings.append("image OCR needs: pip install pytesseract pillow")
        return []
    rows = _ocr_rows(Image.open(path), warnings)
    if rows:
        warnings.append("image parsed via OCR -- verify numbers before trusting")
    return rows


def _ocr_rows(img, warnings):
    """PIL image -> table rows, via word coordinates (rows by y, cells by
    x-gaps). Plain OCR text reads tables column-wise and destroys rows."""
    try:
        import pytesseract
    except ImportError:
        warnings.append("image OCR needs: pip install pytesseract pillow "
                        "+ the Tesseract program (github.com/UB-Mannheim/tesseract)")
        return []
    # default Windows install path when tesseract isn't on PATH
    import shutil as _sh
    if not _sh.which("tesseract"):
        for p in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                  r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
            if os.path.isfile(p):
                pytesseract.pytesseract.tesseract_cmd = p
                break
    try:
        d = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    except Exception as e:
        warnings.append(f"OCR failed: {e}")
        return []
    words = [(d["top"][i], d["left"][i], d["width"][i], d["text"][i].strip())
             for i in range(len(d["text"])) if d["text"][i].strip()]
    if not words:
        warnings.append("OCR found no text in the image")
        return []
    words.sort()
    heights = [d["height"][i] for i in range(len(d["text"])) if d["text"][i].strip()]
    tol = max(8, statistics.median(heights) // 2)
    gap = max(20, img.width // 50)      # x-gap that separates table cells
    lines, cur, cur_top = [], [], None
    for top, left, w, t in words:
        if cur_top is None or abs(top - cur_top) <= tol:
            cur.append((left, w, t))
            cur_top = top if cur_top is None else cur_top
        else:
            lines.append(cur)
            cur, cur_top = [(left, w, t)], top
    lines.append(cur)
    rows = []
    for ln in lines:
        ln.sort()
        cells, prev_end = [], None
        for left, w, t in ln:
            if prev_end is not None and left - prev_end > gap:
                cells.append(t)
            elif cells:
                cells[-1] += " " + t
            else:
                cells.append(t)
            prev_end = left + w
        rows.append(cells)
    return rows


def _gather_rows(path, warnings):
    """(rows, kind) from one file of any supported type."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xls", ".xlsx", ".xlsm"):
        rows = []
        for _nm, sheet_rows in open_workbook(path):
            rows += sheet_rows
        return rows, "excel"
    if ext == ".pdf":
        return _parse_pdf(path, warnings), "pdf"
    if ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp"):
        return _parse_image(path, warnings), "image"
    warnings.append(f"unsupported file type {ext}")
    return [], "?"


def parse_boq(path):
    return parse_many([path])


def parse_many(paths):
    """Parse ONE BOQ spread across several files (e.g. 5 photos of one
    quote, or photos + a sheet). Rows are concatenated in upload order and
    parsed as a single table, so the header on page 1 also drives the
    continuation pages that have no header of their own."""
    warnings, rows, kinds = [], [], []
    for p in paths:
        r, k = _gather_rows(p, warnings)
        rows += r
        kinds.append(k)
    items = _rows_to_items(rows, warnings)
    if not items and not warnings:
        warnings.append("no BOQ items recognised")
    kind = "+".join(sorted(set(kinds))) if len(set(kinds)) > 1 else kinds[0]
    src = os.path.basename(paths[0]) + (f" (+{len(paths)-1} more)"
                                        if len(paths) > 1 else "")
    return {"source": src, "kind": kind, "items": items, "warnings": warnings}


# ----------------------------------------------------------------------
# matching + comparison
# ----------------------------------------------------------------------
def _norm(desc):
    return re.sub(r"[^a-z0-9 ]", "", desc.lower()).strip()


def _match(owner_items, other_items):
    """owner index -> other item (exact normalized, then fuzzy > 0.87)."""
    by_norm = {}
    for it in other_items:
        by_norm.setdefault(_norm(it["desc"]), it)
    taken, out = set(), {}
    for i, o in enumerate(owner_items):
        n = _norm(o["desc"])
        hit = by_norm.get(n)
        if hit is not None and id(hit) not in taken:
            out[i] = hit
            taken.add(id(hit))
    # fuzzy pass for the rest  # ponytail: O(n*m) difflib, fine to ~1k items
    rest = [it for it in other_items if id(it) not in taken]
    rest_norms = [_norm(it["desc"]) for it in rest]
    for i, o in enumerate(owner_items):
        if i in out or not rest:
            continue
        best = difflib.get_close_matches(_norm(o["desc"]), rest_norms,
                                         n=1, cutoff=0.87)
        if best:
            j = rest_norms.index(best[0])
            out[i] = rest[j]
            rest.pop(j)
            rest_norms.pop(j)
    return out


def compare(owner, contractors):
    """owner + [contractor parse results] -> comparison report."""
    names = [c["source"] for c in contractors]
    matches = [_match(owner["items"], c["items"]) for c in contractors]
    rows, totals = [], {n: 0.0 for n in names}
    owner_total = 0.0
    missing = {n: [] for n in names}
    for i, o in enumerate(owner["items"]):
        qty = o["qty"] or 0.0
        rates = {}
        for k, c in enumerate(contractors):
            hit = matches[k].get(i)
            if hit is None:
                missing[names[k]].append(o["desc"])
            else:
                rates[names[k]] = hit["rate"]
                totals[names[k]] += hit["rate"] * qty
        if o["rate"]:
            owner_total += o["rate"] * qty
        quoted = [r for r in rates.values() if r]
        med = statistics.median(quoted) if quoted else None
        flags = {}
        for n, r in rates.items():
            if med and len(quoted) > 1:
                if r < 0.6 * med:
                    flags[n] = "LOW"
                elif r > 1.5 * med:
                    flags[n] = "HIGH"
        rows.append({"desc": o["desc"], "unit": o["unit"], "qty": o["qty"],
                     "owner_rate": o["rate"], "rates": rates, "flags": flags,
                     "lowest": min(rates, key=rates.get) if rates else None})
    ranking = sorted(totals.items(), key=lambda kv: kv[1])
    return {
        "owner": owner["source"], "owner_items": len(owner["items"]),
        "owner_total": round(owner_total, 2),
        "contractors": [{
            "name": n,
            "items_quoted": len(owner["items"]) - len(missing[n]),
            "missing": missing[n],
            "total": round(totals[n], 2),
            "low_flags": sum(1 for r in rows if r["flags"].get(n) == "LOW"),
            "high_flags": sum(1 for r in rows if r["flags"].get(n) == "HIGH"),
        } for n in names],
        "ranking": [{"name": n, "total": round(t, 2)} for n, t in ranking],
        "rows": rows,
    }


# ----------------------------------------------------------------------
# self-check: make_boq's planted defects must be found exactly
# ----------------------------------------------------------------------
if __name__ == "__main__":
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    a = os.path.join(root, "Contractor-A.xlsx")
    b = os.path.join(root, "Contractor-B.xlsx")
    if not (os.path.isfile(a) and os.path.isfile(b)):
        sys.exit("run make_boq.py first (generates Contractor-A/B.xlsx)")
    A, B = parse_boq(a), parse_boq(b)
    assert len(A["items"]) >= 200, f"A parsed only {len(A['items'])}"
    rep = compare(A, [B])
    cb = rep["contractors"][0]
    assert len(cb["missing"]) == 3, f"expected 3 missing, got {len(cb['missing'])}"
    assert cb["low_flags"] == 0  # single contractor -> no median flags
    print(f"A items {len(A['items'])}, B items {len(B['items'])}, "
          f"B missing {len(cb['missing'])} (expected 3) -- OK")
    print("owner total", rep["owner_total"], "| B total", cb["total"])
