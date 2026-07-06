#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
site_server.py -- backend for the Architect's Corner landing page
=================================================================

Serves the marketing page WITH real analytics and a real product demo:

  * total visits / unique visitors / today's visits  (persisted to JSON)
  * "watching now" -- live presence via a 10s heartbeat from each browser
  * peak concurrent watchers
  * demo runs counter
  * /api/demo runs the REAL ArchTools engine on a bundled sample drawing
    (generated with ezdxf at startup), so "Simulate" shows genuine output.

Run:  python site_server.py   ->  http://127.0.0.1:8080
"""

import datetime
import json
import os
import sys
import threading
import time

from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)          # ArchTools.py lives one level up
import ArchTools  # noqa: E402

STATS_FILE = os.path.join(HERE, "site_stats.json")
SAMPLE_DXF = os.path.join(HERE, "sample_plan.dxf")
LOCK = threading.Lock()
ACTIVE = {}          # sid -> last heartbeat timestamp
WINDOW = 30          # seconds of silence before a watcher is "gone"

app = FastAPI(title="Architect's Corner — site")


# ----------------------------------------------------------------------
# stats persistence
# ----------------------------------------------------------------------
def _load():
    if os.path.isfile(STATS_FILE):
        with open(STATS_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"total": 0, "unique": 0, "days": {}, "demo_runs": 0, "peak": 0}


def _save(s):
    with open(STATS_FILE, "w", encoding="utf-8") as f:
        json.dump(s, f, indent=1)


STATS = _load()


def _today():
    return datetime.date.today().isoformat()


def _active_count():
    now = time.time()
    dead = [k for k, t in ACTIVE.items() if now - t > WINDOW]
    for k in dead:
        ACTIVE.pop(k, None)
    return len(ACTIVE)


# ----------------------------------------------------------------------
# sample drawing -- generated once so the demo runs the REAL engine
# ----------------------------------------------------------------------
def build_sample():
    import ezdxf
    doc = ezdxf.new()
    msp = doc.modelspace()

    def txt(s, x, y, layer="ANNOT"):
        msp.add_mtext(s, dxfattribs={"layer": layer}).set_location((x, y))

    # floor level marks (read by extract_levels)
    for s, y in [("GROUND FLOOR LVL +0.00", 0), ("FIRST FLOOR LVL +3.50", -900),
                 ("SECOND FLOOR LVL +7.00", -1800), ("TERRACE FLOOR LVL +10.50", -2700)]:
        txt(s, -15000, y, "LEVELS")

    # two floor plans with titles + tagged columns (C-tags) on a 6x4 m grid
    plans = [("GROUND FLOOR PLAN", 0), ("FIRST FLOOR PLAN", 40000)]
    cols = [("C1", "450x450"), ("C2", "450x450"), ("C3", "300x450"),
            ("C4", "450x450"), ("C5", "300x450"), ("C6", "300x300")]
    for title, ox in plans:
        txt(title, ox + 9000, 14000, "TITLES")
        for i, (tag, sz) in enumerate(cols):
            x, y = ox + (i % 3) * 6000, (i // 3) * 4000
            txt(f"{tag}\n{sz}", x, y, "COLUMNS")
            w, h = (int(v) for v in sz.split("x"))
            msp.add_lwpolyline(
                [(x - w/2, y - h/2), (x + w/2, y - h/2),
                 (x + w/2, y + h/2), (x - w/2, y + h/2)],
                close=True, dxfattribs={"layer": "COLUMNS"})

    # footing markers (F-tags) below the ground plan
    for i, (tag, sz) in enumerate([("F1", "1500x1500"), ("F2", "1500x1500"),
                                   ("F3", "1200x1200")]):
        txt(f"{tag}\n{sz}", i * 6000, -8000, "FOOTINGS")

    # building boundary (slab footprint 12 x 8 m)
    msp.add_lwpolyline([(-3000, -2000), (15000, -2000), (15000, 6000),
                        (-3000, 6000)], close=True,
                       dxfattribs={"layer": "BOUNDARY"})

    # deliberate defects so the QA demo finds something
    msp.add_line((0, -5000), (5000, -5000), dxfattribs={"layer": "WALLS"})
    msp.add_line((0, -5000), (5000, -5000), dxfattribs={"layer": "WALLS"})  # duplicate
    msp.add_line((100, -5200), (104, -5200), dxfattribs={"layer": "WALLS"})  # stray
    msp.add_lwpolyline([(20000, 0), (24000, 0), (24000, 3000), (20010, 20)],
                       dxfattribs={"layer": "WALLS"})  # nearly closed
    txt("STORE", 26000, 0, "TEXT")
    txt("STORE ROOM", 26040, 30, "TEXT")  # overlapping text
    doc.layers.add("UNUSED-LAYER")        # empty layer

    doc.saveas(SAMPLE_DXF)


if not os.path.isfile(SAMPLE_DXF):
    build_sample()

# tools the demo may run (whitelist -- nothing else is callable)
DEMO_TOOLS = {
    "columns":   ["columns", SAMPLE_DXF],
    "concrete":  ["concrete", SAMPLE_DXF],
    "foundation": ["foundation", SAMPLE_DXF],
    "qa-report": ["qa-report", SAMPLE_DXF],
    "levels":    ["levels", SAMPLE_DXF],
    "area-statement": ["area-statement", SAMPLE_DXF],
}


# ----------------------------------------------------------------------
# routes
# ----------------------------------------------------------------------
@app.get("/")
def index():
    return FileResponse(os.path.join(HERE, "index.html"))


class Visit(BaseModel):
    new: bool = False


@app.post("/api/visit")
def visit(v: Visit):
    with LOCK:
        STATS["total"] += 1
        if v.new:
            STATS["unique"] += 1
        STATS["days"][_today()] = STATS["days"].get(_today(), 0) + 1
        _save(STATS)
        return {"visitor_no": STATS["total"]}


class Beat(BaseModel):
    sid: str


@app.post("/api/beat")
def beat(b: Beat):
    ACTIVE[b.sid[:64]] = time.time()
    n = _active_count()
    with LOCK:
        if n > STATS.get("peak", 0):
            STATS["peak"] = n
            _save(STATS)
    return {"active": n}


@app.get("/api/stats")
def stats():
    with LOCK:
        return {
            "total": STATS["total"],
            "unique": STATS["unique"],
            "today": STATS["days"].get(_today(), 0),
            "active": _active_count(),
            "peak": STATS.get("peak", 0),
            "demo_runs": STATS.get("demo_runs", 0),
        }


MSG_FILE = os.path.join(HERE, "site_messages.json")
ADMIN_FILE = os.path.join(HERE, "site_admin.json")
OWNER_EMAIL = "devendranprajapati07@gmail.com"


def _admin_key():
    """Stable secret admin key, generated once and kept out of git."""
    if os.path.isfile(ADMIN_FILE):
        with open(ADMIN_FILE, encoding="utf-8") as f:
            return json.load(f)["key"]
    import secrets
    key = secrets.token_urlsafe(18)
    with open(ADMIN_FILE, "w", encoding="utf-8") as f:
        json.dump({"key": key}, f)
    return key


ADMIN_KEY = _admin_key()
print(f"ADMIN dashboard -> http://127.0.0.1:8080/admin?key={ADMIN_KEY}")


def _messages():
    if os.path.isfile(MSG_FILE):
        with open(MSG_FILE, encoding="utf-8") as f:
            return json.load(f)
    return []


def _notify_email(entry):
    """Email the owner about a new message. Activates when SMTP creds are
    set (Gmail: 2FA + app password):
        set AC_SMTP_USER=you@gmail.com
        set AC_SMTP_PASS=<16-char app password>
    Silent no-op otherwise -- the message is stored regardless."""
    user = os.environ.get("AC_SMTP_USER")
    pw = os.environ.get("AC_SMTP_PASS")
    if not (user and pw):
        return
    try:
        import smtplib
        from email.mime.text import MIMEText
        body = (f"New message on the Architect's Corner site\n\n"
                f"Time   : {entry['time']}\n"
                f"Name   : {entry['name']}\n"
                f"Email  : {entry['email']}\n\n{entry['message']}")
        m = MIMEText(body, _charset="utf-8")
        m["Subject"] = f"[Architect's Corner] message from {entry['name']}"
        m["From"] = user
        m["To"] = OWNER_EMAIL
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15) as s:
            s.login(user, pw)
            s.send_message(m)
    except Exception:
        pass  # never lose the message over a mail hiccup


class Contact(BaseModel):
    name: str
    email: str
    message: str


@app.post("/api/contact")
def contact(c: Contact):
    """Store a contact-form message (name/email/message + timestamp)."""
    if not c.message.strip():
        return JSONResponse({"error": "empty message"}, status_code=400)
    entry = {
        "time": datetime.datetime.now().isoformat(timespec="seconds"),
        "name": c.name.strip()[:120],
        "email": c.email.strip()[:200],
        "message": c.message.strip()[:4000],
    }
    with LOCK:
        msgs = _messages()
        msgs.append(entry)
        with open(MSG_FILE, "w", encoding="utf-8") as f:
            json.dump(msgs, f, indent=1, ensure_ascii=False)
    threading.Thread(target=_notify_email, args=(entry,), daemon=True).start()
    return {"ok": True, "received": entry["time"]}


# ----------------------------------------------------------------------
# owner dashboard -- messages + stats + CSV export (key-protected)
# ----------------------------------------------------------------------
def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


@app.get("/admin")
def admin(key: str = ""):
    if key != ADMIN_KEY:
        return JSONResponse({"error": "wrong key"}, status_code=403)
    with LOCK:
        s = dict(STATS)
        msgs = list(reversed(_messages()))
    days = sorted(s.get("days", {}).items(), reverse=True)[:14]
    rows = "".join(
        f"<tr><td>{_esc(m['time'])}</td><td>{_esc(m['name'])}</td>"
        f"<td><a href='mailto:{_esc(m['email'])}'>{_esc(m['email'])}</a></td>"
        f"<td>{_esc(m['message'])}</td></tr>" for m in msgs) or \
        "<tr><td colspan=4>(no messages yet)</td></tr>"
    dayrows = "".join(f"<tr><td>{d}</td><td>{n}</td></tr>" for d, n in days)
    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>AC — owner dashboard</title><style>
body{{font-family:system-ui;background:#0d0e11;color:#edece7;margin:0;padding:34px;font-size:14px}}
h1{{font-size:18px}} h2{{font-size:13px;text-transform:uppercase;letter-spacing:.1em;color:#8f9097;margin-top:34px}}
.grid{{display:flex;gap:12px;flex-wrap:wrap}}
.card{{background:#131418;border:1px solid #26272d;border-radius:10px;padding:16px 20px;min-width:130px}}
.card b{{font-size:26px;font-family:ui-monospace,monospace;display:block}}
.card i{{font-style:normal;font-size:10px;text-transform:uppercase;letter-spacing:.12em;color:#5c5d65}}
table{{border-collapse:collapse;width:100%;margin-top:10px;background:#131418;border:1px solid #26272d;border-radius:8px}}
th,td{{padding:8px 12px;border-bottom:1px solid #26272d;text-align:left;vertical-align:top;font-size:13px}}
th{{color:#8f9097;font-size:10.5px;text-transform:uppercase;letter-spacing:.1em}}
td:last-child{{white-space:pre-wrap;max-width:520px}}
a{{color:#e8ae3f}} .btn{{display:inline-block;margin-top:10px;border:1px solid #e8ae3f;color:#e8ae3f;
border-radius:7px;padding:7px 16px;text-decoration:none;font-size:13px}}</style></head><body>
<h1>Architect's Corner — owner dashboard</h1>
<div class="grid">
<div class="card"><b>{_active_count()}</b><i>watching now</i></div>
<div class="card"><b>{s.get('days', {}).get(_today(), 0)}</b><i>visits today</i></div>
<div class="card"><b>{s['total']}</b><i>total visits</i></div>
<div class="card"><b>{s['unique']}</b><i>unique visitors</i></div>
<div class="card"><b>{s.get('peak', 0)}</b><i>peak concurrent</i></div>
<div class="card"><b>{s.get('demo_runs', 0)}</b><i>demo runs</i></div>
<div class="card"><b>{len(msgs)}</b><i>messages</i></div>
</div>
<h2>Messages</h2>
<a class="btn" href="/admin/export?key={ADMIN_KEY}">⬇ Download CSV</a>
<table><tr><th>time</th><th>name</th><th>email</th><th>message</th></tr>{rows}</table>
<h2>Visits — last 14 days</h2>
<table><tr><th>date</th><th>visits</th></tr>{dayrows}</table>
</body></html>"""
    from fastapi.responses import HTMLResponse
    return HTMLResponse(html)


@app.get("/admin/export")
def admin_export(key: str = ""):
    if key != ADMIN_KEY:
        return JSONResponse({"error": "wrong key"}, status_code=403)
    import csv
    import io
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["time", "name", "email", "message"])
    for m in _messages():
        w.writerow([m["time"], m["name"], m["email"], m["message"]])
    from fastapi.responses import Response
    return Response(buf.getvalue(), media_type="text/csv",
                    headers={"Content-Disposition":
                             "attachment; filename=messages.csv"})


class Demo(BaseModel):
    tool: str


@app.post("/api/demo")
def demo(d: Demo):
    argv = DEMO_TOOLS.get(d.tool)
    if not argv:
        return JSONResponse({"error": "unknown demo tool"}, status_code=400)
    out = ArchTools.run_command(list(argv))
    with LOCK:
        STATS["demo_runs"] = STATS.get("demo_runs", 0) + 1
        _save(STATS)
    return {"output": out, "tool": d.tool}


if __name__ == "__main__":
    import uvicorn
    print("Landing page -> http://127.0.0.1:8080")
    uvicorn.run(app, host="127.0.0.1", port=8080, log_level="warning")
