#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
server.py  --  web backend for Architect's Corner
==================================================

Serves the single-page tool catalog and runs vetted ArchTools commands
in-process (no AI, no subprocess). Used both by `python server.py` in dev
and by the packaged desktop .exe (via launch.py).

Data (project.json + uploads) lives in ArchTools.DATA_DIR, which the desktop
launcher points at a writable folder (Documents\\ArchitectsCorner). The web
UI is served from the bundle (sys._MEIPASS when frozen).
"""

import os
import shutil
import sys

from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

import ArchTools  # tools + run_command + project manifest helpers

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = ArchTools.DATA_DIR
UPLOADS = os.path.join(DATA_DIR, "uploads")
PROJECT = ArchTools.PROJECT_FILE
os.makedirs(UPLOADS, exist_ok=True)


def _res(*parts):
    """Path to a bundled resource (works packaged and from source)."""
    base = getattr(sys, "_MEIPASS", HERE)
    return os.path.join(base, *parts)


app = FastAPI(title="Architect's Corner")


# ----------------------------------------------------------------------
# Page
# ----------------------------------------------------------------------
@app.get("/")
def index():
    return FileResponse(_res("web", "index.html"))


@app.get("/logo.svg")
def logo():
    return FileResponse(_res("web", "logo.svg"))


# ----------------------------------------------------------------------
# API
# ----------------------------------------------------------------------
@app.get("/api/buildings")
def buildings():
    proj = ArchTools.load_project(PROJECT)
    return {"plot": proj.get("plot", ""), "buildings": proj.get("buildings", {})}


@app.post("/api/plot")
def set_plot(name: str = Form(...)):
    proj = ArchTools.load_project(PROJECT)
    proj["plot"] = name
    proj.setdefault("buildings", {})
    ArchTools.save_project(proj, PROJECT)
    return {"ok": True, "plot": name}


@app.post("/api/upload")
async def upload(name: str = Form(...), note: str = Form(""),
                 file: UploadFile = Form(...)):
    if not name.strip():
        return JSONResponse({"error": "building name required"}, status_code=400)
    safe = os.path.basename(file.filename or "")
    if not safe.lower().endswith(".dxf"):
        return JSONResponse({"error": "please upload a .dxf file"},
                            status_code=400)
    dest = os.path.join(UPLOADS, f"{name.strip()}__{safe}")
    with open(dest, "wb") as f:
        shutil.copyfileobj(file.file, f)

    proj = ArchTools.load_project(PROJECT)
    proj.setdefault("buildings", {})[name.strip()] = {
        # absolute path so resolution is independent of the working directory
        "dxf": dest,
        "note": note.strip(),
    }
    ArchTools.save_project(proj, PROJECT)
    return {"ok": True, "name": name.strip(), "dxf": dest}


@app.post("/api/remove")
def remove(name: str = Form(...)):
    proj = ArchTools.load_project(PROJECT)
    proj.get("buildings", {}).pop(name, None)
    ArchTools.save_project(proj, PROJECT)
    return {"ok": True}


@app.get("/api/catalog")
def catalog():
    """The searchable tool catalog (no AI involved)."""
    return {"tools": ArchTools.build_catalog()}


@app.get("/api/measure")
def measure(building: str = ""):
    """Geometry read straight from a building's DXF (area/perimeter/floors),
    used to auto-fill calculator forms. Empty if the building/file is missing."""
    proj = ArchTools.load_project(PROJECT)
    entry = proj.get("buildings", {}).get(building)
    if not entry:
        return {"metrics": {}}
    return {"metrics": ArchTools.auto_metrics(entry.get("dxf", ""))}


class RunReq(BaseModel):
    command: str
    building: str = ""
    params: dict = {}


@app.post("/api/run")
def run_tool(req: RunReq):
    """Run ONE vetted ArchTools command in-process. Only commands that exist
    in the catalog can be called -- no code-writing, no shell."""
    cat = {t["command"]: t for t in ArchTools.build_catalog()}
    meta = cat.get(req.command)
    if not meta:
        return JSONResponse({"error": f"unknown tool '{req.command}'"},
                            status_code=400)

    argv = [req.command]
    if meta["needs_building"]:
        if not req.building:
            return JSONResponse({"error": "pick a building first"},
                                status_code=400)
        argv += ["--building", req.building]

    pmeta = {p["name"]: p for p in meta["params"]}
    for name, val in (req.params or {}).items():
        p = pmeta.get(name)
        if not p:
            continue
        if p["type"] == "flag":
            if val:
                argv.append(name)
        elif str(val).strip() != "":
            argv += [name, str(val)]

    output = ArchTools.run_command(argv)
    return {"output": output, "command": "ArchTools " + " ".join(argv)}


if __name__ == "__main__":
    import uvicorn
    print("Architect's Corner -> http://127.0.0.1:8000")
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="warning")
