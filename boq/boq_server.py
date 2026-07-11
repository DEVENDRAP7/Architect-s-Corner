#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
boq_server.py -- BOQ Comparer web app
=====================================

Upload YOUR BOQ, then each contractor's quote (Excel / PDF / photo).
One click -> item-matched rate comparison, missing items, lowball/high
flags, totals and ranking.

Run:  python boq_server.py   ->  http://127.0.0.1:8090
"""

import os
import shutil
from typing import List

from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import FileResponse, JSONResponse

import boq_engine

HERE = os.path.dirname(os.path.abspath(__file__))
UPLOADS = os.path.join(HERE, "uploads")
os.makedirs(UPLOADS, exist_ok=True)

app = FastAPI(title="BOQ Comparer")

STATE = {"owner": None, "contractors": []}   # parse results, in memory


@app.get("/")
def index():
    return FileResponse(os.path.join(HERE, "index.html"))


@app.post("/api/upload")
async def upload(role: str = Form(...), name: str = Form(""),
                 files: List[UploadFile] = Form(...)):
    paths = []
    for file in files:
        safe = os.path.basename(file.filename or "boq")
        dest = os.path.join(UPLOADS, safe)
        with open(dest, "wb") as f:
            shutil.copyfileobj(file.file, f)
        paths.append(dest)
    parsed = boq_engine.parse_many(paths)
    if name.strip():
        parsed["source"] = name.strip()
    if not parsed["items"] and role == "owner":
        return JSONResponse({"error": "could not read any BOQ items: "
                             + "; ".join(parsed["warnings"])}, status_code=400)
    if role == "owner":
        STATE["owner"] = parsed
    else:
        # replace an earlier upload with the same name
        STATE["contractors"] = [c for c in STATE["contractors"]
                                if c["source"] != parsed["source"]]
        STATE["contractors"].append(parsed)
    return {"ok": True, "source": parsed["source"], "kind": parsed["kind"],
            "items": len(parsed["items"]), "warnings": parsed["warnings"]}


@app.get("/api/state")
def state():
    return {
        "owner": ({"source": STATE["owner"]["source"],
                   "items": len(STATE["owner"]["items"])}
                  if STATE["owner"] else None),
        "contractors": [{"source": c["source"], "items": len(c["items"]),
                         "warnings": c["warnings"]}
                        for c in STATE["contractors"]],
    }


@app.post("/api/remove")
def remove(source: str = Form(...)):
    STATE["contractors"] = [c for c in STATE["contractors"]
                            if c["source"] != source]
    return {"ok": True}


@app.post("/api/reset")
def reset():
    STATE["owner"] = None
    STATE["contractors"] = []
    return {"ok": True}


@app.get("/api/compare")
def do_compare():
    if not STATE["owner"]:
        return JSONResponse({"error": "upload your BOQ first"}, status_code=400)
    if not STATE["contractors"]:
        return JSONResponse({"error": "upload at least one contractor quote"},
                            status_code=400)
    return boq_engine.compare(STATE["owner"], STATE["contractors"])


if __name__ == "__main__":
    import uvicorn
    print("BOQ Comparer -> http://127.0.0.1:8090")
    uvicorn.run(app, host="127.0.0.1", port=8090, log_level="warning")
