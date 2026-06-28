#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
launch.py  --  desktop entry point for Architect's Corner
=========================================================

Double-clicked by the user (packaged as ArchitectsCorner.exe). It:
  * points the app's data at  Documents\\ArchitectsCorner  (writable, persists)
  * picks a free local port
  * starts the FastAPI server on 127.0.0.1
  * opens the default browser to the app

Everything runs on the user's own machine -- no internet, no host, no account.
"""

import os
import socket
import threading
import time
import webbrowser


def _data_dir():
    home = os.path.expanduser("~")
    docs = os.path.join(home, "Documents")
    base = docs if os.path.isdir(docs) else home
    d = os.path.join(base, "ArchitectsCorner")
    os.makedirs(d, exist_ok=True)
    return d


# Must be set BEFORE importing ArchTools/server (they read it at import time).
os.environ.setdefault("AC_DATA", _data_dir())

import uvicorn          # noqa: E402
from server import app  # noqa: E402


def _free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def main():
    port = int(os.environ.get("AC_PORT") or _free_port())
    url = f"http://127.0.0.1:{port}"

    def _open():
        time.sleep(1.3)
        webbrowser.open(url)
    threading.Thread(target=_open, daemon=True).start()

    print("=" * 54)
    print("  Architect's Corner")
    print(f"  Running at: {url}")
    print(f"  Your files: {os.environ['AC_DATA']}")
    print("  Keep this window open. Close it to quit the app.")
    print("=" * 54)
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")


if __name__ == "__main__":
    main()
