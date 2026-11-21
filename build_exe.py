#!/usr/bin/env python3
"""
Build the Windows app with PyInstaller.

    python -m pip install -r requirements-build.txt
    python build_exe.py

Result:
    dist/CAMIS Marks/CAMIS Marks.exe   (+ the _internal folder next to it)

Zip the whole "dist/CAMIS Marks" folder and share it. The target PC needs
Google Chrome or Microsoft Edge (every Windows 10/11 PC has Edge) and nothing else.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
NAME = "CAMIS Marks"
SEP = ";" if sys.platform == "win32" else ":"


def main() -> int:
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name", NAME,
        "--icon", str(HERE / "assets" / "icon.ico"),
        "--add-data", f"{HERE / 'ui'}{SEP}ui",
        "--add-data", f"{HERE / 'assets' / 'icon.png'}{SEP}assets",
        "--add-data", f"{HERE / 'HOW TO USE.txt'}{SEP}.",
        "--collect-all", "playwright",
        "--collect-all", "webview",
        "--hidden-import", "openpyxl",
        "--hidden-import", "xlrd",
        str(HERE / "app.py"),
    ]
    print("Building:", " ".join(cmd))
    rc = subprocess.call(cmd, cwd=HERE)
    if rc == 0:
        dist = HERE / "dist" / NAME
        shutil.copy2(HERE / "HOW TO USE.txt", dist / "HOW TO USE.txt")
        print(f"\nBuild OK. Zip this folder and share it:\n  {dist}")
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
