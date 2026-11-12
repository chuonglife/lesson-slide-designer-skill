#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Render every slide of a deck to a 1920x1080 PNG so you (and Claude) can look at the real result.

Usage:
  python render_slides.py <deck.pptx> [output folder]

Creates <output>/slide-NN.png for each slide and <output>/contact-sheet.png (3 columns) for a quick
overview. The default output folder is "_preview" next to the deck.
Needs LibreOffice (soffice) and pdftoppm (Poppler).
"""
import glob
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SOFFICE_CANDIDATES = [
    shutil.which("soffice"),
    shutil.which("libreoffice"),
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
]


def find_soffice():
    for c in SOFFICE_CANDIDATES:
        if c and os.path.isfile(c):
            return c
    raise SystemExit("LibreOffice (soffice) not found. Install it from https://www.libreoffice.org/")


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        sys.exit(2)
    pptx = os.path.abspath(sys.argv[1])
    out = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(os.path.dirname(pptx), "_preview")
    os.makedirs(out, exist_ok=True)
    for old in glob.glob(os.path.join(out, "slide-*.png")):
        os.remove(old)
    # A private LibreOffice profile per run: parallel runs on one shared profile fight over a lock and fail.
    prof = tempfile.mkdtemp(prefix="lo_profile_")
    uri = "file:///" + prof.replace("\\", "/").lstrip("/")
    try:
        subprocess.run([find_soffice(), f"-env:UserInstallation={uri}", "--headless", "--convert-to", "pdf",
                        "--outdir", out, pptx], check=True, capture_output=True)
    finally:
        shutil.rmtree(prof, ignore_errors=True)
    pdf = os.path.join(out, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise SystemExit("pdftoppm not found. Install Poppler (poppler-utils).")
    subprocess.run([pdftoppm, "-png", "-r", "144", "-scale-to-x", "1920", "-scale-to-y", "1080", pdf,
                    os.path.join(out, "slide")], check=True, capture_output=True)
    files = sorted(glob.glob(os.path.join(out, "slide-*.png")))
    renamed = []
    for i, f in enumerate(files, 1):
        new = os.path.join(out, f"slide-{i:02d}.png")
        if f != new:
            os.replace(f, new)
        renamed.append(new)
    cols, tw, th = 3, 640, 360
    rows = (len(renamed) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * tw, max(1, rows) * th), "white")
    for i, f in enumerate(renamed):
        im = Image.open(f).convert("RGB").resize((tw - 8, th - 8))
        sheet.paste(im, ((i % cols) * tw + 4, (i // cols) * th + 4))
    sheet_path = os.path.join(out, "contact-sheet.png")
    sheet.save(sheet_path)
    print(f"Rendered {len(renamed)} slides to {out}")
    print(f"Contact sheet: {sheet_path}")


if __name__ == "__main__":
    main()
