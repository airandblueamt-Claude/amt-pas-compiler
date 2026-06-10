#!/usr/bin/env python3
"""
extract_assets.py — ONE-OFF provenance helper.

Re-extract the AMT footer contact banner and the company seal from a completed
sample PAS PDF (they live on page 1). The header logo is supplied separately as a
clean high-resolution file and is NOT extracted here.

Usage:
    python3 extract_assets.py "<path to a sample PAS pdf>"

Writes amt-footer-banner.png and amt-seal.png into ../assets/. Run manually only
when the branding changes; the normal build does not call this.
"""
import os
import sys
from PIL import Image
from pypdf import PdfReader

ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")


def main(sample_pdf: str):
    reader = PdfReader(sample_pdf)
    page = reader.pages[0]
    saved = []
    for im in page.images:
        # Footer banner is the wide PNG; the seal is the small JP2/JPX image.
        w, h = im.image.size
        ratio = w / h
        if ratio > 3.5 and w > 800:          # wide banner
            out = os.path.join(ASSETS, "amt-footer-banner.png")
            im.image.convert("RGBA").save(out)
            saved.append(out)
        elif w < 300 and h < 300:            # small seal
            out = os.path.join(ASSETS, "amt-seal.png")
            im.image.convert("RGBA").save(out)
            saved.append(out)
    if not saved:
        print("No matching banner/seal images found on page 1.")
        return 1
    for s in saved:
        print("wrote", s)
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
