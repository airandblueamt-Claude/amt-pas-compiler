"""
stamp.py — overlay ONLY the AMT logo header (top-left) onto each page of an
already-rendered BOQ table PDF.

This is intentionally lighter than full page chrome: it adds the company logo so
the §1–4 table pages carry AMT identity, but no footer banner — keeping the table
itself faithful and minimising reserved space / overlap risk. The table is rendered
with a reserved top margin (amt_common.logo_header_reserve) so the logo lands in
empty space and never covers a row.
"""
from __future__ import annotations

import io

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

import amt_common as A


def _logo_overlay(w: float, h: float) -> "PdfReader":
    """A one-page overlay (sized w×h) carrying only the AMT logo, top-left."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    lh = A.LOGO_H
    lw = lh / A._img_aspect(A.LOGO_PNG)
    c.drawImage(A.LOGO_PNG, A.LOGO_X, h - A.LOGO_TOP_GAP - lh,
                width=lw, height=lh, preserveAspectRatio=True, mask="auto")
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def stamp_logo(in_pdf: str, out_pdf: str) -> str:
    """Overlay the AMT logo on each page of in_pdf -> out_pdf. One overlay is built
    per distinct page size and reused."""
    reader = PdfReader(in_pdf)
    writer = PdfWriter()
    cache: dict = {}
    for page in reader.pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        key = (round(w), round(h))
        if key not in cache:
            cache[key] = _logo_overlay(w, h)
        page.merge_page(cache[key])
        writer.add_page(page)
    with open(out_pdf, "wb") as fh:
        writer.write(fh)
    return out_pdf
