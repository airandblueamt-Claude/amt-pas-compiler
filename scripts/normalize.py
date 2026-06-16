"""
normalize.py — fit appended source pages (datasheets, certificates, drawings) onto
A4 so the whole submittal is one consistent paper size.

Large engineering drawings (e.g. A1 landscape single-line diagrams) are scaled down
to A4 with their aspect ratio preserved and centred. Orientation is preserved by
default: a landscape sheet lands on A4-landscape, a portrait sheet on A4-portrait
(so drawings stay as large and readable as possible). Pages already ~A4 are left
effectively unchanged (never enlarged).

Uses PyMuPDF (fitz), which respects each source page's /Rotate flag.
"""
from __future__ import annotations

import fitz   # PyMuPDF

A4_W, A4_H = 595.32, 841.92
TOL = 2.0     # points; treat as "already A4" within this tolerance


def _target_size(sw: float, sh: float, mode: str) -> tuple[float, float]:
    if mode == "portrait":
        return A4_W, A4_H
    if mode == "landscape":
        return A4_H, A4_W
    # auto: match the source orientation
    return (A4_H, A4_W) if sw > sh else (A4_W, A4_H)


def needs_normalising(path: str) -> bool:
    """True if any page is not already (within tolerance) an A4 sheet."""
    doc = fitz.open(path)
    try:
        for pg in doc:
            w, h = pg.rect.width, pg.rect.height
            a4 = (abs(w - A4_W) < TOL and abs(h - A4_H) < TOL) or \
                 (abs(w - A4_H) < TOL and abs(h - A4_W) < TOL)
            if not a4:
                return True
        return False
    finally:
        doc.close()


def normalize_to_a4(src: str, out: str, mode: str = "auto") -> str:
    """Write `out` where every page of `src` is placed on an A4 sheet, scaled to
    fit (never enlarged) and centred. Returns `out`."""
    src_doc = fitz.open(src)
    new = fitz.open()
    try:
        for pno in range(src_doc.page_count):
            sp = src_doc[pno]
            sw, sh = sp.rect.width, sp.rect.height
            tw, th = _target_size(sw, sh, mode)
            scale = min(tw / sw, th / sh, 1.0)
            nw, nh = sw * scale, sh * scale
            x0, y0 = (tw - nw) / 2, (th - nh) / 2
            page = new.new_page(width=tw, height=th)
            page.show_pdf_page(fitz.Rect(x0, y0, x0 + nw, y0 + nh), src_doc, pno)
        new.save(out, garbage=3, deflate=True)
        return out
    finally:
        new.close()
        src_doc.close()


def to_a4_portrait(src: str, out: str, side: float = 0.0,
                   top_reserve: float = 0.0) -> str:
    """Make every page of `src` a uniform A4 *portrait* sheet.

      * with no reserve, a page that is already A4 portrait is copied through
        untouched (full size); other sizes are fitted; landscape pages rotate 90°.
      * with `top_reserve` > 0 (used for AMT-generated tables), the content is
        always placed into the area BELOW the reserved top band — guaranteeing a
        clear strip for the stamped logo, independent of the source page's margins.
    Aspect ratio is preserved and content centred. Returns `out`.
    """
    src_doc = fitz.open(src)
    new = fitz.open()
    try:
        for pno in range(src_doc.page_count):
            sp = src_doc[pno]
            sw, sh = sp.rect.width, sp.rect.height
            if top_reserve <= 0 and abs(sw - A4_W) < TOL and abs(sh - A4_H) < TOL:
                new.insert_pdf(src_doc, from_page=pno, to_page=pno)   # keep as-is
                continue
            page = new.new_page(width=A4_W, height=A4_H)
            rect = fitz.Rect(side, side + top_reserve, A4_W - side, A4_H - side)
            rot = 90 if sw > sh else 0
            page.show_pdf_page(rect, src_doc, pno, keep_proportion=True, rotate=rot)
        new.save(out, garbage=3, deflate=True)
        return out
    finally:
        new.close()
        src_doc.close()


if __name__ == "__main__":
    import sys
    print(normalize_to_a4(sys.argv[1], sys.argv[2],
                          sys.argv[3] if len(sys.argv) > 3 else "auto"))
