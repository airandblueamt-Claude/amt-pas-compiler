"""
render_tables.py — render a BOQ spreadsheet (.xlsx) to branded PDF table pages.

Two engines:
  * LibreOffice ("soffice --headless --convert-to pdf"): pixel-faithful to the
    company's Excel formatting (Arabic RTL, merged cells). Used when available.
  * reportlab fallback: reads the sheet with openpyxl and re-typesets a clean,
    AMT-styled bilingual grid. Self-contained (no external binary needed).

Both return a path to a PDF whose pages contain only the table (no logo/banner),
matching how the sample embeds the rendered Excel sheets.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile

import openpyxl
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

import amt_common as A
from amt_common import (PAGE_W, PAGE_H, MARGIN_L, MARGIN_R, CONTENT_W,
                        BLUE_HEADER, BLUE_BORDER, WHITE, BLACK, GREY_REF,
                        F_EN, F_EN_B, F_AR, F_AR_B, ar)

_ARABIC_RE = re.compile(r"[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]")


def has_arabic(s: str) -> bool:
    return bool(_ARABIC_RE.search(s or ""))


def have_soffice() -> str | None:
    """Locate a LibreOffice binary. Checks $AMT_SOFFICE, then PATH, then common
    install locations including a user-extracted AppImage under ~/.local/opt."""
    import glob
    env = os.environ.get("AMT_SOFFICE")
    if env and os.path.exists(env):
        return env
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    patterns = [
        os.path.expanduser("~/.local/opt/squashfs-root/opt/libreoffice*/program/soffice"),
        os.path.expanduser("~/.local/opt/libreoffice*/program/soffice"),
        "/opt/libreoffice*/program/soffice",
        "/usr/lib/libreoffice/program/soffice",
    ]
    for pat in patterns:
        hits = sorted(glob.glob(pat))
        if hits:
            return hits[-1]
    return None


# --------------------------------------------------------------------------- #
# Fit-to-page preparation (so wide BOQ sheets scale onto one A4 width, centred,
# instead of overflowing across several pages)
# --------------------------------------------------------------------------- #
def _prepare_xlsx_for_print(xlsx_path: str, branded: bool = True) -> str:
    """Return a temp copy of the workbook set to A4 LANDSCAPE, fit-all-columns-to-
    one-page-wide and centred. Landscape gives wide / Arabic BOQ tables enough room
    so cells don't shrink into each other. When `branded`, top/bottom margins leave
    clearance for the stamped AMT logo + footer banner; otherwise normal margins."""
    from openpyxl.worksheet.properties import PageSetupProperties
    from openpyxl.styles import Alignment
    wb = openpyxl.load_workbook(xlsx_path)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"     # (#2) landscape tables
        ws.page_setup.paperSize = 9                 # A4
        ws.page_setup.fitToWidth = 1                # all columns on one page wide
        ws.page_setup.fitToHeight = 0               # as many pages tall as needed
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_options.horizontalCentered = True  # centre the table (#4)
        ws.print_options.verticalCentered = True
        # margins (inches). When branded, reserve room so the stamped header/footer
        # never overlap the table (#3); otherwise use tidy default margins.
        ws.page_margins.top = 1.15 if branded else 0.5
        ws.page_margins.bottom = 1.20 if branded else 0.5
        ws.page_margins.left = 0.45
        ws.page_margins.right = 0.45
        ws.page_margins.header = 0.2
        ws.page_margins.footer = 0.2

        # Wrap every cell and DROP manual row heights so each row auto-grows to fit
        # its (often multi-line Arabic) content. Fixed/too-small row heights are the
        # cause of long descriptions overflowing and overlapping neighbouring rows.
        for row in ws.iter_rows():
            for cell in row:
                if cell.value in (None, ""):
                    continue
                try:
                    al = cell.alignment
                    cell.alignment = Alignment(
                        horizontal=al.horizontal,
                        vertical="center",
                        wrap_text=True,
                        text_rotation=al.text_rotation or 0,
                        reading_order=al.reading_order or 0,
                        indent=al.indent or 0,
                    )
                except (AttributeError, TypeError):
                    pass  # merged-cell phantoms / styles we can't touch
        for rd in list(ws.row_dimensions.values()):
            rd.height = None   # auto-fit row height to the wrapped content
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", prefix="amt_fit_")
    os.close(fd)
    wb.save(tmp)
    return tmp


# --------------------------------------------------------------------------- #
# LibreOffice engine
# --------------------------------------------------------------------------- #
def render_with_soffice(xlsx_path: str, out_pdf: str, soffice: str) -> str:
    outdir = tempfile.mkdtemp(prefix="amt_lo_")
    # A dedicated profile dir avoids clashes with a running LibreOffice instance.
    profile = tempfile.mkdtemp(prefix="amt_loprof_")
    cmd = [soffice, "--headless", "--nologo", "--nofirststartwizard",
           f"-env:UserInstallation=file://{profile}",
           "--convert-to", "pdf", "--outdir", outdir, xlsx_path]
    subprocess.run(cmd, check=True, capture_output=True, timeout=180)
    base = os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf"
    produced = os.path.join(outdir, base)
    if not os.path.exists(produced):
        raise RuntimeError(f"LibreOffice did not produce {produced}")
    shutil.move(produced, out_pdf)
    shutil.rmtree(outdir, ignore_errors=True)
    shutil.rmtree(profile, ignore_errors=True)
    return out_pdf


# --------------------------------------------------------------------------- #
# reportlab fallback engine
# --------------------------------------------------------------------------- #
def _cell_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _sheet_matrix(ws):
    """Return (rows, ncols) trimmed to the used area, with merged cells expanded
    so the top-left value shows and a merge-span map is available."""
    rows = [[_cell_text(c) for c in r] for r in ws.iter_rows(values_only=True)]
    while rows and all(c == "" for c in rows[-1]):
        rows.pop()
    if not rows:
        return [], 0, {}
    ncols = max((max((i + 1 for i, c in enumerate(r) if c != ""), default=0) for r in rows),
                default=0)
    rows = [r[:ncols] + [""] * (ncols - len(r)) for r in rows]
    # merged spans -> {(r,c): (rowspan, colspan)} 0-indexed within used area
    spans = {}
    covered = set()
    for mc in ws.merged_cells.ranges:
        r0, c0, r1, c1 = mc.min_row - 1, mc.min_col - 1, mc.max_row - 1, mc.max_col - 1
        if r0 >= len(rows) or c0 >= ncols:
            continue
        spans[(r0, c0)] = (r1 - r0 + 1, c1 - c0 + 1)
        for rr in range(r0, min(r1 + 1, len(rows))):
            for cc in range(c0, min(c1 + 1, ncols)):
                if (rr, cc) != (r0, c0):
                    covered.add((rr, cc))
    return rows, ncols, spans, covered


def _col_widths(ws, ncols, content_w=CONTENT_W):
    """Approximate column widths (points) from Excel column dimensions, scaled
    to the printable content width."""
    widths = []
    for i in range(ncols):
        letter = openpyxl.utils.get_column_letter(i + 1)
        dim = ws.column_dimensions.get(letter)
        w = dim.width if dim and dim.width else 10
        widths.append(max(w, 4))
    total = sum(widths)
    scale = content_w / total
    return [w * scale for w in widths]


def _wrap(text, font, size, max_w, is_ar):
    words = str(text).split()
    if not words:
        return [""]
    lines, cur = [], words[0]
    for w in words[1:]:
        trial = cur + " " + w
        disp = ar(trial) if is_ar else trial
        if stringWidth(disp, font, size) <= max_w:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    lines.append(cur)
    return lines


def render_with_reportlab(xlsx_path: str, out_pdf: str, title: str, ref_no: str,
                          branded: bool = True) -> str:
    A.register_fonts()
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    # pick the first non-empty, prefer a non-Arabic-named sheet (English layout)
    ws = wb.active
    for cand in wb.worksheets:
        if cand.max_row and cand.max_column:
            ws = cand
            break
    rows, ncols, spans, covered = _sheet_matrix(ws)
    if ncols == 0:
        rows, ncols, spans, covered = [["(empty sheet)"]], 1, {}, set()

    # A4 landscape (#2) so wide/Arabic tables have room and don't overlap
    page_w, page_h = PAGE_H, PAGE_W
    content_w = page_w - MARGIN_L - MARGIN_R
    colw = _col_widths(ws, ncols, content_w)
    size = 8
    pad = 3
    line_h = size + 2

    c = canvas.Canvas(out_pdf, pagesize=(page_w, page_h))
    # Reserve top/bottom only when branding will be stamped, so chrome never
    # overlaps the table (#3); unbranded tables use the whole page.
    top_reserve = A.TABLE_TOP_RESERVE if branded else 28
    bottom_reserve = A.TABLE_BOTTOM_RESERVE if branded else 28
    top_y = page_h - top_reserve - 6
    bottom_limit = bottom_reserve + 6
    x_left = (page_w - sum(colw)) / 2   # centre the table horizontally (#4)
    page_no = 1

    def row_height(r):
        h = line_h + 2 * pad
        for cidx in range(ncols):
            if (r, cidx) in covered:
                continue
            txt = rows[r][cidx]
            if not txt:
                continue
            span = spans.get((r, cidx), (1, 1))
            w = sum(colw[cidx:cidx + span[1]]) - 2 * pad
            is_ar = has_arabic(txt)
            font = (F_AR if is_ar else F_EN)
            nlines = len(_wrap(txt, font, size, w, is_ar))
            h = max(h, nlines * line_h + 2 * pad)
        return min(h, page_h - 140)   # never taller than a page

    y = top_y
    is_header_row = True
    for r in range(len(rows)):
        rh = row_height(r)
        if y - rh < bottom_limit:
            c.showPage()
            page_no += 1
            y = top_y
        x = x_left
        for cidx in range(ncols):
            if (r, cidx) in covered:
                x += colw[cidx]
                continue
            span = spans.get((r, cidx), (1, 1))
            cw = sum(colw[cidx:cidx + span[1]])
            # header band for the first row
            header_band = is_header_row
            if header_band:
                c.setFillColor(BLUE_HEADER)
                c.rect(x, y - rh, cw, rh, stroke=0, fill=1)
            c.setStrokeColor(BLUE_BORDER)
            c.setLineWidth(0.5)
            c.rect(x, y - rh, cw, rh, stroke=1, fill=0)
            txt = rows[r][cidx]
            if txt:
                is_ar = has_arabic(txt)
                font = (F_AR_B if header_band else F_AR) if is_ar else (F_EN_B if header_band else F_EN)
                color = WHITE if header_band else BLACK
                lines = _wrap(txt, font, size, cw - 2 * pad, is_ar)
                c.setFont(font, size)
                c.setFillColor(color)
                ty = y - pad - size
                for ln in lines:
                    disp = ar(ln) if is_ar else ln
                    if is_ar:
                        c.drawRightString(x + cw - pad, ty, disp)
                    else:
                        c.drawString(x + pad, ty, disp)
                    ty -= line_h
                c.setFillColor(BLACK)
            x += cw
        y -= rh
        is_header_row = False

    c.showPage()
    c.save()
    return out_pdf


# --------------------------------------------------------------------------- #
# Dispatcher
# --------------------------------------------------------------------------- #
def render_table(xlsx_path: str, out_pdf: str, title: str, ref_no: str,
                 engine: str = "auto", brand: bool = True) -> tuple[str, str]:
    """Render xlsx -> A4 landscape PDF table pages. Returns (out_pdf, engine_used).

    The sheet is scaled to fit one page wide and centred. When brand=True the AMT
    logo + footer banner are stamped on each page (sections like the Tender BOQ,
    Catalogue and Warranty pass brand=False so no AMT logo appears on them)."""
    soffice = have_soffice()
    use_lo = (engine == "libreoffice") or (engine == "auto" and soffice)

    # render the bare table to a temp file, then stamp branding onto it
    if brand:
        fd, raw = tempfile.mkstemp(suffix=".pdf", prefix="amt_tbl_")
        os.close(fd)
    else:
        raw = out_pdf

    eng = "reportlab"
    if use_lo:
        if not soffice:
            raise RuntimeError("render_engine=libreoffice but soffice not found on PATH.")
        prepared = None
        try:
            prepared = _prepare_xlsx_for_print(xlsx_path, branded=brand)
            render_with_soffice(prepared, raw, soffice)
            eng = "libreoffice"
        except Exception as e:
            if engine == "libreoffice":
                raise
            print(f"  ! LibreOffice failed ({e}); using reportlab fallback.")
            render_with_reportlab(xlsx_path, raw, title, ref_no, branded=brand)
        finally:
            if prepared and os.path.exists(prepared):
                os.remove(prepared)
    else:
        render_with_reportlab(xlsx_path, raw, title, ref_no, branded=brand)

    if brand:
        A.stamp_pdf(raw, out_pdf, ref_no, mode="table")
        if os.path.exists(raw):
            os.remove(raw)
    return out_pdf, eng


if __name__ == "__main__":
    import sys
    out, eng = render_table(sys.argv[1], "/tmp/table_test.pdf", "Tender BoQ", "2506038-TCS-010-v.00")
    print("engine:", eng, "->", out)
