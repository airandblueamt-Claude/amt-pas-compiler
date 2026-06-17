"""
build_chrome.py — generate the AMT-authored pages (cover, table of contents,
section dividers, and 'to follow' placeholders) as single-page PDFs that match
the company's submittal template.

Each builder draws onto a reportlab canvas via the shared helpers in amt_common.
"""
from __future__ import annotations

import os
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

import amt_common as A
from amt_common import (PAGE_W, PAGE_H, MARGIN_L, MARGIN_R, CONTENT_W,
                        BLUE_HEADER, BLUE_BORDER, WHITE, BLACK, AMT_RED, GREY_REF,
                        F_EN, F_EN_B, F_AR, F_AR_B, F_SANS, F_SANS_B, ar)


def ref_display(cfg) -> str:
    return f"{cfg['ref_no']}-v.{cfg['version']}"


def mts_display(cfg) -> str:
    return f"{cfg.get('mts_ref_no', cfg['ref_no'])} v.{cfg['version']}"


# --------------------------------------------------------------------------- #
# Generic bordered grid table
# --------------------------------------------------------------------------- #
def draw_grid(c, x, y_top, col_widths, row_heights, cells):
    """Draw a bordered table. `cells[r][c]` is a dict or None:
        text, font, size, align('c'|'l'|'r'), arabic(bool), fill(color), color
    Returns the y of the table bottom edge."""
    total_w = sum(col_widths)
    y = y_top
    for r, rh in enumerate(row_heights):
        x_cur = x
        for cidx, cw in enumerate(col_widths):
            cell = cells[r][cidx] if cells[r][cidx] else {}
            fill = cell.get("fill")
            if fill is not None:
                c.setFillColor(fill)
                c.rect(x_cur, y - rh, cw, rh, stroke=0, fill=1)
            # border
            c.setStrokeColor(BLUE_BORDER)
            c.setLineWidth(0.75)
            c.rect(x_cur, y - rh, cw, rh, stroke=1, fill=0)
            # text (vertically centred, supports wrapping into up to 3 lines)
            txt = cell.get("text", "")
            if txt != "":
                font = cell.get("font", F_EN)
                size = cell.get("size", 9)
                color = cell.get("color", BLACK)
                align = cell.get("align", "c")
                is_ar = cell.get("arabic", False)
                lines = _wrap(c, str(txt), font, size, cw - 8, is_ar)
                lh = size + 2
                block_h = lh * len(lines)
                ty = y - rh / 2 + block_h / 2 - size
                c.setFillColor(color)
                c.setFont(font, size)
                for ln in lines:
                    disp = ar(ln) if is_ar else ln
                    if align == "l":
                        c.drawString(x_cur + 4, ty, disp)
                    elif align == "r":
                        c.drawRightString(x_cur + cw - 4, ty, disp)
                    else:
                        c.drawCentredString(x_cur + cw / 2, ty, disp)
                    ty -= lh
                c.setFillColor(BLACK)
            x_cur += cw
        y -= rh
    return y


def _wrap(c, text, font, size, max_w, is_ar):
    """Greedy word-wrap to fit max_w, max 4 lines."""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    words = text.split()
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
            if len(lines) == 3:
                break
    lines.append(cur)
    return lines[:4]


# --------------------------------------------------------------------------- #
# Cover page
# --------------------------------------------------------------------------- #
def _caption(c, x, y, text):
    """A small red small-caps label above a table block."""
    c.setFont(F_SANS_B, 8.5)
    c.setFillColor(AMT_RED)
    c.drawString(x, y, text.upper())
    c.setFillColor(BLACK)


def draw_cover(c, cfg):
    A.register_fonts()
    logo_bottom = A.page_chrome(c, ref_display(cfg))

    # --- header rule under the logo, with a submittal label on the right -----
    rule_y = logo_bottom - 14
    c.setStrokeColor(AMT_RED)
    c.setLineWidth(1.2)
    c.line(MARGIN_L, rule_y, PAGE_W - MARGIN_R, rule_y)
    c.setFont(F_SANS_B, 8.5)
    c.setFillColor(GREY_REF)
    c.drawRightString(PAGE_W - MARGIN_R, rule_y + 6, "TECHNICAL SUBMITTAL")
    c.setFillColor(BLACK)

    # --- title block (centred between the rule and the control tables) ------
    ty = rule_y - 100
    client_logo = cfg.get("client_logo")
    if client_logo and os.path.exists(client_logo):
        from PIL import Image
        with Image.open(client_logo) as im:
            asp = im.size[1] / im.size[0]
        lw = 96
        c.drawImage(client_logo, PAGE_W / 2 - lw / 2, ty - lw * asp,
                    width=lw, height=lw * asp, preserveAspectRatio=True, mask="auto")
        ty -= lw * asp + 24

    A.text_center(c, PAGE_W / 2, ty, ar(cfg["project_title_ar"]), F_AR_B, 16)
    A.text_center(c, PAGE_W / 2, ty - 25, cfg["project_title_en"], F_EN_B, 15)
    c.setStrokeColor(BLUE_BORDER)
    c.setLineWidth(0.6)
    c.line(PAGE_W / 2 - 95, ty - 38, PAGE_W / 2 + 95, ty - 38)
    A.text_center(c, PAGE_W / 2, ty - 58, f"Ref. #{mts_display(cfg)}", F_EN_B, 12, color=AMT_RED)
    A.text_center(c, PAGE_W / 2, ty - 92,
                  ar(cfg.get("company_name_ar", "شركة الأبعاد المترامية للتقنية")), F_AR_B, 19)

    # --- document-control tables (anchored to the lower half) ---------------
    rev = cfg["revision"]
    rcw = [150, 60, 95, 60, CONTENT_W - 365]
    rev_cells = [
        [_h("Ref"), _h("Version"), _h("Date"), _h("Author"), _h("Remarks")],
        [_d(cfg["ref_no"]), _d(cfg["version"]), _d(cfg["date"]),
         _dash(rev.get("author", "")), _d(rev.get("remarks", ""))],
    ]
    rev_label_y = PAGE_H - 380
    _caption(c, MARGIN_L, rev_label_y, "Revision")
    rev_bottom = draw_grid(c, MARGIN_L, rev_label_y - 9, rcw, [22, 24], rev_cells)

    so = cfg["signoff"]
    scw = [CONTENT_W / 3] * 3
    so_cells = [
        [_h("Ref BY"), _h("Prepared by"), _h("Approved by")],
        [_d(so["prepared_by"]["role_en"], F_EN_B), _d(so["checked_by"]["role_en"], F_EN_B),
         _d(so["approved_by"]["role_en"], F_EN_B)],
        [_dash(so["prepared_by"]["initials"]), _dash(so["checked_by"]["initials"]),
         _dash(so["approved_by"]["initials"])],
    ]
    so_label_y = rev_bottom - 24
    _caption(c, MARGIN_L, so_label_y, "Approval")
    so_bottom = draw_grid(c, MARGIN_L, so_label_y - 9, scw, [22, 24, 22], so_cells)

    # --- company seal, centred in the space above the footer ----------------
    footer_top = A.BANNER_W * A._img_aspect(A.BANNER_PNG) + 24
    seal_cy = (so_bottom + footer_top) / 2
    A.draw_seal(c, PAGE_W / 2, seal_cy, w=118)


# --------------------------------------------------------------------------- #
# Table of contents
# --------------------------------------------------------------------------- #
def draw_toc(c, cfg, page_map, sections, toc_title_en, toc_title_ar):
    A.register_fonts()
    A.page_chrome(c, ref_display(cfg))

    x = MARGIN_L
    col = [42, 250, CONTENT_W - 42 - 250 - 55, 55]   # S.No | EN | AR | Page.No
    # header row
    y = PAGE_H - 215
    draw_grid(c, x, y, col, [26], [[
        {"text": "S.No", "font": F_EN_B, "size": 10, "color": WHITE, "fill": BLUE_HEADER},
        {"text": toc_title_en, "font": F_EN_B, "size": 10, "align": "l", "color": WHITE, "fill": BLUE_HEADER},
        {"text": toc_title_ar, "font": F_AR_B, "size": 11, "align": "r", "arabic": True,
         "color": WHITE, "fill": BLUE_HEADER},
        {"text": "Page.No", "font": F_EN_B, "size": 10, "color": WHITE, "fill": BLUE_HEADER},
    ]])
    y -= 26
    # adapt row height to the number of sections so they always fit the page
    avail = y - 250
    rh = max(20, min(34, avail / max(len(sections), 1)))
    total_w = sum(col)
    rects = []   # (section_no, x0, y0, x1, y1) for clickable links
    for spec in sections:
        row = [[
            {"text": str(spec["no"]), "font": F_EN_B, "size": 10},
            {"text": spec["en"], "font": F_EN_B, "size": 10, "align": "l"},
            {"text": spec["ar"], "font": F_AR_B, "size": 11, "align": "r", "arabic": True},
            {"text": f"{page_map.get(spec['no'], ''):02d}" if spec["no"] in page_map else "",
             "font": F_EN_B, "size": 10},
        ]]
        top = y
        y = draw_grid(c, x, y, col, [rh], row)
        if spec["no"] in page_map:
            rects.append((spec["no"], x, y, x + total_w, top))

    A.draw_seal(c, PAGE_W / 2 + 90, 235, w=110)
    return rects


# --------------------------------------------------------------------------- #
# Section divider
# --------------------------------------------------------------------------- #
def _fit_lines(text, font, base_size, max_w, is_ar, min_size=9):
    """Return (size, [lines]) so the text fits within max_w: first try shrinking
    to one line down to min_size, then wrap to two lines at min_size."""
    disp = ar(text) if is_ar else text
    w = stringWidth(disp, font, base_size)
    if w <= max_w:
        return base_size, [text]
    size = base_size * max_w / w
    if size >= min_size:
        return size, [text]
    # still too wide at min_size -> wrap into <=2 lines at min_size
    words = str(text).split()
    lines, cur = [], words[0] if words else ""
    for word in words[1:]:
        trial = cur + " " + word
        td = ar(trial) if is_ar else trial
        if stringWidth(td, font, min_size) <= max_w:
            cur = trial
        else:
            lines.append(cur)
            cur = word
    lines.append(cur)
    return min_size, lines[:2]


def draw_section_title_row(c, section, y):
    """Draw 'N-  <English>' on the left and the Arabic title on the right, each
    fitted to its own half so the two can never overlap in the middle."""
    x_left = MARGIN_L + 35
    x_right = PAGE_W - MARGIN_R - 20
    centre = PAGE_W / 2
    gap = 16
    max_en = (centre - gap) - x_left
    max_ar = x_right - (centre + gap)

    en_text = f"{section['no']}-    {section['en']}"
    s_en, en_lines = _fit_lines(en_text, F_EN_B, 14, max_en, is_ar=False)
    s_ar, ar_lines = _fit_lines(section["ar"], F_AR_B, 15, max_ar, is_ar=True)

    lh_en = s_en + 3
    yy = y
    c.setFont(F_EN_B, s_en)
    c.setFillColor(BLACK)
    for ln in en_lines:
        c.drawString(x_left, yy, ln)
        yy -= lh_en

    lh_ar = s_ar + 3
    yy = y
    c.setFont(F_AR_B, s_ar)
    for ln in ar_lines:
        c.drawRightString(x_right, yy, ar(ln))
        yy -= lh_ar
    c.setFillColor(BLACK)


def draw_divider(c, cfg, section):
    A.register_fonts()
    logo_bottom = A.page_chrome(c, ref_display(cfg))
    draw_section_title_row(c, section, logo_bottom - 40)
    # Seal centred lower
    A.draw_seal(c, PAGE_W / 2 + 90, PAGE_H / 2 - 30, w=120)


def draw_placeholder(c, cfg, section, note):
    A.register_fonts()
    logo_bottom = A.page_chrome(c, ref_display(cfg))
    draw_section_title_row(c, section, logo_bottom - 40)
    A.text_center(c, PAGE_W / 2, PAGE_H / 2 + 20, note, F_EN_B, 13, color=GREY_REF)


# --------------------------------------------------------------------------- #
# Cell shortcuts
# --------------------------------------------------------------------------- #
def _h(text):
    return {"text": text, "font": F_EN_B, "size": 9, "color": WHITE, "fill": BLUE_HEADER}


def _d(text, font=F_EN, size=9):
    return {"text": text, "font": font, "size": size}


def _dash(text, font=F_EN, size=9):
    """Like _d, but an empty value renders as a faint centred dash so a blank
    Author / initials cell looks intentional rather than unfinished."""
    if text is None or not str(text).strip():
        return {"text": "—", "font": font, "size": size, "color": GREY_REF}
    return {"text": text, "font": font, "size": size}


# --------------------------------------------------------------------------- #
# Page emitters (one single-page PDF each)
# --------------------------------------------------------------------------- #
def _emit(path, draw_fn):
    c = canvas.Canvas(path, pagesize=(PAGE_W, PAGE_H))
    draw_fn(c)
    c.showPage()
    c.save()
    return path


def cover_pdf(cfg, out):    return _emit(out, lambda c: draw_cover(c, cfg))
def toc_pdf(cfg, page_map, sections, toc_en, toc_ar, out):
    """Emit the TOC page and return the clickable row rectangles
    [(section_no, x0, y0, x1, y1), …] in PDF coordinates."""
    c = canvas.Canvas(out, pagesize=(PAGE_W, PAGE_H))
    rects = draw_toc(c, cfg, page_map, sections, toc_en, toc_ar)
    c.showPage()
    c.save()
    return rects
def divider_pdf(cfg, section, out): return _emit(out, lambda c: draw_divider(c, cfg, section))
def placeholder_pdf(cfg, section, note, out):
    return _emit(out, lambda c: draw_placeholder(c, cfg, section, note))


if __name__ == "__main__":
    # smoke test
    import json, sys
    import pas_spec as SPEC
    cfg = json.load(open(sys.argv[1]))
    tpl = SPEC.resolve_template(cfg)
    secs = tpl["sections"]
    os.makedirs("/tmp/chrome", exist_ok=True)
    cover_pdf(cfg, "/tmp/chrome/cover.pdf")
    pm = {s["no"]: 3 + 2 * i for i, s in enumerate(secs)}
    toc_pdf(cfg, pm, secs, tpl["toc_title_en"], tpl["toc_title_ar"], "/tmp/chrome/toc.pdf")
    divider_pdf(cfg, secs[0], "/tmp/chrome/div1.pdf")
    placeholder_pdf(cfg, secs[-1], "Certificate to follow", "/tmp/chrome/ph6.pdf")
    print("chrome smoke test written to /tmp/chrome")
