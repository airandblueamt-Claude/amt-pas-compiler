"""
amt_common.py — shared branding, geometry, Arabic shaping and page-chrome helpers
for the AMT PAS (Technical Submittal) compiler.

All AMT-authored pages (cover, table-of-contents, section dividers, placeholders)
are drawn on an A4 reportlab canvas using the helpers here so that the header logo,
the contact-banner footer, the reference line, fonts and colours stay identical
across every page and every submittal.
"""
from __future__ import annotations

import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import arabic_reshaper
from bidi.algorithm import get_display

# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS_DIR = os.path.join(SKILL_DIR, "assets")
FONTS_DIR = os.path.join(ASSETS_DIR, "fonts")

LOGO_PNG = os.path.join(ASSETS_DIR, "amt-logo.png")
BANNER_PNG = os.path.join(ASSETS_DIR, "amt-footer-banner.png")
SEAL_PNG = os.path.join(ASSETS_DIR, "amt-seal.png")

# --------------------------------------------------------------------------- #
# Page geometry (A4, points)
# --------------------------------------------------------------------------- #
PAGE_W, PAGE_H = A4                       # 595.32 x 841.92
MARGIN_L = 40
MARGIN_R = 40
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R

# Header logo box (top-left). Sized by a target HEIGHT so any logo aspect works
# (the AMT lockup is a stacked diamond + bilingual wordmark).
LOGO_X = 38
LOGO_H = 56                               # target header-logo height (points)
LOGO_TOP_GAP = 16                         # gap from page top to logo top
# Reserved top strip on generated table pages so the stamped logo never overlaps
# the table (logo gap + logo height + breathing room).
HEADER_BAND = LOGO_TOP_GAP + LOGO_H + 14

# Footer banner (full width, bottom)
BANNER_X = 16
BANNER_W = PAGE_W - 2 * BANNER_X
REF_LINE_GAP = 6                          # gap between banner top and the Ref line

# --------------------------------------------------------------------------- #
# Brand colours
# --------------------------------------------------------------------------- #
BLUE_HEADER = HexColor("#5B9BD5")         # table header fill
BLUE_BORDER = HexColor("#2E75B6")         # table grid lines
AMT_RED = HexColor("#C00000")
GREY_REF = HexColor("#404040")
WHITE = HexColor("#FFFFFF")
BLACK = HexColor("#000000")

# --------------------------------------------------------------------------- #
# Fonts
# --------------------------------------------------------------------------- #
F_EN = "Times-Roman"          # reportlab built-in (matches the serif body of the sample)
F_EN_B = "Times-Bold"
F_AR = "Amiri"                # registered below
F_AR_B = "Amiri-Bold"
F_SANS = "Helvetica"
F_SANS_B = "Helvetica-Bold"

_FONTS_READY = False


def register_fonts() -> None:
    """Register the Arabic TTFs with reportlab (idempotent)."""
    global _FONTS_READY
    if _FONTS_READY:
        return
    pdfmetrics.registerFont(TTFont(F_AR, os.path.join(FONTS_DIR, "Amiri-Regular.ttf")))
    pdfmetrics.registerFont(TTFont(F_AR_B, os.path.join(FONTS_DIR, "Amiri-Bold.ttf")))
    _FONTS_READY = True


# --------------------------------------------------------------------------- #
# Arabic shaping
# --------------------------------------------------------------------------- #
def ar(text: str) -> str:
    """Shape + reorder Arabic text for correct visual (RTL) rendering in a PDF."""
    if not text:
        return ""
    return get_display(arabic_reshaper.reshape(str(text)))


# --------------------------------------------------------------------------- #
# Image helpers
# --------------------------------------------------------------------------- #
def _img_aspect(path: str) -> float:
    """height / width of an image."""
    from PIL import Image
    with Image.open(path) as im:
        w, h = im.size
    return h / w


def draw_header_logo(c) -> float:
    """Draw the AMT logo top-left, sized to LOGO_H tall. Returns the logo bottom y."""
    aspect = _img_aspect(LOGO_PNG)          # h / w
    h = LOGO_H
    w = h / aspect
    y = PAGE_H - LOGO_TOP_GAP - h
    c.drawImage(LOGO_PNG, LOGO_X, y, width=w, height=h,
                preserveAspectRatio=True, mask="auto")
    return y


def draw_footer(c, ref_no: str) -> None:
    """Draw the 'Ref.: <ref>' line and the full-width contact banner at the bottom."""
    register_fonts()
    aspect = _img_aspect(BANNER_PNG)
    bh = BANNER_W * aspect
    c.drawImage(BANNER_PNG, BANNER_X, 0, width=BANNER_W, height=bh,
                preserveAspectRatio=True, mask="auto")
    # Reference line just above the banner, left aligned
    c.setFont(F_SANS, 8)
    c.setFillColor(GREY_REF)
    c.drawString(MARGIN_L + 20, bh + REF_LINE_GAP, f"Ref.: {ref_no}")
    c.setFillColor(BLACK)


def draw_seal(c, cx: float, cy: float, w: float = 120) -> None:
    """Draw the company oval seal centred on (cx, cy)."""
    aspect = _img_aspect(SEAL_PNG)
    h = w * aspect
    c.drawImage(SEAL_PNG, cx - w / 2, cy - h / 2, width=w, height=h,
                preserveAspectRatio=True, mask="auto")


def page_chrome(c, ref_no: str, with_logo: bool = True) -> float:
    """Apply the standard AMT chrome (logo + footer). Returns logo-bottom y."""
    logo_bottom = draw_header_logo(c) if with_logo else PAGE_H - LOGO_TOP_GAP
    draw_footer(c, ref_no)
    return logo_bottom


# --------------------------------------------------------------------------- #
# Small drawing utilities
# --------------------------------------------------------------------------- #
def text_center(c, x, y, text, font, size, color=BLACK):
    c.setFont(font, size)
    c.setFillColor(color)
    c.drawCentredString(x, y, text)
    c.setFillColor(BLACK)


def text_left(c, x, y, text, font, size, color=BLACK):
    c.setFont(font, size)
    c.setFillColor(color)
    c.drawString(x, y, text)
    c.setFillColor(BLACK)


def text_right(c, x, y, text, font, size, color=BLACK):
    c.setFont(font, size)
    c.setFillColor(color)
    c.drawRightString(x, y, text)
    c.setFillColor(BLACK)


# --------------------------------------------------------------------------- #
# Page-branding overlay (stamp the AMT logo/footer onto already-rendered pages)
# --------------------------------------------------------------------------- #
# Space (points) reserved at top/bottom of a GENERATED table page so the stamped
# header logo and footer banner never sit on top of the table. Kept in sync with
# the LibreOffice print margins set in render_tables._prepare_xlsx_for_print.
TABLE_TOP_RESERVE = 84
TABLE_BOTTOM_RESERVE = 92


def _stamp_header(c, pw, ph, ref_no):
    """AMT logo, top-LEFT, sized to LOGO_H — matches the cover/TOC/divider header so
    every page is consistent. No footer ref (the section divider already carries the
    reference), which also avoids a duplicate ref where the source sheet has its own.
    The page reserves a top band so this never overlaps the content."""
    aspect = _img_aspect(LOGO_PNG)          # h / w
    h = LOGO_H
    w = h / aspect
    c.drawImage(LOGO_PNG, LOGO_X, ph - LOGO_TOP_GAP - h, width=w, height=h,
                preserveAspectRatio=True, mask="auto")


def stamp_pdf(in_pdf: str, out_pdf: str, ref_no: str, mode: str = "header") -> str:
    """Overlay the AMT header logo (top-left) on every page of `in_pdf`. Page count
    and sizes are preserved."""
    from pypdf import PdfReader, PdfWriter
    register_fonts()
    reader = PdfReader(in_pdf)
    writer = PdfWriter()
    draw = _stamp_header
    for page in reader.pages:
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(pw, ph))
        draw(c, pw, ph, ref_no)
        c.showPage()
        c.save()
        buf.seek(0)
        page.merge_page(PdfReader(buf).pages[0])
        writer.add_page(page)
    with open(out_pdf, "wb") as fh:
        writer.write(fh)
    return out_pdf
