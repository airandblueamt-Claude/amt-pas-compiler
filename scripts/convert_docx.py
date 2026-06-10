"""
convert_docx.py — turn a .docx into PDF pages.

  * LibreOffice when available  -> faithful rendering of the Word document.
  * fallback: extract the paragraph text straight from the .docx XML (no external
    library) and lay it out on a branded AMT page, Arabic-aware.

The fallback is intentionally simple (text + basic paragraph breaks); it exists so
the submittal still assembles where LibreOffice is not installed. For final,
client-facing output, LibreOffice (or a pre-exported PDF) is recommended.
"""
from __future__ import annotations

import os
import re
import shutil
import zipfile
import tempfile
import subprocess

from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

import amt_common as A
from amt_common import (PAGE_W, PAGE_H, MARGIN_L, MARGIN_R, CONTENT_W,
                        F_EN, F_AR, BLACK, ar)
from render_tables import has_arabic, have_soffice


def _docx_paragraphs(path: str) -> list[str]:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    # split into paragraphs, pull the visible text runs from each
    paras = re.split(r"</w:p>", xml)
    out = []
    for p in paras:
        texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, flags=re.S)
        line = "".join(texts)
        line = (line.replace("&amp;", "&").replace("&lt;", "<")
                    .replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'"))
        out.append(line.strip())
    # collapse runs of blank lines
    cleaned, blank = [], False
    for ln in out:
        if ln == "":
            if not blank:
                cleaned.append("")
            blank = True
        else:
            cleaned.append(ln)
            blank = False
    return cleaned


def is_cover_docx(path: str, max_words: int = 12) -> bool:
    """True if the .docx is essentially just a heading/title (a redundant cover
    sheet), i.e. it has at most one non-empty paragraph with very few words.
    Real documents like the warranty letter have many paragraphs and fail this."""
    try:
        paras = [p for p in _docx_paragraphs(path) if p.strip()]
    except Exception:
        return False
    if len(paras) > 1:
        return False
    words = sum(len(p.split()) for p in paras)
    return words <= max_words


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


def render_with_soffice(docx_path: str, out_pdf: str, soffice: str) -> str:
    outdir = tempfile.mkdtemp(prefix="amt_lo_")
    profile = tempfile.mkdtemp(prefix="amt_loprof_")
    cmd = [soffice, "--headless", "--nologo", "--nofirststartwizard",
           f"-env:UserInstallation=file://{profile}",
           "--convert-to", "pdf", "--outdir", outdir, docx_path]
    subprocess.run(cmd, check=True, capture_output=True, timeout=180)
    base = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    produced = os.path.join(outdir, base)
    if not os.path.exists(produced):
        raise RuntimeError(f"LibreOffice did not produce {produced}")
    shutil.move(produced, out_pdf)
    shutil.rmtree(outdir, ignore_errors=True)
    shutil.rmtree(profile, ignore_errors=True)
    return out_pdf


def render_with_reportlab(docx_path: str, out_pdf: str, ref_no: str) -> str:
    A.register_fonts()
    paras = _docx_paragraphs(docx_path)
    c = canvas.Canvas(out_pdf, pagesize=(PAGE_W, PAGE_H))
    size = 11
    line_h = size + 5

    def new_page():
        A.page_chrome(c, ref_no)
        return PAGE_H - 175

    y = new_page()
    for para in paras:
        if para == "":
            y -= line_h
            if y < 150:
                c.showPage(); y = new_page()
            continue
        is_ar = has_arabic(para)
        font = F_AR if is_ar else F_EN
        for ln in _wrap(para, font, size, CONTENT_W, is_ar):
            if y < 150:
                c.showPage(); y = new_page()
            c.setFont(font, size)
            c.setFillColor(BLACK)
            disp = ar(ln) if is_ar else ln
            if is_ar:
                c.drawRightString(PAGE_W - MARGIN_R, y, disp)
            else:
                c.drawString(MARGIN_L, y, disp)
            y -= line_h
        y -= 3
    c.showPage()
    c.save()
    return out_pdf


def convert_docx(docx_path: str, out_pdf: str, ref_no: str, engine: str = "auto") -> tuple[str, str]:
    soffice = have_soffice()
    use_lo = (engine == "libreoffice") or (engine == "auto" and soffice)
    if use_lo:
        if not soffice:
            raise RuntimeError("render_engine=libreoffice but soffice not found.")
        try:
            return render_with_soffice(docx_path, out_pdf, soffice), "libreoffice"
        except Exception as e:
            if engine == "libreoffice":
                raise
            print(f"  ! LibreOffice failed on docx ({e}); using text fallback.")
    return render_with_reportlab(docx_path, out_pdf, ref_no), "reportlab-text"
