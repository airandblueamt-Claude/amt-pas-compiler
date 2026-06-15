"""
assemble.py — build the full PAS PDF from a discovered manifest + config.

Two-pass build:
  pass 1  render each section's content (tables -> PDF, docx -> PDF, PDFs as-is)
          and measure page counts.
  pass 2  knowing the page counts, compute the physical page where each section
          starts, draw the cover + TOC (with real page numbers) + dividers /
          placeholders, then merge everything in order.

Only one artifact is written: config["output_pdf"].
"""
from __future__ import annotations

import os
import tempfile

from pypdf import PdfReader, PdfWriter

import build_chrome as CH
import render_tables as RT
import convert_docx as CD
import normalize as NORM
import amt_common as A
from pas_spec import TABLE, APPEND


def _pdf_pages(path: str) -> int:
    return len(PdfReader(path).pages)


def _render_section_content(sec, cfg, ref_disp, work, engines):
    """Produce a list of PDF file paths for a section's CONTENT (no divider).
    Returns (list_of_pdfs, is_placeholder)."""
    no = sec["no"]
    if sec["kind"] == TABLE:
        out = os.path.join(work, f"sec{no}_table.pdf")
        _, eng = RT.render_table(sec["xlsx"], out, sec["en"], ref_disp,
                                 engine=cfg.get("render_engine", "auto"))
        engines.add(eng)
        return [out], False

    # APPEND: real Word content (e.g. the warranty letter) is kept and placed
    # before the PDFs. A heading-only ".docx" (e.g. a "Shop Drawings/Single Line
    # Diagram" cover sheet) is redundant with our generated section divider and
    # carries stale numbering/refs, so it is skipped.
    converted = []
    for i, d in enumerate(sec["docx"]):
        if cfg.get("drop_cover_docx", True) and CD.is_cover_docx(d):
            print(f"  · §{no}: skipping cover-only doc '{os.path.basename(d)}' "
                  f"(redundant with the section divider).")
            continue
        out = os.path.join(work, f"sec{no}_docx{i}.pdf")
        _, eng = CD.convert_docx(d, out, ref_disp, engine=cfg.get("render_engine", "auto"))
        engines.add(eng)
        converted.append(out)
    parts = converted + list(sec["pdfs"])
    if not parts:
        return [], True   # empty -> caller decides placeholder/skip

    # Fit oversized / off-size pages (e.g. A1 layout drawings) onto A4 so the
    # whole submittal is one consistent paper size. With drawing_fit="native"
    # (matching the sample document) drawings are appended at original size.
    mode = cfg.get("drawing_fit", "native")
    if cfg.get("normalize_appended", True) and mode != "native":
        normed = []
        for j, p in enumerate(parts):
            try:
                if NORM.needs_normalising(p):
                    out = os.path.join(work, f"sec{no}_norm{j}.pdf")
                    NORM.normalize_to_a4(p, out, mode=mode)
                    normed.append(out)
                else:
                    normed.append(p)
            except Exception as e:
                print(f"  ! normalise failed for {os.path.basename(p)} ({e}); using original size.")
                normed.append(p)
        parts = normed

    # Brand third-party datasheet/diagram pages with a small AMT logo + ref line
    # so every page of the submittal carries AMT identity (kept unobtrusive).
    if cfg.get("brand_appended", True):
        stamped = []
        for k, p in enumerate(parts):
            sp = os.path.join(work, f"sec{no}_brand{k}.pdf")
            try:
                A.stamp_pdf(p, sp, ref_disp, mode="appended")
                stamped.append(sp)
            except Exception as e:
                print(f"  ! branding failed for {os.path.basename(p)} ({e}); using unbranded.")
                stamped.append(p)
        parts = stamped
    return parts, False


def build(manifest: dict, cfg: dict, template: dict | None = None, log=print) -> dict:
    ref_disp = CH.ref_display(cfg)
    work = tempfile.mkdtemp(prefix="amt_pas_")
    engines = set()
    mode = cfg.get("missing_section_mode", "placeholder")
    if template is None:
        import pas_spec as SPEC
        template = SPEC.resolve_template(cfg)

    # ---- PASS 1: render content + measure ---------------------------------
    log("Pass 1/2 — rendering section content …")
    rendered = {}   # no -> dict(content=[paths], pages=int, placeholder=bool)
    for sec in manifest["sections"]:
        parts, is_empty = _render_section_content(sec, cfg, ref_disp, work, engines)
        if is_empty:
            rendered[sec["no"]] = dict(content=[], pages=0, placeholder=True)
            log(f"  · §{sec['no']} {sec['en']}: empty")
        else:
            pages = sum(_pdf_pages(p) for p in parts)
            rendered[sec["no"]] = dict(content=parts, pages=pages, placeholder=False)
            log(f"  · §{sec['no']} {sec['en']}: {pages} page(s)")

    # ---- compute pagination ----------------------------------------------
    # page 1 cover, page 2 TOC, then per section: 1 divider/placeholder + content
    page_map = {}
    running = 2  # cover + toc
    for sec in manifest["sections"]:
        info = rendered[sec["no"]]
        page_map[sec["no"]] = running + 1            # the section's first page
        if info["placeholder"] and mode == "skip":
            running += 1                             # divider only
        elif info["placeholder"]:
            running += 1                             # placeholder page acts as divider
        else:
            running += 1 + info["pages"]             # divider + content
    total_pages = running

    # ---- PASS 2: build chrome --------------------------------------------
    log("Pass 2/2 — building cover, contents and dividers …")
    cover = os.path.join(work, "cover.pdf")
    toc = os.path.join(work, "toc.pdf")
    CH.cover_pdf(cfg, cover)
    toc_rects = CH.toc_pdf(cfg, page_map, template["sections"],
                           template["toc_title_en"], template["toc_title_ar"], toc)

    order = [cover, toc]
    for sec in manifest["sections"]:
        info = rendered[sec["no"]]
        if info["placeholder"]:
            if mode == "skip":
                div = os.path.join(work, f"div{sec['no']}.pdf")
                CH.divider_pdf(cfg, sec, div)
                order.append(div)
            else:
                ph = os.path.join(work, f"ph{sec['no']}.pdf")
                note = cfg.get("placeholder_note", "To be submitted / Certificate to follow")
                CH.placeholder_pdf(cfg, sec, note, ph)
                order.append(ph)
        else:
            div = os.path.join(work, f"div{sec['no']}.pdf")
            CH.divider_pdf(cfg, sec, div)
            order.append(div)
            order.extend(info["content"])

    # ---- merge ------------------------------------------------------------
    log("Merging …")
    writer = PdfWriter()
    for p in order:
        writer.append(p)

    # Make the Table of Contents clickable: each row jumps to that section's
    # first page (cover=index 0, TOC=index 1). page_map is 1-based.
    try:
        from pypdf.annotations import Link
        from pypdf.generic import Fit
        for (no, x0, y0, x1, y1) in toc_rects:
            tgt = page_map.get(no)
            if not tgt:
                continue
            writer.add_annotation(
                page_number=1,
                annotation=Link(rect=(x0, y0, x1, y1),
                                target_page_index=tgt - 1, fit=Fit.fit()),
            )
    except Exception as e:
        print(f"  ! could not add clickable TOC links ({e}); TOC text page numbers still apply.")

    out_pdf = cfg["output_pdf"]
    os.makedirs(os.path.dirname(os.path.abspath(out_pdf)), exist_ok=True)
    with open(out_pdf, "wb") as fh:
        writer.write(fh)

    # ---- QA ---------------------------------------------------------------
    actual = _pdf_pages(out_pdf)
    qa = dict(expected_pages=total_pages, actual_pages=actual,
              page_map=page_map, engines=sorted(engines), output=out_pdf,
              consistent=(actual == total_pages))
    return qa
