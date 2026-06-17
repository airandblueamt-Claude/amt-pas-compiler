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
from pas_spec import TABLE, APPEND


def _pdf_pages(path: str) -> int:
    return len(PdfReader(path).pages)


def _render_section_content(sec, cfg, ref_disp, work, engines):
    """Render whatever a section contains — spreadsheets -> faithful table PDFs,
    Word -> PDF, PDFs as-is — then place every page on a uniform A4 portrait sheet.
    AMT-authored sections (branded_sections, default §2/§3/§4 — Material Sheet,
    Traceability, Material Selection) carry the AMT logo on their table pages; the
    client's Tender BOQ and third-party datasheets/certs/drawings stay clean and the
    divider carries identity. Drawing sections (default §8) keep their native size.
    Returns (parts, is_empty)."""
    no = sec["no"]

    # AMT-authored sheets (§2 Material Sheet, §3 Traceability, §4 Material Selection)
    # carry the AMT logo on their table pages; third-party sections stay clean and
    # their divider carries the identity. Configurable via `branded_sections`.
    branded = no in cfg.get("branded_sections", [2, 3, 4])

    raw = []
    # 1) spreadsheets -> faithful table pages (logo only when branded)
    for i, x in enumerate(sec.get("xlsx_list") or ([sec["xlsx"]] if sec.get("xlsx") else [])):
        out = os.path.join(work, f"sec{no}_table{i}.pdf")
        _, eng = RT.render_table(x, out, sec["en"], ref_disp,
                                 engine=cfg.get("render_engine", "auto"), branded=branded)
        engines.add(eng)
        raw.append(out)

    # 2) Word docs -> PDF (skip a heading-only cover doc, redundant with the divider)
    for i, d in enumerate(sec["docx"]):
        if cfg.get("drop_cover_docx", True) and CD.is_cover_docx(d):
            print(f"  · §{no}: skipping cover-only doc '{os.path.basename(d)}'.")
            continue
        out = os.path.join(work, f"sec{no}_docx{i}.pdf")
        _, eng = CD.convert_docx(d, out, ref_disp, engine=cfg.get("render_engine", "auto"))
        engines.add(eng)
        raw.append(out)

    # 3) appended source PDFs (datasheets, diagrams, …)
    raw += list(sec["pdfs"])

    if not raw:
        return [], True

    # 4) uniform A4 portrait; drawing sections keep native size; nothing stamped
    uniform = cfg.get("uniform_pages", True)
    keep_native = no in cfg.get("native_sections", [8])
    parts = []
    for j, p in enumerate(raw):
        page = p
        if uniform and not keep_native:
            a4 = os.path.join(work, f"sec{no}_a4_{j}.pdf")
            try:
                NORM.to_a4_portrait(p, a4)
                page = a4
            except Exception as e:
                print(f"  ! A4 fit failed for {os.path.basename(p)} ({e}); using original.")
        parts.append(page)
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
