# Faithful table rendering — design

**Date:** 2026-06-16
**Status:** Approved (direction) — pending spec review
**Component:** `amt-pas-compiler` (the PDF engine behind the PAS Generator)

## Problem

BOQ / material / selection tables keep rendering with overlapping text, clipped
cells, inconsistent fonts and logo overlap. Eight successive bug-fix commits
(`8b53af9` → `17b4d8b`) all target the same area and none fully fixed it.

Root cause: the engine **modifies the user's spreadsheet** before converting it —
it rewrites every cell's font, clears manual row heights, forces text wrapping,
rewrites alignment, and reserves margins. Those modifications are exactly what
breaks the layout. The user's sheet already looks correct in Excel; the tool's
attempt to "improve" it is the source of the overlap/clipping, and every real
sheet (wide, merged cells, embedded images, mixed Arabic/English) breaks it in a
new way that synthetic test sheets don't reproduce.

## Decision (user-approved)

**Convert each table faithfully and never re-format it. Branding lives on the
section divider pages, not on the table pages.**

- Tables come out **exactly as designed in Excel** — same fonts, row heights,
  wrapping, merged cells, images.
- **No AMT logo on any table page.** The section divider page in front of each
  section carries the full AMT logo + bilingual title (unchanged).
- This matches the user's existing branding preference: Tender (§1), Catalogue
  (§5), Vendor (§6), Warranty (§7) and Drawings (§8) already have no logo; now
  §2/§3/§4 are likewise clean-and-faithful.

## Architecture

### `render_tables.py` — convert, don't re-typeset
- `_prepare_xlsx_for_print(xlsx)` is reduced to **page setup only**: A4 paper,
  `fitToWidth = 1`, `fitToHeight = 0`, portrait. **No** cell, row-height, font or
  alignment changes whatsoever. (Fit-to-width prevents wide sheets from splitting
  columns across pages; it scales uniformly and does not alter layout. If the
  sheet already carries its own fit/scale settings they are preserved.)
- Keep the high-quality PDF export filter (lossless, no image down-sampling) so
  embedded reference images stay sharp.
- **Delete** the re-typesetting machinery: `_estimate_row_height`,
  `_rows_with_images`, `_col_points`, the per-cell `Alignment`/`Font` rewriting,
  the row-height clearing, and `TABLE_FONT` / `TABLE_FONT_SIZE`.
- `render_table(..., brand=...)` no longer stamps anything; the `brand` parameter
  and the logo-stamping path are removed from the table flow.

### `assemble.py` — one clean path for all content
- A section's content (rendered tables, converted Word docs, appended PDFs) is
  all handled the same way: each page is placed onto a uniform **A4 portrait**
  sheet via `normalize.to_a4_portrait` (scale-to-fit, rotate landscape, never
  enlarge). **No top-reserve band and no logo stamp** on any content page.
- Drawing sections (default §8) keep their **native size** (A1/A3 fold-outs).
- Remove the table-vs-appended branding split, the `top_reserve` logo band, and
  the content-page stamping. `_is_branded` / `unbranded_sections` no longer affect
  content pages (every content page is clean); dividers are always branded.

### `amt_common.py` — branding only on AMT-authored pages
- `stamp_pdf` / `_stamp_header` (content-page logo overlay) is **removed** — no
  content page is stamped any more.
- The header-logo helpers used by the cover / TOC / dividers stay (new AMT logo,
  sized by height). `HEADER_BAND` constant is removed.

### Unchanged
- Cover, Table of Contents, section dividers and placeholders — full AMT chrome,
  bilingual, clickable TOC. These carry all AMT identity.
- Background upload + manifest session (wizard) — unrelated, stays.
- Reference datasheets / certificates kept full-size and clean.

## Page size & orientation
Every content page is uniform **A4 portrait** (the user's prior requirement).
Tables are fit-to-width portrait; a long table flows onto multiple A4 pages; a
very wide sheet scales down to fit the portrait width (its layout/proportions
preserved). Drawings (§8) stay native size. This keeps a consistent document
while preserving each table's design.

## What this removes
All of: font normalization, row-height estimation, wrap forcing, alignment
rewriting, logo-band reservation, content-page stamping. ~150 lines of the most
bug-prone code, and the cause of every recurring table issue.

## Testing
- **Faithful:** a sheet converts to a PDF whose text/positions match the source
  Excel (no font swap, no row-height change). Verified by diffing cell text and
  spot-checking layout against the source.
- **No overlap / no clip:** because nothing is modified, the output matches Excel;
  verified on the real KFU selection sheet (wide, merged cells, images, Arabic).
- **Images sharp & full-size:** embedded images keep their resolution and row size.
- **Uniform pages:** every non-drawing page is A4 portrait; drawings native.
- **Branding:** dividers branded; no table page carries a logo.

## Risks / tradeoffs
- **Very wide sheets become small** in A4 portrait. Mitigation: the user sets
  "Fit Sheet on One Page" (or a sensible print area) in Excel; they control it.
- **Long tables span multiple pages** (expected and fine).
- **openpyxl page-setup round-trip**: limited to `page_setup` only (no cell/row
  edits), which is low-risk for image/format preservation.
- **reportlab fallback** (only when LibreOffice is absent) still re-typesets and
  is lower fidelity — acceptable, since the Space always has LibreOffice.

## Out of scope
Per-table logo options, landscape document mode, image re-anchoring. Can revisit
if a specific sheet needs it.
