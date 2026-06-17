---
name: amt-pas-compiler
description: Compile an AMT Product/Material Approval Submittal (PAS / Technical Submittal) into a single branded PDF — cover, bilingual table of contents, section dividers, rendered BOQ tables, appended datasheets, partnership/warranty letters and single-line diagrams — from a standard numbered input folder. Use when building, assembling, or generating a technical submittal, PAS, material submittal, or tender compliance package for AMT (Advanced Micro Technologies Co.).
---

# AMT PAS (Technical Submittal) Compiler

Turn a folder of submittal inputs (BOQ spreadsheets, datasheets, certificates,
drawings) into one polished, AMT-branded PDF that matches the company template:
cover → table of contents → 8 sections, each with a bilingual (English/Arabic)
divider page followed by its content.

## When to use
- "Build / generate / assemble the PAS (or technical submittal) for <project>"
- "Compile the submittal package from this folder"
- "Make the tender compliance / material submittal PDF"

## What it produces (fixed 8-section architecture)
1. Tender BOQ · 2. Material Sheet (AMT) · 3. Material Traceability & Compliance ·
4. Material Selection · 5. Product Datasheets · 6. Vendor Partnership Certificate ·
7. Warranty Certificate · 8. Layout & Single Line Diagram.

Cover carries the ref no., bilingual title, revision table and sign-off table.
AMT-authored pages (cover, TOC, dividers) carry the AMT logo header + contact
banner footer. Sections 1–4 are rendered **faithfully** from the Excel BOQ files
(LibreOffice — content untouched), then the **AMT logo header is stamped** on each
table page (see `scripts/stamp.py`); before conversion the sheet's print layout is
adjusted page-setup-only (centre horizontally with symmetric margins, fit to one
page wide, top margin reserved for the logo) — cell content/formatting is never
modified. Sections 5–8 are the source datasheets/letters/drawings appended as-is.
See `references/document-spec.md` for the full spec.

## Templates (config-driven section list)
The document's section list is a **template**, so the same engine builds other
submittal types — not just the default 8-section PAS. Select one in the job config
(first match wins):
1. `"sections": [ … ]` — an inline section list in the config.
2. `"template": "drawing-submittal"` — a name resolved to `templates/<name>.json`,
   or a path to any template `.json`.
3. *(neither)* — the built-in default `material-submittal` (the 8 sections below).

Each section: `{ "no", "prefix", "kind": "table"|"append", "optional", "en", "ar" }`
where `prefix` is the leading token of the input sub-folder name (e.g. `"1"` →
`1-Tender BOQ`) and must be unique. `table` = one `.xlsx` rendered; `append` =
source PDFs (+ real `.docx`) appended. See `templates/README.md` and the two
shipped examples. Adding a new submittal type = a new JSON file, no code changes.

## Input folder convention
Sub-folders are matched by their **leading number** (1–8), so as-received name
typos are fine (`2-AMT-VENDOR BOQ`, `8-OVERALL Single line digram`, etc.):

| # | Folder (any name starting with the number) | Content |
|---|---|---|
| 1 | `1-...` | one `.xlsx` (Tender BOQ) |
| 2 | `2-...` | one `.xlsx` (Vendor/AMT BOQ) |
| 3 | `3-...` | one `.xlsx` (Traceability) |
| 4 | `4-...` | one `.xlsx` (Material Selection) |
| 5 | `5-...` | datasheet `.pdf`s (recursive) |
| 6 | `6-...` | partnership cert `.pdf`(s) — may be empty |
| 7 | `7-...` | warranty `.docx`/`.pdf` |
| 8 | `8-...` | SLD + layout `.pdf`s (+ `.docx`), recursive |

## How to run
1. Copy `config.example.json`, fill in ref no., bilingual title, date, client and
   sign-off initials, and set `input_dir` + `output_pdf`.
2. Dry run to validate discovery: shows each section's resolved files and warns on
   anything empty.
   ```bash
   python3 scripts/build_pas.py my.config.json --dry-run
   ```
3. Build:
   ```bash
   python3 scripts/build_pas.py my.config.json
   ```
   Output is a single PDF at `output_pdf`. The run prints a page total, a
   pass/fail page-consistency check, and the TOC page map.

## Rendering engines (`render_engine` in config)
- `auto` (default): use **LibreOffice** if `soffice` is on PATH (pixel-faithful
  Excel/Word incl. Arabic RTL and merged cells), otherwise the built-in
  **reportlab** renderer (self-contained, AMT-styled, no external binary).
- `libreoffice`: force LibreOffice (errors if not installed).
- `reportlab`: force the built-in renderer.

For best fidelity install LibreOffice once, then re-run (auto-detected):
```bash
sudo apt-get install -y libreoffice-calc libreoffice-writer
```

## Oversized drawings (`drawing_fit`, `normalize_appended`)
Large source pages — e.g. A1 single-line diagrams and floor layouts in §8.
- `drawing_fit`:
  - `native` (default, **matches the sample document**): append drawings at their
    original size/rotation (the sample keeps them full A1 — a mixed-size PDF).
  - `auto`: scale each oversized sheet down onto A4, keeping its orientation
    (landscape drawing → A4-landscape), aspect preserved, centred, never enlarged.
  - `portrait` / `landscape`: force every appended page to that A4 orientation.
- `normalize_appended` (default `true`): when `drawing_fit` is not `native`, fit
  appended pages to A4; set `false` to never resize anything.

## Missing / empty sections (`missing_section_mode`)
- `placeholder` (default): keep the section, insert a divider page with a
  "To be submitted / Certificate to follow" note (correct pagination preserved).
- `skip`: divider page only, no content.
- `error`: refuse to build until content is provided. *(Required sections 1–5, 8
  always error if missing; only optional sections 6–7 honor this setting.)*

## Dependencies
Python 3 with: `reportlab`, `pypdf`, `openpyxl`, `Pillow`, `arabic-reshaper`,
`python-bidi`, `pymupdf` (install:
`pip install reportlab pypdf openpyxl pillow arabic-reshaper python-bidi pymupdf`).
Optional but strongly recommended: LibreOffice (`soffice`) for pixel-faithful
Excel/Word rendering — auto-detected on PATH, in `$AMT_SOFFICE`, or under
`~/.local/opt/.../program/soffice` (a user-extracted AppImage works).
Fonts (Amiri) and branding assets ship in `assets/`.

## Files
- `scripts/build_pas.py` — CLI entry point (validate → discover → build)
- `scripts/discover.py` — folder discovery + validation report
- `scripts/render_tables.py` — §1–4 xlsx → PDF (LibreOffice or reportlab)
- `scripts/convert_docx.py` — §7/§8 docx → PDF (LibreOffice or text fallback)
- `scripts/normalize.py` — fit oversized appended pages (A1 drawings) onto A4
- `scripts/build_chrome.py` — cover / TOC / dividers / placeholders
- `scripts/assemble.py` — two-pass pagination + merge + QA
- `scripts/amt_common.py` — branding, geometry, Arabic shaping, page chrome
- `scripts/pas_spec.py` — template resolution + default section list
- `templates/*.json` — submittal templates (section lists); `templates/README.md`
- `scripts/extract_assets.py` — one-off: re-extract footer banner + seal from a sample
- `assets/` — `amt-logo.png`, `amt-footer-banner.png`, `amt-seal.png`, `fonts/`
- `references/document-spec.md`, `references/sample-page-map.md`
- `config.example.json`
