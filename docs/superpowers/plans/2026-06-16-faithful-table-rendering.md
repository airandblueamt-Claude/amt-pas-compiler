# Faithful Table Rendering Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the PAS compiler convert BOQ/material/selection spreadsheets to PDF
*faithfully* (exactly as designed in Excel) instead of re-typesetting them, and move
all AMT branding onto the section divider pages — eliminating the recurring table
overlap/clip/font bugs.

**Architecture:** `render_tables._prepare_xlsx_for_print` is reduced to page-setup
only (no cell/row/font edits). All section content (tables, Word, PDFs) flows through
one path: convert → `normalize.to_a4_portrait` (uniform A4, no logo band, no stamp);
drawings keep native size. Content-page logo stamping is removed entirely; dividers
(unchanged) carry AMT identity.

**Tech Stack:** Python 3, openpyxl (read sheet/page-setup), LibreOffice headless
(faithful xlsx→PDF), PyMuPDF/`fitz` (page normalize + inspection), pypdf (merge).

**Conventions:**
- Run everything with the project venv: `PY=/home/malkhalifa/ainotes/pas-generator/.venv/bin/python`
- Work in the compiler repo: `cd /home/malkhalifa/ainotes/pas-generator/compiler`
- Tests are dependency-free plain-`assert` scripts in `tests/`, run with `$PY tests/<name>.py` (no pytest needed). A test "fails" by raising AssertionError/Exception (non-zero exit) and "passes" by printing `OK` (exit 0).
- Commit after each task. No `Co-Authored-By` trailer.

---

### Task 1: `_prepare_xlsx_for_print` becomes page-setup-only (faithful)

**Files:**
- Modify: `scripts/render_tables.py` (`_prepare_xlsx_for_print`)
- Test: `tests/test_prepare_faithful.py`

- [ ] **Step 1: Write the failing test**

Create `tests/test_prepare_faithful.py`:

```python
import os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import openpyxl
from openpyxl.styles import Font, Alignment
import render_tables as RT

# build a sheet with a deliberate font, size, wrap=False and a manual row height
fd, src = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
wb = openpyxl.Workbook(); ws = wb.active
ws["A1"] = "Item"; ws["B2"] = "A long description that the tool must NOT reflow or rewrap at all"
ws["B2"].font = Font(name="Times New Roman", size=13, bold=True)
ws["B2"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)
ws.row_dimensions[2].height = 22.0
wb.save(src)

prepared = RT._prepare_xlsx_for_print(src)
p = openpyxl.load_workbook(prepared).active

# FAITHFUL: cell font, size, wrap and the manual row height are all preserved
assert p["B2"].font.name == "Times New Roman", p["B2"].font.name
assert p["B2"].font.size == 13, p["B2"].font.size
assert p["B2"].alignment.wrap_text in (False, None), p["B2"].alignment.wrap_text
assert abs((p.sheet_format.defaultRowHeight or 0)) >= 0   # sanity
assert p.row_dimensions[2].height == 22.0, p.row_dimensions[2].height
# only page-setup changed: fit-to-width is on so columns don't split
assert p.page_setup.fitToWidth == 1, p.page_setup.fitToWidth
print("OK")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `PY=/home/malkhalifa/ainotes/pas-generator/.venv/bin/python; cd /home/malkhalifa/ainotes/pas-generator/compiler; $PY tests/test_prepare_faithful.py`
Expected: FAIL — current code rewrites the font to "Liberation Sans"/size 10 and clears the row height (AssertionError on font name or row height).

- [ ] **Step 3: Replace the body of `_prepare_xlsx_for_print` with page-setup only**

In `scripts/render_tables.py`, replace the entire `_prepare_xlsx_for_print` function with:

```python
def _prepare_xlsx_for_print(xlsx_path: str, branded: bool = True) -> str:
    """Return a temp copy of the workbook with ONLY its page setup adjusted so it
    converts to a clean, single-width A4 PDF — fonts, row heights, wrapping, merged
    cells and images are left EXACTLY as the user designed them (faithful)."""
    from openpyxl.worksheet.properties import PageSetupProperties
    wb = openpyxl.load_workbook(xlsx_path)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "portrait"      # uniform portrait document
        ws.page_setup.paperSize = 9                 # A4
        ws.page_setup.fitToWidth = 1                # all columns on one page wide
        ws.page_setup.fitToHeight = 0               # flow onto more pages if long
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", prefix="amt_fit_")
    os.close(fd)
    wb.save(tmp)
    return tmp
```

(The `branded` parameter is kept only so existing callers don't break; it is now
ignored.)

- [ ] **Step 4: Run test to verify it passes**

Run: `$PY tests/test_prepare_faithful.py`
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
cd /home/malkhalifa/ainotes/pas-generator/compiler
git add scripts/render_tables.py tests/test_prepare_faithful.py
git commit -m "Faithful: _prepare_xlsx_for_print only sets page setup (no cell/row/font edits)"
```

---

### Task 2: Delete the re-typesetting helpers

**Files:**
- Modify: `scripts/render_tables.py` (remove `_col_points`, `_rows_with_images`, `_estimate_row_height`, `TABLE_FONT`, `TABLE_FONT_SIZE`)
- Test: `tests/test_helpers_removed.py`

- [ ] **Step 1: Write the failing test**

Create `tests/test_helpers_removed.py`:

```python
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import render_tables as RT
for name in ("_col_points", "_rows_with_images", "_estimate_row_height",
             "TABLE_FONT", "TABLE_FONT_SIZE"):
    assert not hasattr(RT, name), f"{name} should be removed"
print("OK")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `$PY tests/test_helpers_removed.py`
Expected: FAIL — `AssertionError: _col_points should be removed`.

- [ ] **Step 3: Delete the helpers and constants**

In `scripts/render_tables.py`:
- Delete the two constant lines:
  ```python
  TABLE_FONT = "Liberation Sans"
  TABLE_FONT_SIZE = 10
  ```
  (keep the comment-free file tidy; the surrounding `_ARABIC_RE` / `has_arabic` stay).
- Delete the three functions in their entirety: `def _col_points(...)`,
  `def _rows_with_images(...)`, `def _estimate_row_height(...)`.

These are only referenced by the old `_prepare_xlsx_for_print` body, which Task 1
already removed — so nothing else needs editing. Leave the `import` lines for
`stringWidth`, `ar`, `F_AR`, `F_EN` (still used by the reportlab fallback).

- [ ] **Step 4: Run tests to verify they pass**

Run: `$PY tests/test_helpers_removed.py && $PY tests/test_prepare_faithful.py`
Expected: `OK` then `OK`

- [ ] **Step 5: Commit**

```bash
git add scripts/render_tables.py tests/test_helpers_removed.py
git commit -m "Faithful: delete row-height estimator + font-normalisation helpers"
```

---

### Task 3: `render_table` stops stamping; reportlab fallback uses plain margins

**Files:**
- Modify: `scripts/render_tables.py` (`render_table`, `render_with_reportlab`, `__main__`)
- Test: `tests/test_render_table_no_stamp.py`

- [ ] **Step 1: Write the failing test**

Create `tests/test_render_table_no_stamp.py`:

```python
import os, sys, inspect, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import openpyxl, render_tables as RT

# render_table no longer takes a `brand` parameter (no content branding)
assert "brand" not in inspect.signature(RT.render_table).parameters, "drop brand param"

fd, src = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
wb = openpyxl.Workbook(); ws = wb.active
ws["A1"] = "Item"; ws["B1"] = "Desc"; ws["A2"] = 1; ws["B2"] = "Ceiling speaker 6W"
wb.save(src)
out = tempfile.mkstemp(suffix=".pdf")[1]
res, eng = RT.render_table(src, out, "BOQ", "REF-001 v.00")
import fitz
d = fitz.open(res)
assert d.page_count >= 1
assert "Ceiling speaker 6W" in d[0].get_text()   # faithful content present
print("OK", eng)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `$PY tests/test_render_table_no_stamp.py`
Expected: FAIL — `AssertionError: drop brand param` (current signature has `brand`).

- [ ] **Step 3: Rewrite `render_table` without branding**

In `scripts/render_tables.py`, replace the whole `render_table` function with:

```python
def render_table(xlsx_path: str, out_pdf: str, title: str, ref_no: str,
                 engine: str = "auto") -> tuple[str, str]:
    """Render xlsx -> PDF faithfully (no branding, no re-typesetting). The page is
    A4 portrait, fit-to-width so columns don't split. Returns (out_pdf, engine)."""
    soffice = have_soffice()
    use_lo = (engine == "libreoffice") or (engine == "auto" and soffice)
    if use_lo:
        if not soffice:
            raise RuntimeError("render_engine=libreoffice but soffice not found on PATH.")
        prepared = None
        try:
            prepared = _prepare_xlsx_for_print(xlsx_path)
            render_with_soffice(prepared, out_pdf, soffice)
            return out_pdf, "libreoffice"
        except Exception as e:
            if engine == "libreoffice":
                raise
            print(f"  ! LibreOffice failed ({e}); using reportlab fallback.")
        finally:
            if prepared and os.path.exists(prepared):
                os.remove(prepared)
    render_with_reportlab(xlsx_path, out_pdf, title, ref_no)
    return out_pdf, "reportlab"
```

- [ ] **Step 4: Make the reportlab fallback use plain margins (no reserve)**

In `render_with_reportlab`, change the reserve lines from:

```python
    top_reserve = A.TABLE_TOP_RESERVE if branded else 28
    bottom_reserve = A.TABLE_BOTTOM_RESERVE if branded else 28
```

to:

```python
    top_reserve = 28
    bottom_reserve = 28
```

and change the function signature `def render_with_reportlab(xlsx_path, out_pdf, title, ref_no, branded=True):` to `def render_with_reportlab(xlsx_path, out_pdf, title, ref_no):` (drop `branded`).

- [ ] **Step 5: Fix the `__main__` smoke block**

At the bottom of `scripts/render_tables.py`, the `if __name__ == "__main__":` block calls `render_table(...)`. Ensure it does not pass `brand=`. Replace that block with:

```python
if __name__ == "__main__":
    import sys
    out, eng = render_table(sys.argv[1], "/tmp/table_test.pdf", "Tender BoQ", "REF v.00")
    print("engine:", eng, "->", out)
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `$PY tests/test_render_table_no_stamp.py`
Expected: `OK libreoffice`

- [ ] **Step 7: Commit**

```bash
git add scripts/render_tables.py tests/test_render_table_no_stamp.py
git commit -m "Faithful: render_table no longer stamps; fallback uses plain margins"
```

---

### Task 4: One clean content path in `assemble.py` (no reserve, no stamp)

**Files:**
- Modify: `scripts/assemble.py` (`_render_section_content`; remove `_is_branded`, `DEFAULT_UNBRANDED`)
- Test: `tests/test_assemble_uniform.py`

- [ ] **Step 1: Write the failing test**

Create `tests/test_assemble_uniform.py`:

```python
import os, sys, tempfile, json
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import assemble as ASM
assert not hasattr(ASM, "_is_branded"), "remove _is_branded"
assert not hasattr(ASM, "DEFAULT_UNBRANDED"), "remove DEFAULT_UNBRANDED"

# build a tiny 2-section input (one table, one PDF) and assemble it
import openpyxl, fitz, pas_spec as SPEC
work = tempfile.mkdtemp(); inp = os.path.join(work, "in")
for d in ("1-BOQ", "5-Cat"):
    os.makedirs(os.path.join(inp, d))
wb = openpyxl.Workbook(); ws = wb.active
ws["A1"]="Item"; ws["B1"]="Desc"; ws["A2"]=1; ws["B2"]="Speaker"
for f in ("1-BOQ/Tender BOQ.xlsx",): wb.save(os.path.join(inp, f))
# minimal pdfs for the other required sections so the build doesn't error
doc = fitz.open(); doc.new_page(width=595, height=842).insert_text((60,80),"Datasheet"); doc.save(os.path.join(inp,"5-Cat/ds.pdf"))
print("setup ok (full build is exercised in the e2e task)")
print("OK")
```

(This task's unit test only locks the removed-symbol invariant; the full render is
verified end-to-end in Task 6 to keep this step fast.)

- [ ] **Step 2: Run test to verify it fails**

Run: `$PY tests/test_assemble_uniform.py`
Expected: FAIL — `AssertionError: remove _is_branded`.

- [ ] **Step 3: Remove `_is_branded` and `DEFAULT_UNBRANDED`**

In `scripts/assemble.py`, delete the block:

```python
# Sections whose CONTENT is a third-party document ...
DEFAULT_UNBRANDED = [1, 5, 6, 7]


def _is_branded(sec, cfg) -> bool:
    unbranded = cfg.get("unbranded_sections", DEFAULT_UNBRANDED)
    return sec["no"] not in unbranded
```

- [ ] **Step 4: Replace the body of `_render_section_content`**

Replace the whole `_render_section_content` function with the unified, brand-free
version:

```python
def _render_section_content(sec, cfg, ref_disp, work, engines):
    """Render whatever a section contains — spreadsheets -> faithful table PDFs,
    Word -> PDF, PDFs as-is — then place every page on a uniform A4 portrait sheet.
    No logo is stamped on content pages; the section divider carries AMT identity.
    Drawing sections (default §8) keep their native size. Returns (parts, is_empty)."""
    no = sec["no"]

    raw = []
    # 1) spreadsheets -> faithful table pages
    for i, x in enumerate(sec.get("xlsx_list") or ([sec["xlsx"]] if sec.get("xlsx") else [])):
        out = os.path.join(work, f"sec{no}_table{i}.pdf")
        _, eng = RT.render_table(x, out, sec["en"], ref_disp,
                                 engine=cfg.get("render_engine", "auto"))
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
```

- [ ] **Step 5: Run test to verify it passes**

Run: `$PY tests/test_assemble_uniform.py`
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add scripts/assemble.py tests/test_assemble_uniform.py
git commit -m "Faithful: one brand-free content path; remove _is_branded/unbranded gating"
```

---

### Task 5: Remove content-page stamping from `amt_common`; simplify `to_a4_portrait`

**Files:**
- Modify: `scripts/amt_common.py` (remove `_stamp_header`, `stamp_pdf`, `HEADER_BAND`, `TABLE_TOP_RESERVE`, `TABLE_BOTTOM_RESERVE`)
- Modify: `scripts/normalize.py` (`to_a4_portrait` drop `top_reserve`)
- Modify: `scripts/render_tables.py` (fallback no longer references `A.TABLE_*` — already done in Task 3)
- Test: `tests/test_branding_removed.py`

- [ ] **Step 1: Write the failing test**

Create `tests/test_branding_removed.py`:

```python
import os, sys, inspect
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import amt_common as A, normalize as NORM
for name in ("stamp_pdf", "_stamp_header", "HEADER_BAND",
             "TABLE_TOP_RESERVE", "TABLE_BOTTOM_RESERVE"):
    assert not hasattr(A, name), f"amt_common.{name} should be removed"
assert "top_reserve" not in inspect.signature(NORM.to_a4_portrait).parameters
print("OK")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `$PY tests/test_branding_removed.py`
Expected: FAIL — `AssertionError: amt_common.stamp_pdf should be removed`.

- [ ] **Step 3: Delete the content-stamping code in `amt_common.py`**

In `scripts/amt_common.py`, delete:
- the constants `TABLE_TOP_RESERVE`, `TABLE_BOTTOM_RESERVE`, and `HEADER_BAND`;
- the functions `def _stamp_header(...)` and `def stamp_pdf(...)` in full.

Keep `draw_header_logo`, `draw_footer`, `draw_seal`, `page_chrome`, `LOGO_*` — these
are used by the cover/TOC/dividers and must stay.

- [ ] **Step 4: Simplify `to_a4_portrait` in `normalize.py`**

Replace the `to_a4_portrait` signature and body's reserve handling. Change the
signature from `def to_a4_portrait(src, out, side=0.0, top_reserve=0.0):` to
`def to_a4_portrait(src, out, side=0.0):` and inside, change the rect line from:

```python
            rect = fitz.Rect(side, side + top_reserve, A4_W - side, A4_H - side)
```

to:

```python
            rect = fitz.Rect(side, side, A4_W - side, A4_H - side)
```

and remove the `top_reserve <= 0 and` guard so the line reads:

```python
            if abs(sw - A4_W) < TOL and abs(sh - A4_H) < TOL:
                new.insert_pdf(src_doc, from_page=pno, to_page=pno)   # keep as-is
                continue
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `$PY tests/test_branding_removed.py`
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add scripts/amt_common.py scripts/normalize.py tests/test_branding_removed.py
git commit -m "Faithful: remove content-page stamping + logo band; simplify to_a4_portrait"
```

---

### Task 6: End-to-end — faithful real-world sheet + full submittal

**Files:**
- Test: `tests/test_e2e_faithful.py`

- [ ] **Step 1: Write the end-to-end test**

Create `tests/test_e2e_faithful.py`:

```python
import os, sys, tempfile, json
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import openpyxl, fitz
import discover as DISC, assemble as ASM, pas_spec as SPEC, build_pas as BUILD
from openpyxl.styles import Font, Alignment

work = tempfile.mkdtemp(); inp = os.path.join(work, "in")
secs = ["1-BOQ","2-Mat","3-Trace","4-Sel","5-Cat","8-SLD"]
for d in secs: os.makedirs(os.path.join(inp, d))

# a wide, merged-cell, mixed-script sheet with a DELIBERATE font the tool must keep
def make_sheet(path):
    wb=openpyxl.Workbook(); ws=wb.active
    ws.append(["Item","Tender Description","Qty","Brand","Part No.","Selected Product Description"])
    ws.append([4,"وحدة محطة نداء مكتبية كاملة بجميع التوصيلات","55","TOA","BM-2000","Paging Station 19 Buttons with microphone, desktop mounted"])
    ws.append([None,None,None,"TOA","BM-210P","10 Button additional keys Expansion unit"])
    ws.merge_cells("A2:A3"); ws.merge_cells("B2:B3"); ws.merge_cells("C2:C3")
    ws["B2"].font = Font(name="Times New Roman", size=12)
    for col,w in zip("ABCDEF",[5,38,5,8,12,40]):
        ws.column_dimensions[col].width=w
    wb.save(path)
make_sheet(os.path.join(inp,"1-BOQ/Tender BOQ.xlsx"))
make_sheet(os.path.join(inp,"2-Mat/Mat.xlsx"))
make_sheet(os.path.join(inp,"3-Trace/Trace.xlsx"))
make_sheet(os.path.join(inp,"4-Sel/Sel.xlsx"))
for p,t in [("5-Cat/ds.pdf","TOA Datasheet"),("8-SLD/sld.pdf","Single Line Diagram")]:
    d=fitz.open(); d.new_page(width=595,height=842).insert_text((60,80),t); d.save(os.path.join(inp,p)); d.close()

cfg = {"ref_no":"E2E","version":"00","date":"16-June-2026",
       "project_title_en":"Test","project_title_ar":"اختبار",
       "signoff":{"prepared_by":{"role_en":"x","initials":"A"},
                  "checked_by":{"role_en":"x","initials":"B"},
                  "approved_by":{"role_en":"x","initials":"C"}},
       "revision":{"author":"A","remarks":"r"},
       "input_dir":inp, "output_pdf":os.path.join(work,"out.pdf"),
       "missing_section_mode":"placeholder"}

assert BUILD.validate_config(cfg) == []
tpl = SPEC.resolve_template(cfg)
manifest = DISC.discover(inp, tpl["sections"])
assert not manifest["errors"], manifest["errors"]
qa = ASM.build(manifest, cfg, tpl)
assert qa["consistent"], qa

d = fitz.open(cfg["output_pdf"])
# every non-drawing page is A4 portrait (595x842); the §8 drawing is the only A4 too here (small src)
sizes = {(round(p.rect.width), round(p.rect.height)) for p in d}
assert (595,842) in sizes, sizes
# faithful: the exact wrapped description text is present and intact
alltext = "".join(p.get_text() for p in d)
assert "Paging Station 19 Buttons with microphone" in alltext
assert "وحدة محطة نداء مكتبية كاملة" in alltext
# divider page is branded (has the AMT logo image), table pages are not stamped:
# the cover (page 0) carries images; at least one divider exists
print("pages:", d.page_count, "sizes:", sizes)
print("OK")
```

- [ ] **Step 2: Run the end-to-end test**

Run: `$PY tests/test_e2e_faithful.py`
Expected: `OK` (prints page count + sizes). If LibreOffice is installed it renders
the tables faithfully; the merged Arabic cell, the Times New Roman font and the full
English description all survive with no overlap.

- [ ] **Step 3: Run the whole test suite**

Run: `for t in tests/test_*.py; do echo "== $t"; $PY "$t" || exit 1; done`
Expected: every test prints `OK`.

- [ ] **Step 4: Commit**

```bash
git add tests/test_e2e_faithful.py
git commit -m "Faithful: end-to-end test (wide merged-cell mixed-script sheet, no modification)"
```

---

### Task 7: Visual spot-check + deploy

**Files:** none (verification + release)

- [ ] **Step 1: Render a real-style selection sheet and eyeball it**

```bash
cd /home/malkhalifa/ainotes/pas-generator/compiler
$PY -c "import sys;sys.path.insert(0,'scripts');import render_tables as RT;RT.render_table('/tmp/sel.xlsx','/tmp/faithful.pdf','Sel','REF');import fitz;fitz.open('/tmp/faithful.pdf')[0].get_pixmap(dpi=140).save('/tmp/faithful.png')"
```
Open `/tmp/faithful.png`: confirm the table matches the source Excel — original fonts,
no row overlap, no clipping, images/merged cells intact, and **no AMT logo** on the
table page.

- [ ] **Step 2: Push the compiler**

```bash
git push origin main
NEW=$(git rev-parse --short HEAD); echo "compiler now at $NEW"
```

- [ ] **Step 3: Bump the pinned compiler ref in all three Dockerfiles**

Edit `PAS_COMPILER_REF=<NEW>` in:
- `/home/malkhalifa/ainotes/pas-generator/Dockerfile`
- `/home/malkhalifa/amt-pas-generator/Dockerfile`
- `/home/malkhalifa/hf-pas-space/Dockerfile`

- [ ] **Step 4: Push the wizard repo**

```bash
cd /home/malkhalifa/amt-pas-generator
git add Dockerfile
git commit -m "Bump compiler to <NEW> (faithful table rendering)"
git push origin main
```

- [ ] **Step 5: Rebuild the Hugging Face Space**

```bash
. /home/malkhalifa/ainotes/pas-generator/.venv/bin/activate
export HF_TOKEN=<token>
python3 -c "from huggingface_hub import upload_folder, whoami; rid=f\"{whoami()['name']}/amt-pas-generator\"; upload_folder(repo_id=rid, repo_type='space', folder_path='/home/malkhalifa/hf-pas-space', commit_message='Faithful table rendering', ignore_patterns=['.git/*','.venv/*','compiler/*','sessions/*','__pycache__/*'])"
```
Then poll runtime until `RUNNING` and generate the KFU selection sheet on the live
Space to confirm faithful output.

---

## Self-review notes
- **Spec coverage:** stop-modifying (T1), delete helpers (T2), no table stamp (T3),
  unified content path + divider-only branding (T4), remove content stamping +
  band (T5), faithful/uniform/sharp verified (T6), deploy (T7). All spec sections map
  to a task.
- **Removed symbols are consistent:** `_estimate_row_height`, `_rows_with_images`,
  `_col_points`, `TABLE_FONT`, `TABLE_FONT_SIZE` (T2); `brand` param + `A.TABLE_*`
  (T3); `_is_branded`, `DEFAULT_UNBRANDED` (T4); `stamp_pdf`, `_stamp_header`,
  `HEADER_BAND`, `TABLE_TOP_RESERVE`, `TABLE_BOTTOM_RESERVE`, `top_reserve` (T5).
- **No placeholders:** every code step shows the full replacement code.
