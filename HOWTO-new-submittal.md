# How to generate a different type of PAS / submittal

There are two cases. Pick the one that matches you.

- **A. Same 8-section PAS, different project or system** (e.g. a CCTV or Fire-Alarm
  material submittal instead of IPTV). → You only change the **config**. No template.
- **B. A different document structure** (different number/order/titles of sections,
  e.g. a Method-Statement or Drawings-only package). → You add a small **template**
  JSON, then make a config that selects it.

---

## A. Same template, new project (5 minutes)

1. **Lay out the input folder** with the standard numbered sub-folders. Names can be
   anything as long as they start with the section number:
   ```
   CCTV Submittal/
     1-Tender BOQ/            <one .xlsx>
     2-Vendor BOQ/            <one .xlsx>
     3-Traceability/          <one .xlsx>
     4-Material Selection/    <one .xlsx>
     5-Datasheets/            <datasheet .pdf files>
     6-Partnership Cert/      <cert .pdf>   (optional, may be empty)
     7-Warranty/              <warranty .docx or .pdf>
     8-Drawings/              <SLD + layout .pdf files>
   ```
2. **Copy the example config** and edit the fields:
   ```bash
   cp ~/.claude/skills/amt-pas-compiler/config.example.json cctv.config.json
   ```
   Change `ref_no`, `version`, `mts_ref_no`, `date`, `project_title_en/ar`,
   `client`, the three `signoff` initials, `input_dir`, and `output_pdf`.
   Leave `template` as `"material-submittal"` (the default).
3. **Validate, then build:**
   ```bash
   python3 ~/.claude/skills/amt-pas-compiler/scripts/build_pas.py cctv.config.json --dry-run
   python3 ~/.claude/skills/amt-pas-compiler/scripts/build_pas.py cctv.config.json
   ```
   Done — that's a complete CCTV PAS.

> In Claude Code you can skip the CLI and just say: *"build the PAS for the CCTV
> project, inputs in <folder>"* and answer the questions it asks.

---

## B. A different submittal structure (add a template)

Worked example: a **Method-Statement Submittal** with these sections:
1. Cover Letter, 2. Method Statement, 3. Risk Assessment, 4. Approvals.

### Step 1 — Write the template JSON
Create `~/.claude/skills/amt-pas-compiler/templates/method-statement.json`:
```json
{
  "name": "method-statement",
  "toc_title_en": "Table of Content",
  "toc_title_ar": "جدول المحتويات",
  "sections": [
    { "no": 1, "prefix": "1", "kind": "append", "optional": false,
      "en": "Cover Letter",      "ar": "خطاب التقديم" },
    { "no": 2, "prefix": "2", "kind": "append", "optional": false,
      "en": "Method Statement",  "ar": "بيان طريقة العمل" },
    { "no": 3, "prefix": "3", "kind": "append", "optional": false,
      "en": "Risk Assessment",   "ar": "تقييم المخاطر" },
    { "no": 4, "prefix": "4", "kind": "append", "optional": true,
      "en": "Approvals",         "ar": "الموافقات" }
  ]
}
```
Rules for each section:
- `prefix` = the **number the input sub-folder starts with**; must be unique.
- `kind` = `"table"` for a single `.xlsx` (rendered as a BOQ-style table) or
  `"append"` for source PDFs / Word docs added as-is.
- `optional: true` → an empty folder becomes a "to follow" placeholder instead of
  failing the build.
- `en` / `ar` = the bilingual titles used on the cover/TOC/divider.

### Step 2 — Arrange the input folder to match the prefixes
```
Method Statement - Project X/
  1-Cover Letter/        <.docx or .pdf>
  2-Method Statement/    <.docx or .pdf>
  3-Risk Assessment/     <.pdf / .xlsx / .docx>
  4-Approvals/           <.pdf>   (optional)
```
(Folder names just need to start with 1- / 2- / 3- / 4-.)

### Step 3 — Make a config that selects the template
`method.config.json`:
```json
{
  "ref_no": "2506100-MST-001",
  "version": "00",
  "mts_ref_no": "2506100-MST-001",
  "date": "15-June-2026",
  "project_title_en": "Method Statement Submittal - Cable Installation",
  "project_title_ar": "بيان طريقة العمل - تمديد الكابلات",
  "client": "Project X",
  "signoff": {
    "prepared_by": { "role_en": "Site Engineer",     "initials": "ABC" },
    "checked_by":  { "role_en": "QA/QC Engineer",     "initials": "DEF" },
    "approved_by": { "role_en": "Project Manager",    "initials": "GHI" }
  },
  "revision": { "author": "ABC", "remarks": "Issued for approval" },

  "template": "method-statement",

  "input_dir":  "/abs/path/to/Method Statement - Project X",
  "output_pdf": "/abs/path/to/2506100-MST-001 v.00 Method Statement.pdf",

  "render_engine": "auto",
  "missing_section_mode": "placeholder",
  "drawing_fit": "native"
}
```
You can instead skip the file and put the sections **inline** in the config with a
`"sections": [ ... ]` key (same shape) — handy for a one-off.

### Step 4 — Validate, then build
```bash
python3 ~/.claude/skills/amt-pas-compiler/scripts/build_pas.py method.config.json --dry-run
python3 ~/.claude/skills/amt-pas-compiler/scripts/build_pas.py method.config.json
```
The dry run prints `Template: method-statement (4 sections)` and lists what it found
in each folder. The build then produces cover → TOC → the 4 sections.

### Step 5 — Share the new template with the team
```bash
cd ~/.claude/skills/amt-pas-compiler
git add templates/method-statement.json
git commit -m "Add method-statement template"
git push
```
Teammates `git pull` and can use `"template": "method-statement"` immediately.

---

## Config options you can tune (any submittal)
| Key | Values | Effect |
|---|---|---|
| `render_engine` | `auto` (default) / `libreoffice` / `reportlab` | How Excel & Word are rendered (LibreOffice = pixel-faithful) |
| `drawing_fit` | `native` (default) / `auto` / `portrait` / `landscape` | Keep big drawings full-size (matches sample) or scale onto A4 |
| `missing_section_mode` | `placeholder` (default) / `skip` / `error` | What happens when an optional section's folder is empty |
| `drop_cover_docx` | `true` (default) / `false` | Skip heading-only Word "cover" pages that duplicate the divider |
| `client_logo` | path | Optional client/project logo centred on the cover |

## Quick troubleshooting
- **"required sections are missing"** → a non-optional section's folder is absent or
  its `prefix` doesn't match the folder's leading number. Run `--dry-run` to see what
  was found per section.
- **Arabic looks wrong / tables cramped** → install LibreOffice (see README) and
  rebuild; the skill auto-switches to it.
- **A blank/odd page in an append section** → a heading-only Word file; it's skipped
  by default (`drop_cover_docx`).
- **Page numbers in the TOC look off** → they're computed from real content length;
  re-running always recomputes them.
