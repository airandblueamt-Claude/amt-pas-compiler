# Walkthrough — clone the skill and build a PAS (end to end)

A copy-paste guide a teammate can follow on a fresh machine to go from nothing to a
finished AMT PAS / Technical Submittal PDF. The worked example is the **IPTV &
Digital Signage** submittal (KFUH – Building B02); swap the values for any project.

> TL;DR: clone → `bash install.sh` → lay out the numbered input folder → copy &
> edit a config → **dry-run** → build. Always dry-run before you build.

---

## Step 0 — Prerequisites (once per machine)

- **Python 3.9+** and **git** — verify: `python3 --version`, `git --version`
- *(Recommended)* **LibreOffice** — pixel-faithful Excel/Word rendering incl. Arabic.
  Without it the skill still builds with a built-in fallback renderer (lower fidelity).

```bash
# Ubuntu/Debian (recommended):
sudo apt-get install -y libreoffice-calc libreoffice-writer
# macOS:  brew install --cask libreoffice
```

## Step 1 — Clone into Claude's skills folder

The skill must live at `~/.claude/skills/amt-pas-compiler/` so Claude Code detects it.

```bash
git clone https://github.com/airandblueamt-Claude/amt-pas-compiler.git \
  ~/.claude/skills/amt-pas-compiler
cd ~/.claude/skills/amt-pas-compiler
```

## Step 2 — Run the installer

Installs Python dependencies, checks for LibreOffice, verifies imports, and runs the
test suite.

```bash
bash install.sh
```

Healthy output ends with:

```
Verifying the skill imports cleanly …
  ✓ OK — default template 'material-submittal' with 8 sections.
Running tests …
PASS — required empty -> error, optional empty -> warning, content -> ok
  ✓ All tests passed.
```

## Step 3 — Lay out the input folder

Sub-folders are matched by their **leading number** — the rest of the name is free
text, so as-received typos (`VENDEOR`, `digram`) are fine. One `.xlsx` per BOQ
section (1–4); PDFs in sections 5 and 8.

```
Technical Submital_IPTV_04June2026/
  1-Tender BOQ/                          Tender BoQ.xlsx
  2-AMT-VENDOR BOQ/                      VENDEOR BoQ.xlsx
  3-AMT - VENDOR Material Traceability/  AMT - VENDEOR BoQ.xlsx
  4-Material Selection/                  Material Selection.xlsx
  5-Product Datasheet & Catalogue new/   10 datasheet PDFs
  6- Vendor Partnership Cert/            partnership cert PDF   (optional)
  7-Warranty/                            Warranty Letter -IPTV.docx
  8-OVERALL Single line digram/          SLD PDF + 9 layout PDFs + Shop Drawings.docx
```

| # | Required? | Holds |
|---|-----------|-------|
| 1–4 | yes | one `.xlsx` each (rendered as BOQ tables) |
| 5 | yes | datasheet PDFs |
| 6 | optional | partnership certificate (empty → "to follow" placeholder) |
| 7 | optional | warranty `.docx`/`.pdf` |
| 8 | yes | SLD + layout drawings |

## Step 4 — Create the job config

```bash
cp config.example.json iptv.config.json
```

Edit `iptv.config.json` — the IPTV values:

```jsonc
{
  "ref_no": "2506038-TCS-009",
  "version": "00",
  "date": "04-June-2026",
  "project_title_en": "Material Submittal for IPTV and Digital Signage System",
  "project_title_ar": "التقديم الفني لنظام IPTV والشاشات الرقمية",
  "client": "King Fahd University Hospital (KFUH) - Building B02",
  "building": "B02",
  "template": "material-submittal",            // standard 8-section PAS
  "signoff": {
    "prepared_by": { "role_en": "Sales Coordinator",  "initials": "MSD" },
    "checked_by":  { "role_en": "Document Controller", "initials": "MZK" },
    "approved_by": { "role_en": "Presales Manager",    "initials": "MGT" }
  },
  "input_dir":  "/ABSOLUTE/path/to/Technical Submital_IPTV_04June2026",
  "output_pdf": "/ABSOLUTE/path/to/2506038-TCS-009 v.00 PAS.pdf",
  "render_engine": "auto",
  "drawing_fit": "native"
}
```

`input_dir` and `output_pdf` **must** be absolute paths on your machine.

## Step 5 — Dry-run first (validates, writes no PDF)

```bash
python3 scripts/build_pas.py iptv.config.json --dry-run
```

Real IPTV output:

```
Template: material-submittal (8 sections)
  1. Tender BOQ                       Tender BoQ.xlsx
  2. Material Sheet Provided by AMT   VENDEOR BoQ.xlsx
  3. Material Traceability Sheet …    AMT - VENDEOR BoQ.xlsx
  4. Material Selection               Material Selection.xlsx
  5. Product's Datasheet & Catalogue  10 pdf
  6. Vendor Partnership Certificate   — empty (placeholder) —
  7. Warranty Certificate             1 docx
  8. Layout and Single Line Diagram   10 pdf, 1 docx

Warnings:
  ! Section 6 (Vendor Partnership Certificate): no documents found — a placeholder will be inserted.

Dry run — no PDF written.
```

**Read the warnings.** §6 (partnership cert) is empty → optional, so the build
proceeds with a placeholder. If a *required* section (1–5, 8) were empty, the tool
would **stop with an error** instead — that's the safeguard.

## Step 6 — Build the PDF

```bash
python3 scripts/build_pas.py iptv.config.json
```

Ends with:

```
============================================================
Output : /…/2506038-TCS-009 v.00 PAS.pdf
Pages  : <N> (expected <N>) -> OK
Engine : libreoffice        # or 'reportlab' if LibreOffice isn't installed
```

`-> OK` = the page-consistency check passed (TOC page numbers match section starts).

## Step 7 — Quality-check the result

Open the PDF and confirm:

- **Cover** — ref no., bilingual title, revision + sign-off tables.
- **TOC** — page numbers line up with where each section actually starts.
- Every AMT-authored page has the **logo (header)** + **contact banner (footer)** with
  the `Ref.:` line.
- §6 shows the "to follow" placeholder (expected — the cert is missing).

---

## Shortcut — inside Claude Code

Steps 4–6 collapse to one sentence:

> *"Build the IPTV PAS — inputs in `<folder>`, ref 2506038-TCS-009 v.00, the usual
> KFUH B02 cover. Dry-run first, then build."*

The skill triggers automatically. See [PROMPTING.md](PROMPTING.md) for prompt
patterns, and [HOWTO-new-submittal.md](HOWTO-new-submittal.md) for other submittal
types (new project, or a different document structure via a template).

---

## Notes for the team

- The **partnership certificate (§6) is genuinely missing** from this package — a real
  open item to chase before issuing, not a tool problem.
- Install **LibreOffice** on whichever machine produces the final issued copy — it is
  what renders the Arabic and merged BOQ cells crisply.
- The actual submittal contents (BOQ pricing, datasheets, drawings) are **controlled
  documents** — keep them in your document-control store, not in this public repo.
  This repo holds only the *tool*.
```
