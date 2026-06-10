# Prompting the `amt-pas-compiler` skill

How to ask Claude Code to build an AMT PAS / Technical Submittal so it works the
first time and produces a clean, correct PDF. This skill auto-triggers on the right
words — you do **not** need to touch the CLI. The quality of the output depends
mostly on (a) the input folder being laid out right and (b) giving the cover details
in your prompt.

---

## The one-line version

> **"Build the PAS for `<project/system>`. Inputs are in `<absolute folder path>`.
> Ref `<ref-no>`, version `00`, dated `<date>`. Title EN: `<English title>`,
> AR: `<Arabic title>`. Client: `<client / building>`. Sign-off — prepared `<XX>`,
> checked `<XX>`, approved `<XX>`."**

Give that and the skill has everything it needs. Everything below is just detail.

---

## Before you prompt: lay out the input folder

The single biggest cause of a bad build is a wrong input folder. The skill matches
sub-folders by their **leading number** — the rest of the name is free text, so
as-received typos are fine.

```
<Project> Submittal/
  1-Tender BOQ/            one .xlsx
  2-Vendor BOQ/            one .xlsx
  3-Traceability/          one .xlsx
  4-Material Selection/    one .xlsx
  5-Datasheets/            datasheet .pdf files (sub-folders OK)
  6-Partnership Cert/      cert .pdf            (optional — may be empty)
  7-Warranty/              warranty .docx or .pdf
  8-Drawings/              SLD + layout .pdf files (sub-folders OK)
```

Rules that keep the build clean:
- **One** `.xlsx` per BOQ folder (sections 1–4). Two spreadsheets in one folder is
  ambiguous — move the extra out.
- Sections 1–5 and 8 are **required**; 6 and 7 are optional (empty → "to follow"
  placeholder, build still succeeds).
- Don't pre-merge PDFs — append sections take the source files as-is.

---

## Always dry-run first

Ask for a **dry run** before the real build. It validates discovery — lists what was
found in every section and warns on anything empty — without spending time rendering.

> **"Dry-run the PAS build for the inputs in `<folder>` and show me what it found per
> section."**

Read the output. If a section says *empty* or *missing* and it shouldn't be, fix the
folder, then ask to build. This is the difference between a perfect first PDF and
three rebuilds.

---

## Good prompts (copy, then fill the brackets)

**New project, standard 8-section PAS:**
> "Build the PAS for the CCTV system. Inputs: `/home/me/CCTV Submittal`. Ref
> 2506039-TCS-011, version 00, dated 15-June-2026. Title EN: *Material Submittal for
> CCTV System*, AR: *التقديم الفني لنظام المراقبة*. Client: KFUH – Building B02.
> Sign-off: prepared MSD, checked MZK, approved MGT. Dry-run first, then build."

**Re-run / rebuild an existing package:**
> "Rebuild the IPTV PAS from `/home/malkhalifa/zakir-folder/Technical Submital_IPTV_04June2026`
> with the same cover as before."

**Best fidelity (Arabic + merged BOQ cells):**
> "…and use LibreOffice rendering."  *(append this; it forces `render_engine:
> libreoffice` for pixel-faithful tables. Requires LibreOffice installed.)*

**A different document type (not the 8-section PAS):**
> "Build a method-statement submittal using the `method-statement` template. Inputs in
> `<folder>` with sub-folders 1-Cover Letter, 2-Method Statement, 3-Risk Assessment,
> 4-Approvals." *(See `HOWTO-new-submittal.md` for adding a new template.)*

---

## What to put in the prompt for a *perfect* cover

The cover/TOC are only as good as the details you give. Include all of these or the
skill will ask (or fall back to defaults):

| Field | Example | If you omit it |
|---|---|---|
| Ref no. | `2506039-TCS-011` | Skill asks |
| Version | `00` | Defaults to `00` |
| Date | `15-June-2026` | Skill asks |
| Title EN **and** AR | both | AR blank looks wrong on a bilingual cover |
| Client / building | `KFUH – Building B02` | Skill asks |
| Sign-off initials ×3 | prepared / checked / approved | Defaults / blanks |

Tip: paste the Arabic title yourself — don't ask Claude to translate it unless you'll
verify it. A wrong Arabic title on a formal submittal is worse than asking you for it.

---

## Knobs you can ask for in plain language

| Say this | Effect |
|---|---|
| "use LibreOffice rendering" | Pixel-faithful Excel/Word (best for Arabic) |
| "scale the drawings to A4" | `drawing_fit: auto` — fit big A1 sheets onto A4 |
| "keep drawings full size" | `drawing_fit: native` (default, matches the sample) |
| "skip empty optional sections" | `missing_section_mode: skip` (no placeholder page) |
| "fail if anything is missing" | `missing_section_mode: error` |
| "put the client logo on the cover" | Provide a path → `client_logo` |

---

## Checking the result

After the build the skill prints a **page total**, a **pass/fail page-consistency
check**, and the **TOC page map**. A good build:
- ends with `consistency check: PASS`,
- has a TOC whose page numbers match the actual section starts,
- has every AMT-authored page carrying the logo (header) + contact banner (footer).

If the page count looks off or a section is blank, ask:
> "Show me the dry-run discovery again — which files went into each section?"

---

## Quick troubleshooting by prompt

| Symptom | Say this |
|---|---|
| "required sections are missing" | "Dry-run and list what was found per section" — then fix the folder whose prefix didn't match |
| Arabic looks broken / tables cramped | "Rebuild using LibreOffice rendering" (install LibreOffice first) |
| A blank/odd page in an append section | It's a heading-only Word cover; it's dropped by default — usually fine |
| TOC page numbers look wrong | "Rebuild" — page numbers are recomputed from real content each run |
| Big drawing got cut off | "Rebuild with the drawings scaled to A4" |

---

## TL;DR

1. Lay out the numbered input folder correctly (one `.xlsx` per BOQ section).
2. Prompt with the folder path **and** the cover details (ref, version, date,
   bilingual title, client, three sign-off initials).
3. Ask for a **dry-run first**, read the discovery, then build.
4. Confirm `consistency check: PASS` and a correct TOC.

That sequence produces a clean PDF on the first build.
