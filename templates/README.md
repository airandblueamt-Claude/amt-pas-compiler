# Submittal templates

Each `*.json` here defines a submittal's **section list** (the document
architecture). The compiler builds the cover, table of contents, and one divider
per section from whichever template the job's config selects.

## Selecting a template (in the job config)
Precedence — first match wins:
1. `"sections": [ ... ]` — an inline section list right in the config.
2. `"template": "drawing-submittal"` — a name resolved to `templates/<name>.json`,
   or a full path to any `.json`.
3. *(neither)* — the built-in default = `material-submittal` (AMT's 8-section PAS).

## Template / section schema
```json
{
  "name": "material-submittal",
  "toc_title_en": "Table of Content",
  "toc_title_ar": "جدول المحتويات",
  "sections": [
    { "no": 1, "prefix": "1", "kind": "table",  "optional": false,
      "en": "Tender BOQ", "ar": "جدول كميات العقد" }
  ]
}
```
Per section:
- `no` — section number shown on the divider/TOC.
- `prefix` — the **leading token of the input sub-folder name** (e.g. `"1"` matches
  `1-Tender BOQ`). Must be unique across sections.
- `kind` — `"table"` (one `.xlsx` rendered to PDF) or `"append"` (source PDFs, plus
  any real `.docx`, appended as-is).
- `optional` — if `true`, an empty folder produces a "to follow" placeholder instead
  of failing the build.
- `en` / `ar` — bilingual section titles (cover/TOC/divider). Arabic is shaped RTL.

## Make a new template
Copy `material-submittal.json`, rename it, edit the `sections`, and point the job
config at it with `"template": "<your-name>"`. Folder numbers in the input set just
need to match each section's `prefix`. No code changes.

Provided examples:
- `material-submittal.json` — the default 8-section PAS.
- `drawing-submittal.json` — a 4-section drawings package (register + drawings).
