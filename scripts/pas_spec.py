"""
pas_spec.py — the submittal architecture (section list + bilingual titles).

The default is AMT's 8-section Material/Product Approval Submittal, but the section
list is now CONFIG-DRIVEN so the same engine can produce other submittal templates:

  * cfg["sections"]  — an inline list of section dicts, or
  * cfg["template"]  — a name (looked up in ../templates/<name>.json) or a path to
                       a template JSON, or
  * neither          — the built-in DEFAULT_SECTIONS below.

Each section dict:
  { "no": 1, "prefix": "1", "kind": "table"|"append",
    "optional": false, "en": "English title", "ar": "Arabic title" }

kind:
  "table"  -> a single .xlsx rendered to PDF
  "append" -> source PDFs (+ optional .docx) appended as-is
"""
from __future__ import annotations

import os
import json

# Section content kinds
TABLE = "table"
APPEND = "append"

TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                             "templates")

# Document control labels (TOC heading) — overridable per template
TOC_TITLE_EN = "Table of Content"
TOC_TITLE_AR = "جدول المحتويات"

DEFAULT_SECTIONS = [
    dict(no=1, prefix="1", kind=TABLE,  optional=False,
         en="Tender BOQ",
         ar="جدول كميات العقد"),
    dict(no=2, prefix="2", kind=TABLE,  optional=False,
         en="Material Sheet Provided by AMT",
         ar="جدول المواد المقدمة من AMT"),
    dict(no=3, prefix="3", kind=TABLE,  optional=False,
         en="Material Traceability Sheet & Tender BoQ Compliance",
         ar="جدول تتبع المواد والمطابقة مع جدول كميات العقد"),
    dict(no=4, prefix="4", kind=TABLE,  optional=False,
         en="Material Selection",
         ar="اختيار المواد"),
    dict(no=5, prefix="5", kind=APPEND, optional=False,
         en="Product's Datasheet & Catalogue",
         ar="النشرة الفنية والكتالوج للمنتجات"),
    dict(no=6, prefix="6", kind=APPEND, optional=True,
         en="Vendor Partnership Certificate",
         ar="شهادة شراكة مع المورد"),
    dict(no=7, prefix="7", kind=APPEND, optional=True,
         en="Warranty Certificate",
         ar="شهادة الضمان"),
    dict(no=8, prefix="8", kind=APPEND, optional=False,
         en="Layout and Single Line Diagram",
         ar="مخطط أحادي لجميع الأنظمة"),
]

# Backwards-compatible alias (used by smoke tests / external callers)
SECTIONS = DEFAULT_SECTIONS

# File classification
PDF_EXTS = (".pdf",)
XLSX_EXTS = (".xlsx", ".xls")
DOCX_EXTS = (".docx", ".doc")

# Streams / junk to ignore during discovery
IGNORE_SUFFIXES = (":Zone.Identifier", ":com.dropbox.attributes")
IGNORE_PREFIXES = ("~$", ".~")


# --------------------------------------------------------------------------- #
# Template resolution
# --------------------------------------------------------------------------- #
def _normalize_section(s: dict, idx: int) -> dict:
    no = s.get("no", idx + 1)
    kind = str(s.get("kind", APPEND)).lower()
    if kind not in (TABLE, APPEND):
        raise ValueError(f"Section {no}: kind must be 'table' or 'append', got '{kind}'.")
    if not s.get("en") and not s.get("ar"):
        raise ValueError(f"Section {no}: needs at least an 'en' or 'ar' title.")
    return dict(
        no=no,
        prefix=str(s.get("prefix", no)),
        kind=kind,
        optional=bool(s.get("optional", False)),
        en=s.get("en", ""),
        ar=s.get("ar", ""),
    )


def _load_template_file(ref: str) -> dict:
    """Resolve a template by path or by name under ../templates/."""
    candidates = [ref,
                  os.path.join(TEMPLATES_DIR, ref),
                  os.path.join(TEMPLATES_DIR, ref + ".json")]
    for c in candidates:
        if os.path.isfile(c):
            with open(c, encoding="utf-8") as fh:
                return json.load(fh)
    raise FileNotFoundError(
        f"Template '{ref}' not found (looked in {TEMPLATES_DIR} and as a path).")


def resolve_template(cfg: dict) -> dict:
    """Return {sections, toc_title_en, toc_title_ar, name} from the config.

    Precedence: inline cfg['sections'] > cfg['template'] > built-in default."""
    if cfg.get("sections"):
        raw = {"name": cfg.get("template_name", "custom"),
               "toc_title_en": cfg.get("toc_title_en", TOC_TITLE_EN),
               "toc_title_ar": cfg.get("toc_title_ar", TOC_TITLE_AR),
               "sections": cfg["sections"]}
    elif cfg.get("template"):
        raw = _load_template_file(cfg["template"])
    else:
        raw = {"name": "material-submittal",
               "toc_title_en": TOC_TITLE_EN, "toc_title_ar": TOC_TITLE_AR,
               "sections": DEFAULT_SECTIONS}

    secs = [_normalize_section(s, i) for i, s in enumerate(raw.get("sections", []))]
    if not secs:
        raise ValueError("Template has no sections.")
    # prefixes must be unique so folder discovery is unambiguous
    prefixes = [s["prefix"] for s in secs]
    if len(set(prefixes)) != len(prefixes):
        raise ValueError(f"Section prefixes must be unique, got {prefixes}.")
    return dict(sections=secs,
                toc_title_en=raw.get("toc_title_en", TOC_TITLE_EN),
                toc_title_ar=raw.get("toc_title_ar", TOC_TITLE_AR),
                name=raw.get("name", "custom"))
