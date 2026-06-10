#!/usr/bin/env python3
"""
build_pas.py — CLI entry point for the AMT PAS (Technical Submittal) compiler.

Usage:
    python3 build_pas.py <config.json>
    python3 build_pas.py <config.json> --dry-run      # discovery report only
    python3 build_pas.py <config.json> --open         # print path to open after build

The config drives everything; see config.example.json and SKILL.md.
"""
from __future__ import annotations

import os
import sys
import json
import argparse

import discover as DISC
import assemble as ASM
import pas_spec as SPEC

REQUIRED = ["ref_no", "version", "date", "project_title_en", "project_title_ar",
            "signoff", "revision", "input_dir", "output_pdf"]
SIGNOFF_ROLES = ["prepared_by", "checked_by", "approved_by"]


def validate_config(cfg: dict) -> list[str]:
    errs = []
    for k in REQUIRED:
        if k not in cfg or cfg[k] in (None, ""):
            errs.append(f"missing config key: {k}")
    if "signoff" in cfg:
        for r in SIGNOFF_ROLES:
            if r not in cfg["signoff"]:
                errs.append(f"signoff.{r} missing")
            else:
                for f in ("role_en", "initials"):
                    if f not in cfg["signoff"][r]:
                        errs.append(f"signoff.{r}.{f} missing")
    return errs


def main(argv=None):
    ap = argparse.ArgumentParser(description="Compile an AMT PAS submittal PDF.")
    ap.add_argument("config", help="path to the submittal config JSON")
    ap.add_argument("--dry-run", action="store_true",
                    help="show the discovery/validation report and exit")
    ap.add_argument("--open", action="store_true",
                    help="print the output path on success")
    args = ap.parse_args(argv)

    with open(args.config, encoding="utf-8") as fh:
        cfg = json.load(fh)

    cfg_errs = validate_config(cfg)
    if cfg_errs:
        print("Config errors:")
        for e in cfg_errs:
            print(f"  x {e}")
        return 2

    try:
        template = SPEC.resolve_template(cfg)
    except (ValueError, FileNotFoundError) as e:
        print(f"Template error: {e}")
        return 2
    print(f"Template: {template['name']} ({len(template['sections'])} sections)")

    manifest = DISC.discover(cfg["input_dir"], template["sections"])
    print(DISC.format_report(manifest))
    print()

    if manifest["errors"]:
        print("Cannot build — required sections are missing (see ERRORS above).")
        return 3

    if args.dry_run:
        print("Dry run — no PDF written.")
        return 0

    qa = ASM.build(manifest, cfg, template)

    print()
    print("=" * 60)
    print(f"Output : {qa['output']}")
    print(f"Pages  : {qa['actual_pages']} (expected {qa['expected_pages']}) "
          f"-> {'OK' if qa['consistent'] else 'MISMATCH!'}")
    print(f"Engine : {', '.join(qa['engines'])}")
    print("TOC page map:")
    for no, pg in sorted(qa["page_map"].items()):
        print(f"   §{no} -> p{pg}")
    if "reportlab" in qa["engines"] or "reportlab-text" in qa["engines"]:
        print()
        print("Note: LibreOffice was not used for some sections (reportlab fallback).")
        print("      For pixel-faithful Excel/Word rendering install LibreOffice")
        print("      (sudo apt-get install -y libreoffice-calc libreoffice-writer)")
        print("      and re-run; the skill auto-detects it.")
    if args.open:
        print(qa["output"])
    return 0 if qa["consistent"] else 4


if __name__ == "__main__":
    sys.exit(main())
