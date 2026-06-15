"""
discover.py — locate and validate the 8 input sections inside a submittal folder.

Folders are matched by their leading number (1..8) so renamed / typo'd folder
names still resolve. Returns a structured manifest and a human-readable report.
"""
from __future__ import annotations

import os
import glob

from pas_spec import (DEFAULT_SECTIONS, TABLE, APPEND,
                      PDF_EXTS, XLSX_EXTS, DOCX_EXTS,
                      IGNORE_SUFFIXES, IGNORE_PREFIXES)


def _is_junk(name: str) -> bool:
    if any(name.endswith(s) for s in IGNORE_SUFFIXES):
        return True
    if any(name.startswith(p) for p in IGNORE_PREFIXES):
        return True
    return False


def _find_section_dir(input_dir: str, prefix: str) -> str | None:
    """Find the sub-folder whose name starts with '<prefix>-' or '<prefix> '."""
    for entry in sorted(os.listdir(input_dir)):
        full = os.path.join(input_dir, entry)
        if not os.path.isdir(full):
            continue
        # match "1-...", "1 ...", "1_..." or exactly "1"
        if entry == prefix or entry[:len(prefix) + 1] in (prefix + "-", prefix + " ", prefix + "_"):
            return full
    return None


def _collect_files(root: str, exts: tuple[str, ...]) -> list[str]:
    """Recursively collect files with the given extensions, skipping junk."""
    out = []
    for dirpath, _dirs, files in os.walk(root):
        for f in files:
            if _is_junk(f):
                continue
            if f.lower().endswith(exts):
                out.append(os.path.join(dirpath, f))
    return out


def _ordered(paths: list[str]) -> list[str]:
    """Stable ordering: shallower paths first, then case-insensitive name,
    with 'master' plans before 'part' plans (matches the sample's drawing order)."""
    def key(p):
        rel = p.lower()
        depth = rel.count(os.sep)
        master_bias = 0 if "master" in os.path.basename(rel) else 1
        return (depth, master_bias, os.path.basename(rel))
    return sorted(paths, key=key)


def discover(input_dir: str, sections: list | None = None) -> dict:
    """Build the manifest for every section. `sections` defaults to the built-in
    template; pass a resolved list (from pas_spec.resolve_template) for custom ones."""
    input_dir = os.path.abspath(input_dir)
    if not os.path.isdir(input_dir):
        raise FileNotFoundError(f"Input folder not found: {input_dir}")
    if sections is None:
        sections = DEFAULT_SECTIONS

    manifest = {"input_dir": input_dir, "sections": [], "warnings": [], "errors": []}

    for spec in sections:
        sec = dict(no=spec["no"], en=spec["en"], ar=spec["ar"],
                   kind=spec["kind"], optional=spec["optional"],
                   dir=None, xlsx=None, xlsx_list=[], pdfs=[], docx=[], status="ok")
        sdir = _find_section_dir(input_dir, spec["prefix"])
        sec["dir"] = sdir

        if sdir is None:
            sec["status"] = "missing-folder"
            msg = f"Section {spec['no']} ({spec['en']}): folder not found."
            (manifest["warnings"] if spec["optional"] else manifest["errors"]).append(msg)
            manifest["sections"].append(sec)
            continue

        # Collect EVERY supported file type regardless of the section's declared
        # kind. Upload format must not matter: spreadsheets are rendered as tables,
        # PDF/Word are appended — so a section provided either way always lays out
        # well (#4 "some docs as PDF, some as Excel; layout always good").
        xlsx = _collect_files(sdir, XLSX_EXTS)
        # spreadsheets sitting directly in the section folder come first
        xlsx.sort(key=lambda p: (os.path.dirname(p) != sdir, os.path.basename(p).lower()))
        sec["xlsx_list"] = xlsx
        sec["xlsx"] = xlsx[0] if xlsx else None
        sec["pdfs"] = _ordered(_collect_files(sdir, PDF_EXTS))
        sec["docx"] = _ordered(_collect_files(sdir, DOCX_EXTS))

        if not (xlsx or sec["pdfs"] or sec["docx"]):
            sec["status"] = "empty"
            if spec["optional"]:
                manifest["warnings"].append(
                    f"Section {spec['no']} ({spec['en']}): no documents found — "
                    f"a placeholder will be inserted.")
            else:
                manifest["errors"].append(
                    f"Section {spec['no']} ({spec['en']}): no documents found — "
                    f"required section is empty.")

        manifest["sections"].append(sec)

    return manifest


def format_report(manifest: dict) -> str:
    lines = [f"Input: {manifest['input_dir']}", ""]
    for sec in manifest["sections"]:
        tag = f"  {sec['no']}. {sec['en']}"
        parts = []
        if sec.get("xlsx_list"):
            parts.append(f"{len(sec['xlsx_list'])} table")
        if sec["pdfs"]:
            parts.append(f"{len(sec['pdfs'])} pdf")
        if sec["docx"]:
            parts.append(f"{len(sec['docx'])} docx")
        detail = ", ".join(parts) if parts else "— empty (placeholder) —"
        lines.append(f"{tag:<58} {detail}")
    if manifest["warnings"]:
        lines += ["", "Warnings:"] + [f"  ! {w}" for w in manifest["warnings"]]
    if manifest["errors"]:
        lines += ["", "ERRORS:"] + [f"  x {e}" for e in manifest["errors"]]
    return "\n".join(lines)


if __name__ == "__main__":
    import sys
    m = discover(sys.argv[1])
    print(format_report(m))
