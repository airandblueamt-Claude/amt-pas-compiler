#!/usr/bin/env bash
# One-step setup for the amt-pas-compiler skill.
# Run from inside the skill folder:  bash install.sh
set -e

HERE="$(cd "$(dirname "$0")" && pwd)"
echo "Installing Python dependencies for amt-pas-compiler …"
python3 -m pip install --user --break-system-packages -r "$HERE/requirements.txt" \
  || python3 -m pip install --user -r "$HERE/requirements.txt"

echo
echo "Checking for LibreOffice (needed for pixel-faithful Excel/Word rendering) …"
if command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1; then
  echo "  ✓ LibreOffice found."
else
  echo "  ! LibreOffice not found. The skill still runs (reportlab fallback), but for"
  echo "    pixel-faithful tables/Arabic install it once:"
  echo "      Ubuntu/Debian : sudo apt-get install -y libreoffice-calc libreoffice-writer"
  echo "      macOS (brew)  : brew install --cask libreoffice"
  echo "      Windows       : install LibreOffice, ensure soffice.exe is on PATH"
  echo "    No sudo? Set AMT_SOFFICE=/path/to/soffice (a portable/AppImage build works)."
fi

echo
echo "Verifying the skill imports cleanly …"
python3 - "$HERE" <<'PY'
import sys, os
sys.path.insert(0, os.path.join(sys.argv[1], "scripts"))
import build_pas, pas_spec, render_tables, normalize  # noqa
t = pas_spec.resolve_template({})
print(f"  ✓ OK — default template '{t['name']}' with {len(t['sections'])} sections.")
PY

echo
echo "Done. Quick start:"
echo "  cp $HERE/config.example.json my.config.json   # edit ref/title/client/paths"
echo "  python3 $HERE/scripts/build_pas.py my.config.json --dry-run   # validate inputs"
echo "  python3 $HERE/scripts/build_pas.py my.config.json             # build the PDF"
