# amt-pas-compiler

Compile an AMT Technical Submittal / PAS (Product/Material Approval Submittal) into
one branded PDF — cover, bilingual table of contents, section dividers, rendered BOQ
tables, appended datasheets, warranty and drawings — from a standard numbered input
folder. See `SKILL.md` for full usage.

## Install (per machine)
1. Put the skill where Claude Code looks for it — `~/.claude/skills/amt-pas-compiler/`.

   **From GitHub (recommended):**
   ```bash
   git clone https://github.com/airandblueamt-Claude/amt-pas-compiler.git \
     ~/.claude/skills/amt-pas-compiler
   cd ~/.claude/skills/amt-pas-compiler
   ```
   Or, if you already have the folder, just copy it into `~/.claude/skills/`.
2. From inside the folder, run:
   ```bash
   bash install.sh
   ```
   This installs the Python dependencies and checks for LibreOffice.
3. (Recommended) Install LibreOffice for pixel-faithful Excel/Word/Arabic rendering:
   - Ubuntu/Debian: `sudo apt-get install -y libreoffice-calc libreoffice-writer`
   - macOS: `brew install --cask libreoffice`
   - No sudo? Point the skill at a portable build with `export AMT_SOFFICE=/path/to/soffice`.

Requires Python 3.9+. Without LibreOffice the skill still runs using a built-in
renderer, but the bilingual tables are lower fidelity.

## Use
```bash
cp config.example.json my.config.json        # edit ref no., title, client, paths
python3 scripts/build_pas.py my.config.json --dry-run   # validate the input folder
python3 scripts/build_pas.py my.config.json             # build the submittal PDF
```
In Claude Code you can also just ask: "build the PAS for <project>" and the skill
triggers automatically. For the exact prompts that make it build cleanly first time,
see **[PROMPTING.md](PROMPTING.md)**.

## Input folder
Sub-folders matched by their leading number (1–8); as-received name typos are fine:
`1-Tender BOQ`, `2-AMT-VENDOR BOQ`, … `8-OVERALL Single line digram`.

## Templates
The section list is config-driven — see `templates/README.md`. The default is the
8-section material submittal; add a JSON for other submittal types.
