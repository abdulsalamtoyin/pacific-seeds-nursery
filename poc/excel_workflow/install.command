#!/usr/bin/env bash
# Pacific Seeds — assist with VBA install on macOS Excel.
#
# Mac Excel does not allow EXTERNAL tools to inject VBA modules (Windows only).
# So this installer prepares the .xlsx files, opens them in Excel, and copies a
# tiny bootstrap macro to your clipboard. Paste it once into each workbook and
# press F5 — it installs the rest from disk automatically and deletes itself.

set -u

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$HERE/.." && pwd)"
VENV_PY="$PROJECT_ROOT/.venv/bin/python"
VBA_DIR="$HERE/vba"
OUT_DIR="$HERE/output"
BOOTSTRAP="$VBA_DIR/bootstrap.bas"

echo "──────────────────────────────────────────────"
echo "  Pacific Seeds — Nursery VBA installer"
echo "──────────────────────────────────────────────"
echo

if [[ ! -x "$VENV_PY" ]]; then
  echo "❌ Python venv not found at $VENV_PY"
  read -n 1 -s -r -p "Press any key to close…"
  exit 1
fi

echo "→ Regenerating .xlsx skeletons…"
"$VENV_PY" "$HERE/build_workbooks.py" || { echo "❌ build failed"; exit 2; }
echo

if [[ ! -f "$BOOTSTRAP" ]]; then
  echo "❌ Missing $BOOTSTRAP"; exit 3
fi

# Copy the bootstrap to the clipboard, stripping the leading
# `Attribute VB_Name = "..."` line (valid in .bas files, but VBA refuses
# to compile it when pasted into a module by hand).
grep -v '^Attribute VB_Name' "$BOOTSTRAP" | pbcopy
echo "✅ Bootstrap copied to your clipboard ($(grep -vc '^Attribute VB_Name' "$BOOTSTRAP") lines)."
echo

cat <<INSTR

╔══════════════════════════════════════════════════════════════════════════╗
║  Two workbooks need this — repeat the steps below for each.              ║
╚══════════════════════════════════════════════════════════════════════════╝

ONE-TIME (if you haven't already):
  Excel ▸ Preferences ▸ Security & Privacy ▸ Macro Security
  ✓ Trust access to the VBA project object model

FOR EACH WORKBOOK (Nursery_Template.xlsx, then Nursery_Hub.xlsx):

  1.  Tools ▸ Macros ▸ Visual Basic Editor          (or  ⌥F11)
  2.  Insert ▸ Module
  3.  ⌘V   (the bootstrap is already on your clipboard)
  4.  F5
        ─ pick the folder:  $VBA_DIR
        ─ confirm any dialog
        ─ the bootstrap deletes itself when done
  5.  File ▸ Save As… ▸ Excel Macro-Enabled Workbook (.xlsm)
        save next to the .xlsx in: $OUT_DIR

After step 5 of the Template, the second pass for the Hub is the same
steps — the clipboard still has the bootstrap.

INSTR

# Open both .xlsx files in Excel so they're ready.
echo "→ Opening both .xlsx files in Excel now…"
open -a "Microsoft Excel" "$OUT_DIR/Nursery_Template.xlsx" 2>/dev/null
sleep 1
open -a "Microsoft Excel" "$OUT_DIR/Nursery_Hub.xlsx" 2>/dev/null
echo
echo "(Excel should now be in front. The bootstrap is on your clipboard.)"
echo
read -n 1 -s -r -p "Press any key to close this window…"
echo
