#!/bin/bash
# Excel Updater launcher (Mac / Linux)
cd "$(dirname "$0")"
command -v python3 >/dev/null 2>&1 || { echo "ERROR: Python 3 not found."; exit 1; }
read -p "Excel file name (e.g. MyWorkbook.xlsx): " WB
[ -f "$WB" ] || { echo "File not found: $WB"; exit 1; }
read -p "Unique id (e.g. BRP1E): " UID_IN
[ -n "$UID_IN" ] || { echo "No unique id given."; exit 1; }
echo; echo "=== STEP 1: PREVIEW for $UID_IN (nothing is saved) ==="
python3 excel_updater.py --workbook "$WB" --uid "$UID_IN" --config template_config.json --dry-run || exit 1
echo; read -p "Look OK? Type Y to write the output file: " GO
[ "$GO" = "Y" ] || [ "$GO" = "y" ] || { echo "Cancelled."; exit 0; }
python3 excel_updater.py --workbook "$WB" --uid "$UID_IN" --config template_config.json --output "UPDATED_$WB"
echo "Done. Your original file was not modified."
