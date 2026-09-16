#!/bin/bash
# Open the workbook in Microsoft Excel (macOS), force a full recalculation, save and close,
# so that cached values exist for verification (python verify_excel.py).
set -euo pipefail
FILE="${1:-../excel/Miniproject2_ZCB_Term_Structure.xlsx}"
ABS="$(cd "$(dirname "$FILE")" && pwd)/$(basename "$FILE")"
osascript <<APPLESCRIPT
tell application "Microsoft Excel"
  set display alerts to false
  open POSIX file "$ABS"
  set wb to active workbook
  calculate full
  save wb
  close wb saving no
end tell
APPLESCRIPT
echo "Recalculated and saved: $ABS"
