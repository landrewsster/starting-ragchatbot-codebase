#!/bin/bash
# run_survey.sh
#
# Runs the full survey analysis pipeline:
#   1. survey_frequencies.py  — builds frequency tables → Excel
#   2. survey_hbar_charts.R   — horizontal bar charts for SATA questions
#
# Usage:
#   ./run_survey.sh                        # uses default filename (survey_data.csv)
#   ./run_survey.sh my_data_file.csv       # specify CSV filename

set -e   # stop on first error

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# ── Step 1: frequency tables ───────────────────────────────────────────────────
echo "=== Step 1: survey_frequencies.py ==="
if [ -n "$1" ]; then
    python3 "$SCRIPT_DIR/survey_frequencies.py" "$1"
else
    python3 "$SCRIPT_DIR/survey_frequencies.py"
fi

# ── Step 2: horizontal bar charts ─────────────────────────────────────────────
echo ""
echo "=== Step 2: survey_hbar_charts.R ==="
Rscript "$SCRIPT_DIR/survey_hbar_charts.R"

echo ""
echo "Done."
