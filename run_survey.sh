#!/bin/bash
# run_survey.sh
#
# Runs the full survey analysis pipeline.
# Place this script in: ~/Downloads/CRC MDH Project/MDH analysis/
#
# Usage:
#   ./run_survey.sh                  # uses the default CSV filename below
#   ./run_survey.sh my_new_file.csv  # override with a different filename

set -e   # stop on first error

DIR="$(cd "$(dirname "$0")" && pwd)"
CSV="${1:-MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-06-01_1832_EDITED.csv}"

echo "=== Step 1: survey_frequencies.py ($CSV) ==="
python3 "$DIR/survey_frequencies.py" "$CSV"

echo ""
echo "=== Step 2: survey_likert_chart.R ==="
Rscript "$DIR/survey_likert_chart.R"

echo ""
echo "=== Step 3: survey_hbar_charts.R ==="
Rscript "$DIR/survey_hbar_charts.R"

echo ""
echo "=== Step 4: survey_report.R ==="
Rscript "$DIR/survey_report.R"

echo ""
echo "Done."
