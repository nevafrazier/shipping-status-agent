#!/bin/bash
# Double-click this file to run the daily tracking report.
# It picks up the most recent .xlsx in this folder automatically.

cd "$(dirname "$0")"

echo "================================================"
echo "  Daily Shipping Tracking Report"
echo "  $(date '+%A, %B %d %Y')"
echo "================================================"
echo ""

python3 scrape_tracking.py

echo ""
echo "================================================"
echo "  Done! You can close this window."
echo "================================================"
read -p "Press Enter to close..."
