#!/bin/bash
echo "===================================="
echo "  RISE TVP TRADE ROUTE OPTIMIZER"
echo "===================================="
echo ""
echo "Select mode:"
echo "1. Regular Trade (all opportunities)"
echo "2. City-Specific Opportunities"
read -p "Enter 1 or 2: " mode

if [ "$mode" == "1" ]; then
    python3 ocr_to_excel.py --mode regular
elif [ "$mode" == "2" ]; then
    python3 ocr_to_excel.py --mode city
else
    echo "Invalid choice."
fi