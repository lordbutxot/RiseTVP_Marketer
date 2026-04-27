#!/bin/bash
cd "$(dirname "$0")"

if [ -f ".venv/bin/python" ]; then
    PYTHON=".venv/bin/python"
else
    PYTHON="python3"
fi

echo ""
echo "===================================="
echo "  RISE TVP TRADE ROUTE OPTIMIZER"
echo "===================================="
echo ""
echo "Select mode:"
echo "1. Regular Trade (all opportunities)"
echo "2. City-Specific Opportunities"
echo "3. Trade Route (multi-hop chained routes)"
read -p "Enter 1, 2 or 3: " mode

if [ "$mode" == "1" ]; then
    "$PYTHON" "ocr_to_excel.py" --mode regular "$@"
elif [ "$mode" == "2" ]; then
    "$PYTHON" "ocr_to_excel.py" --mode city "$@"
elif [ "$mode" == "3" ]; then
    "$PYTHON" "ocr_to_excel.py" --mode route "$@"
else
    echo "Invalid choice."
    exit 1
fi

read -p "Press Enter to continue..."
