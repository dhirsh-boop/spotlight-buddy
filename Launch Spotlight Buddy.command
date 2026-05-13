#!/bin/bash
# One-click launcher for Spotlight Buddy.
# Double-click this file in Finder to start the Streamlit app.

set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

PYBIN="venv/bin/python"

# (Re)create the venv if missing or broken.
if [ ! -x "$PYBIN" ] || ! "$PYBIN" -c "import sys" >/dev/null 2>&1; then
    echo "Creating fresh virtual environment..."
    rm -rf venv
    python3 -m venv venv
fi

echo "Ensuring dependencies are installed..."
"$PYBIN" -m pip install --quiet --upgrade pip
"$PYBIN" -m pip install --quiet -r requirements.txt

echo ""
echo "=========================================="
echo "  Spotlight Buddy is starting..."
echo "  A browser tab will open automatically."
echo "  Close this Terminal window to stop."
echo "=========================================="
echo ""

"$PYBIN" -m streamlit run main.py
