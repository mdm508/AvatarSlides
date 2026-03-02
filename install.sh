#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="${1:-$HOME/anki-slides}"

mkdir -p "$PROJECT_DIR"
cd "$PROJECT_DIR"

python3 -m venv .venv
source .venv/bin/activate

python -m pip install --upgrade pip
pip install python-pptx pillow

echo
echo "Done."
echo "Project directory: $PROJECT_DIR"
echo
echo "To activate later:"
echo "  cd \"$PROJECT_DIR\""
echo "  source .venv/bin/activate"
echo
echo "Optional for PDF conversion:"
echo "  brew install --cask libreoffice"