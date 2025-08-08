#!/usr/bin/env bash
set -euo pipefail

APP_NAME="program"      # => dist/program
ENTRY="program.py"

rm -rf build dist release
mkdir -p release/data release/config

# (optional) venv
python3 -m venv .venv
source .venv/bin/activate
pip install -U pip wheel setuptools
[ -f requirements.txt ] && pip install -r requirements.txt
pip install pyinstaller

# Build single binary
pyinstaller --noconfirm --clean --onefile "$ENTRY"

# Stage beside the binary
cp "dist/$APP_NAME"* "release/"

# Config: copy if present, else create placeholder so folder is kept
if [[ -f config/config.example.ini ]]; then
  cp config/config.example.ini "release/config/config.ini"
else
  : > "release/config/.keep"
fi

# Data: you said there's nothing to copy—just ensure folder exists in artifacts
: > "release/data/.keep"

# Zip it (placeholders ensure empty dirs are included)
cd release
zip -r "../${APP_NAME}_$(date +%Y%m%d)_macos.zip" .