#!/bin/bash
# Build script for macOS (.app bundle)
echo "Building AutoPPTX Studio V2 for macOS..."

python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
python3 -m pip install pyinstaller

python3 -m PyInstaller --noconfirm AutoPPTX_Studio_V2.spec

echo "Build complete! Output located in dist/AutoPPTX_Studio_V2.app or dist/AutoPPTX_Studio_V2"

