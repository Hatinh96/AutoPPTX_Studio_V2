#!/bin/bash
# Build script for macOS (.app / binary)
echo "Building AutoPPTX Studio V2 for macOS..."

python3 -m pip install -r requirements.txt
python3 -m pip install pyinstaller

python3 -m PyInstaller --noconsole --onefile --name AutoPPTX_Studio_V2 --collect-all customtkinter AutoPPTX_Studio_V2.py

echo "Build complete! Output located in dist/AutoPPTX_Studio_V2"
