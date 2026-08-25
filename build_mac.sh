#!/bin/bash
# Build script for macOS (.app bundle & zip archive)
echo "Building AutoPPTX Studio V2 for macOS..."

python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
python3 -m pip install pyinstaller

python3 -m PyInstaller --noconfirm AutoPPTX_Studio_V2.spec

if [ -d "dist/AutoPPTX_Studio_V2.app" ]; then
    echo "Compressing .app bundle to dist/AutoPPTX_Studio_V2-macOS.zip..."
    ditto -c -k --sequesterRsrc --keepParent "dist/AutoPPTX_Studio_V2.app" "dist/AutoPPTX_Studio_V2-macOS.zip"
    echo "Build complete! Output: dist/AutoPPTX_Studio_V2-macOS.zip"
else
    echo "Build complete! Output located in dist/"
fi


