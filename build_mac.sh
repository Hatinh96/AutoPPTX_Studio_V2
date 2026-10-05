#!/bin/bash
# Build script for macOS (.app bundle & zip archive)
# Bản .app chỉ chạy đúng chip của máy đang build:
#   máy Intel  -> x86_64  (chạy trên Mac Intel)
#   máy M1/M2+ -> arm64   (KHÔNG chạy trên Mac Intel)
ARCH="$(uname -m)"
case "$ARCH" in
    x86_64) ARCH_LABEL="Intel" ;;
    arm64)  ARCH_LABEL="AppleSilicon" ;;
    *)      ARCH_LABEL="$ARCH" ;;
esac
echo "Building AutoPPTX Studio V2 for macOS ($ARCH -> $ARCH_LABEL)..."

python3 -m pip install --upgrade pip
# Mac Intel: cryptography >= 49 không còn wheel -> build từ nguồn sẽ link OpenSSL
# Homebrew và xung đột libssl.3.dylib khi đóng gói. Chỉ nhận wheel.
python3 -m pip install --only-binary=cryptography -r requirements.txt
python3 -m pip install pyinstaller

python3 -m PyInstaller --noconfirm AutoPPTX_Studio_V2.spec

"dist/AutoPPTX_Studio_V2.app/Contents/MacOS/AutoPPTX_Studio_V2" --selftest || {
    echo "Self-test FAIL — xem dòng FAIL ở trên trước khi phát hành."
    exit 1
}

ZIP_OUT="dist/AutoPPTX_Studio_V2-macOS-$ARCH_LABEL.zip"

if [ -d "dist/AutoPPTX_Studio_V2.app" ]; then
    echo "Removing quarantine xattr (Gatekeeper)..."
    xattr -cr "dist/AutoPPTX_Studio_V2.app" 2>/dev/null || true
    echo "Compressing .app bundle to $ZIP_OUT..."
    ditto -c -k --sequesterRsrc --keepParent "dist/AutoPPTX_Studio_V2.app" "$ZIP_OUT"
    echo "Build complete! Output: $ZIP_OUT"
else
    echo "Build complete! Output located in dist/"
fi


