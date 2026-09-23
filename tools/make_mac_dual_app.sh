#!/bin/bash
# Gộp 2 bản .app (arm64 + x86_64) thành MỘT .app chạy được trên cả hai loại Mac.
#
# Cách hoạt động: .app ngoài chỉ chứa 1 script launcher; script đọc `uname -m`
# rồi chạy đúng bản bên trong Contents/Resources/<arch>/.
# Không dùng lipo để ghép 2 file PyInstaller onefile — bản ghép kiểu đó chỉ chạy
# được 1 chip vì PKG archive nhúng bên trong chỉ có của 1 slice.
#
# Dùng: tools/make_mac_dual_app.sh <arm64.zip> <x86_64.zip> [thư mục xuất]
set -euo pipefail

ARM_ZIP="${1:?thiếu file zip bản arm64 (Apple Silicon)}"
X86_ZIP="${2:?thiếu file zip bản x86_64 (Intel)}"
OUT_DIR="${3:-release_out}"

APP_NAME="AutoPPTX_Studio_V2"
BIN_NAME="AutoPPTX_Studio_V2"
BUNDLE_ID="com.autopptx.studio"
VERSION="${APP_VERSION:-2.4.0}"
VERSION="${VERSION#v}"

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

extract() {                       # $1 = zip, $2 = thư mục đích
    mkdir -p "$2"
    ditto -x -k "$1" "$2"
}

find_app() {                      # tìm .app trong thư mục vừa giải nén
    find "$1" -maxdepth 3 -name "$APP_NAME.app" -type d | head -n 1
}

check_arch() {                    # $1 = .app, $2 = arch bắt buộc phải có
    local exe="$1/Contents/MacOS/$BIN_NAME"
    local got
    got="$(lipo -archs "$exe" 2>/dev/null || true)"
    if ! printf '%s\n' $got | grep -qx "$2"; then
        echo "LỖI: $exe không có slice '$2' (đang có: ${got:-không đọc được})." >&2
        exit 1
    fi
}

echo "Giải nén 2 bản…"
extract "$ARM_ZIP" "$WORK/arm64"
extract "$X86_ZIP" "$WORK/x86_64"

ARM_APP="$(find_app "$WORK/arm64")"
X86_APP="$(find_app "$WORK/x86_64")"
[ -n "$ARM_APP" ] && [ -d "$ARM_APP" ] || { echo "Không thấy $APP_NAME.app trong $ARM_ZIP" >&2; exit 1; }
[ -n "$X86_APP" ] && [ -d "$X86_APP" ] || { echo "Không thấy $APP_NAME.app trong $X86_ZIP" >&2; exit 1; }

# Chặn trường hợp lấy trùng 2 bản cùng chip — lỗi này im lặng rất khó tìm.
check_arch "$ARM_APP" arm64
check_arch "$X86_APP" x86_64

DEST="$WORK/out/$APP_NAME.app"
mkdir -p "$DEST/Contents/MacOS" "$DEST/Contents/Resources"

echo "Chép payload…"
ditto "$ARM_APP" "$DEST/Contents/Resources/arm64/$APP_NAME.app"
ditto "$X86_APP" "$DEST/Contents/Resources/x86_64/$APP_NAME.app"

# Icon cho .app ngoài: ưu tiên file trong repo, không có thì lấy từ bản arm64.
ICON_NAME=""
if [ -f "app_icon.icns" ]; then
    cp "app_icon.icns" "$DEST/Contents/Resources/app_icon.icns"
    ICON_NAME="app_icon"
else
    ICNS="$(find "$ARM_APP/Contents/Resources" -maxdepth 1 -name '*.icns' | head -n 1)"
    if [ -n "$ICNS" ]; then
        cp "$ICNS" "$DEST/Contents/Resources/$(basename "$ICNS")"
        ICON_NAME="$(basename "$ICNS" .icns)"
    fi
fi

cat > "$DEST/Contents/MacOS/launch" <<'LAUNCHER'
#!/bin/bash
# Chạy đúng bản theo chip: arm64 (M1/M2/M3+) hoặc x86_64 (Intel).
set -u
HERE="$(cd "$(dirname "$0")/../Resources" && pwd)"
APP="AutoPPTX_Studio_V2"
BIN="$HERE/$(uname -m)/$APP.app/Contents/MacOS/$APP"
if [ ! -x "$BIN" ]; then
    for alt in arm64 x86_64; do
        if [ -x "$HERE/$alt/$APP.app/Contents/MacOS/$APP" ]; then
            BIN="$HERE/$alt/$APP.app/Contents/MacOS/$APP"
            break
        fi
    done
fi
if [ ! -x "$BIN" ]; then
    osascript -e 'display alert "AutoPPTX Studio V2" message "Bản cài thiếu tệp chạy cho chip của máy này."' >/dev/null 2>&1 || true
    exit 1
fi
exec "$BIN" "$@"
LAUNCHER
chmod +x "$DEST/Contents/MacOS/launch"

ICON_ENTRY=""
if [ -n "$ICON_NAME" ]; then
    ICON_ENTRY=$(printf '\t<key>CFBundleIconFile</key>\n\t<string>%s</string>' "$ICON_NAME")
fi

cat > "$DEST/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>CFBundleName</key>
	<string>AutoPPTX Studio V2</string>
	<key>CFBundleDisplayName</key>
	<string>AutoPPTX Studio V2</string>
	<key>CFBundleExecutable</key>
	<string>launch</string>
	<key>CFBundleIdentifier</key>
	<string>$BUNDLE_ID</string>
	<key>CFBundlePackageType</key>
	<string>APPL</string>
	<key>CFBundleShortVersionString</key>
	<string>$VERSION</string>
	<key>CFBundleVersion</key>
	<string>$VERSION</string>
$ICON_ENTRY
	<key>NSHighResolutionCapable</key>
	<true/>
</dict>
</plist>
PLIST
plutil -lint "$DEST/Contents/Info.plist"

xattr -cr "$DEST" 2>/dev/null || true

mkdir -p "$OUT_DIR"
ZIP="$OUT_DIR/$APP_NAME-macOS-Universal.zip"
rm -f "$ZIP"
ditto -c -k --sequesterRsrc --keepParent "$DEST" "$ZIP"

echo "Xong: $ZIP"
du -sh "$ZIP" | awk '{print "Dung lượng zip: "$1}'
