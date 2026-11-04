#!/usr/bin/env bash
# Đưa bản cài app máy tính lên VPS 360 → nút "Windows / macOS" ở trang đăng nhập web.
#   bash server/publish_downloads.sh 2.1.0 file1.exe file2.zip ...
# Tên file quyết định hệ điều hành: *.exe/*Windows* → Windows; *Universal* → macOS;
# *AppleSilicon*/*arm64* → macOS Apple Silicon; *Intel* → macOS Intel.
# Bản cũ trong thư mục bị xoá, chỉ giữ đúng các file vừa đưa lên.
set -euo pipefail
HOST=${AUTOPPTX_HOST:-root@103.153.74.177}
KEY=${AUTOPPTX_KEY:-$HOME/.ssh/ga360_deploy}
VER=${1:?Cần số phiên bản, vd 2.1.0}; shift
[ $# -gt 0 ] || { echo "Cần ít nhất 1 file"; exit 2; }
SSH=(ssh -i "$KEY" -o BatchMode=yes "$HOST")
"${SSH[@]}" 'rm -rf /tmp/autopptx-dl && mkdir -p /tmp/autopptx-dl'
for f in "$@"; do
  [ -f "$f" ] || { echo "Không thấy $f"; exit 2; }
  echo "  tải lên $(basename "$f") ($(du -h "$f" | cut -f1))"
  scp -q -i "$KEY" -o BatchMode=yes "$f" "$HOST:/tmp/autopptx-dl/"
done
"${SSH[@]}" "set -e
D=/var/lib/autopptx/store/downloads
install -d -o autopptx -g www-data -m 2750 \$D
find \$D -maxdepth 1 -type f -delete
for f in /tmp/autopptx-dl/*; do install -o autopptx -g www-data -m 0640 \"\$f\" \$D/; done
printf '{\"version\": \"%s\"}\n' '$VER' > \$D/downloads.json
chown autopptx:www-data \$D/downloads.json; chmod 0640 \$D/downloads.json
rm -rf /tmp/autopptx-dl
ls -la \$D"
curl -fsS https://talent.goldenasia.vn/pptx/api/downloads | head -c 600; echo
