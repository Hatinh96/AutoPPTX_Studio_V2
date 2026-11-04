#!/usr/bin/env bash
# Đưa bản mới của máy chủ kho lên VPS 360 (chạy từ thư mục BD, Git Bash / macOS).
#   bash server/deploy.sh
# Cần SSH key ~/.ssh/ga360_deploy (root@103.153.74.177).
set -euo pipefail
HOST=${AUTOPPTX_HOST:-root@103.153.74.177}
KEY=${AUTOPPTX_KEY:-$HOME/.ssh/ga360_deploy}
cd "$(dirname "$0")/.."
SSH=(ssh -i "$KEY" -o BatchMode=yes "$HOST")
"${SSH[@]}" 'rm -rf /tmp/autopptx-server && mkdir -p /tmp/autopptx-server/autopptx2'
scp -q -i "$KEY" -o BatchMode=yes \
  server/autopptx_server.py server/admin.html server/setup.sh \
  server/autopptx-store.service server/nginx-autopptx.conf \
  "$HOST:/tmp/autopptx-server/"
# Đối chiếu mã RP dùng đúng code của app (đọc Excel + khớp avatar).
scp -q -i "$KEY" -o BatchMode=yes \
  autopptx2/__init__.py autopptx2/constants.py autopptx2/datasource.py \
  "$HOST:/tmp/autopptx-server/autopptx2/"
"${SSH[@]}" 'bash /tmp/autopptx-server/setup.sh'
