#!/usr/bin/env bash
# Đưa bản mới của máy chủ (kho + xuất PPTX trên web) lên VPS 360 — chạy từ thư mục BD.
#   bash server/deploy.sh
# Cần SSH key ~/.ssh/ga360_deploy (root@103.153.74.177).
set -euo pipefail
HOST=${AUTOPPTX_HOST:-root@103.153.74.177}
KEY=${AUTOPPTX_KEY:-$HOME/.ssh/ga360_deploy}
cd "$(dirname "$0")/.."
SSH=(ssh -i "$KEY" -o BatchMode=yes "$HOST")
"${SSH[@]}" 'rm -rf /tmp/autopptx-server && mkdir -p /tmp/autopptx-server/autopptx2'
scp -q -i "$KEY" -o BatchMode=yes \
  server/autopptx_server.py server/render_job.py server/admin.html server/report.html server/editor.html \
  server/setup.sh server/requirements.txt server/autopptx-store.service server/nginx-autopptx.conf \
  "$HOST:/tmp/autopptx-server/"
"${SSH[@]}" 'mkdir -p /tmp/autopptx-server/static'
scp -q -i "$KEY" -o BatchMode=yes server/static/* "$HOST:/tmp/autopptx-server/static/"
# Đúng code xuất của app (đọc Excel, khớp mã/avatar, dựng PPTX) — không kèm giao diện Tk.
scp -q -i "$KEY" -o BatchMode=yes \
  autopptx2/__init__.py autopptx2/constants.py autopptx2/datasource.py \
  autopptx2/effects.py autopptx2/exporter.py autopptx2/fonts.py autopptx2/geometry.py \
  autopptx2/qa.py autopptx2/pipeline.py \
  "$HOST:/tmp/autopptx-server/autopptx2/"
"${SSH[@]}" 'bash /tmp/autopptx-server/setup.sh'
