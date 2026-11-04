#!/usr/bin/env bash
# Chạy trên VPS (root). Cài / cập nhật dịch vụ AutoPPTX kho tài nguyên.
# Gọi bởi server/deploy.sh sau khi chép bộ cài vào /tmp/autopptx-server.
set -euo pipefail
SRC=/tmp/autopptx-server
APP=/opt/autopptx-server
DATA=/var/lib/autopptx
SITE=/etc/nginx/sites-available/ga360

echo "== 1/5 gói hệ thống"
need=()
python3 -c 'import PIL' 2>/dev/null || need+=(python3-pil)
python3 -c 'import openpyxl' 2>/dev/null || need+=(python3-openpyxl)
if [ ${#need[@]} -gt 0 ]; then
  DEBIAN_FRONTEND=noninteractive apt-get install -y -qq "${need[@]}" >/dev/null
fi

echo "== 2/5 tài khoản hệ thống + thư mục"
id autopptx >/dev/null 2>&1 || useradd --system --home-dir "$DATA" --shell /usr/sbin/nologin autopptx
install -d -m 755 "$APP" "$APP/autopptx2"
# DB/khóa chỉ autopptx đọc; nginx (www-data) chỉ đọc được store/ để gửi file.
install -d -o autopptx -g www-data -m 0710 "$DATA"
install -d -o autopptx -g www-data -m 2750 "$DATA/store"
for d in avatars backgrounds excel; do
  install -d -o autopptx -g www-data -m 2750 "$DATA/store/$d"
done

echo "== 3/5 mã nguồn"
install -m 644 "$SRC/autopptx_server.py" "$SRC/admin.html" "$APP/"
install -m 644 "$SRC"/autopptx2/*.py "$APP/autopptx2/"
install -m 644 "$SRC/autopptx-store.service" /etc/systemd/system/autopptx-store.service
python3 -m py_compile "$APP/autopptx_server.py"
systemctl daemon-reload
systemctl enable autopptx-store >/dev/null 2>&1
systemctl restart autopptx-store
sleep 1
systemctl is-active --quiet autopptx-store || { journalctl -u autopptx-store -n 30 --no-pager; exit 1; }
curl -fsS http://127.0.0.1:8780/api/health >/dev/null

echo "== 4/5 nginx"
install -m 644 "$SRC/nginx-autopptx.conf" /etc/nginx/snippets/talent-autopptx.conf
cat > /etc/nginx/conf.d/autopptx-zones.conf <<'EOF'
# AutoPPTX: chậm việc dò mật khẩu (cả văn phòng chung 1 IP nên để rộng tay).
limit_req_zone $binary_remote_addr zone=pptxlogin:1m rate=30r/m;
EOF
if ! grep -q 'snippets/talent-\*.conf' "$SITE"; then
  cp -a "$SITE" "$SITE.bak-$(date +%Y%m%d-%H%M%S)"
  # Chèn vào server{} HTTPS, ngay trước "location / {" cuối cùng của khối đó.
  python3 - "$SITE" <<'PY'
import sys
p = sys.argv[1]
s = open(p, encoding='utf-8').read()
marker = '    location / {\n        try_files $uri $uri/ =404;\n    }\n'
assert s.count(marker) == 1, 'không tìm thấy khối location / trong ga360'
s = s.replace(marker, '    # Ứng dụng khác chạy chung tên miền (vd AutoPPTX /pptx/).\n'
                      '    include /etc/nginx/snippets/talent-*.conf;\n\n' + marker)
open(p, 'w', encoding='utf-8').write(s)
PY
fi
nginx -t
systemctl reload nginx
sleep 2   # reload chạy nền: chờ worker mới nhận cấu hình trước khi kiểm

echo "== 5/5 kiểm tra qua HTTPS"
code=$(curl -s -o /dev/null -w '%{http_code}' https://talent.goldenasia.vn/pptx/api/health)
echo "health: $code"
code=$(curl -s -o /dev/null -w '%{http_code}' https://talent.goldenasia.vn/)
echo "talent home: $code"
code=$(curl -s -o /dev/null -w '%{http_code}' https://talent.goldenasia.vn/pptx-store/excel/shared.xlsx)
echo "store trực tiếp (phải 404): $code"
echo "XONG"
