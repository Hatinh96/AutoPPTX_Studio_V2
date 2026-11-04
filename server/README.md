# Máy chủ kho tài nguyên AutoPPTX

Thay Supabase. Lưu FILE TỔNG Excel, kho avatar (ảnh theo mã RP), ảnh nền và tài khoản
người dùng trên VPS của công ty.

| | |
|---|---|
| VPS | GAM-DC-S04-WEB · `103.153.74.177` (VPS 360, chung với talent.goldenasia.vn) |
| Địa chỉ | `https://talent.goldenasia.vn/pptx/` (trang quản trị) · API `…/pptx/api/…` |
| Dịch vụ | `autopptx-store` (systemd) chạy `127.0.0.1:8780`, nginx chuyển `/pptx/` vào |
| Mã | `/opt/autopptx-server/` |
| Dữ liệu | `/var/lib/autopptx/` — `autopptx.db` (tài khoản, phiên, danh mục file) + `store/{avatars,backgrounds,excel}` |
| Log | `journalctl -u autopptx-store -f` |

## Cách hoạt động

- **Đăng nhập**: email + mật khẩu (băm PBKDF2 trên máy chủ). App nhận một *phiên* (token)
  và chỉ nhớ phiên đó (Windows Credential Manager / macOS Keychain), không nhớ mật khẩu.
  Phiên tự gia hạn khi dùng, hết hạn sau 60 ngày không dùng. Khoá tài khoản / cấp mật
  khẩu tạm / đổi mật khẩu thì các phiên cũ bị huỷ.
- **Vai trò**: `user` (tải FILE TỔNG, avatar, nền; đẩy nền riêng và FILE TỔNG riêng),
  `admin` (thêm đẩy FILE TỔNG cho tất cả, avatar, nền chung, quản lý tài khoản),
  `super_admin` (như admin, không ai khác sửa/xoá được).
- **FILE TỔNG**: một bản *chung* cho mọi người. Tài khoản có thể có bản *riêng* (ưu tiên hơn).
  «Đẩy cho TẤT CẢ» thay bản chung và bỏ hết bản riêng.
- **Avatar**: giữ nguyên cây thư mục theo kênh. Tên file = mã RP, bỏ đuôi `A` / `AA` / `_Ava`
  (`1994HAIRA.jpg` → `1994HAIR`), cùng quy tắc app dùng khi xuất. Ảnh quá lớn được thu nhỏ
  về cạnh dài 1600px. Đẩy lại cả thư mục: ảnh không đổi được bỏ qua (so sha ảnh gốc).
- **Đồng bộ trên app**: lần đầu tải cả kho (~1,8 GB, 4 luồng); lần sau chỉ tải ảnh mới/đổi.
  Ảnh đã có sẵn trên máy (kho Supabase cũ) được nhận luôn, không tải lại.
- **Đối chiếu mã RP** (tab trên trang quản trị): dùng đúng `ExcelSource` + `AvatarIndex` của
  app → mã thiếu ảnh, ảnh không khớp mã nào, mã có nhiều ảnh; tải CSV mã thiếu ảnh.

## Việc thường làm

Cập nhật mã máy chủ (từ thư mục BD, Git Bash):

```bash
bash server/deploy.sh
```

Quên mật khẩu Super Admin (in mật khẩu tạm mới):

```bash
ssh -i ~/.ssh/ga360_deploy root@103.153.74.177 "sudo -u autopptx env AP_DATA=/var/lib/autopptx python3 /opt/autopptx-server/autopptx_server.py set-password EMAIL"
```

Đưa hàng loạt ảnh lên mà không qua trang web (nhanh hơn nhiều khi vài GB):

```bash
python server/bulk_prepare.py avatars "Data avatar" OUT
tar -C OUT -cf - avatars | ssh -i ~/.ssh/ga360_deploy root@103.153.74.177 "mkdir -p /tmp/ap_import && tar -C /tmp/ap_import -xf -"
scp -i ~/.ssh/ga360_deploy OUT/src_map_avatars.json root@103.153.74.177:/tmp/src_map.json
ssh -i ~/.ssh/ga360_deploy root@103.153.74.177 "cp -a /tmp/ap_import/avatars/. /var/lib/autopptx/store/avatars/ && chown -R autopptx:www-data /var/lib/autopptx/store && chown autopptx /tmp/src_map.json && sudo -u autopptx env AP_DATA=/var/lib/autopptx python3 /opt/autopptx-server/autopptx_server.py reindex /tmp/src_map.json && rm -rf /tmp/ap_import /tmp/src_map.json"
```

Sao lưu (DB + kho):

```bash
ssh -i ~/.ssh/ga360_deploy root@103.153.74.177 "sqlite3 /var/lib/autopptx/autopptx.db '.backup /root/autopptx-db-$(date +%F).db' && tar -C /var/lib/autopptx -czf /root/autopptx-store-$(date +%F).tgz store"
```

## Chạy thử trên máy

```bash
AP_DATA=server/data python server/autopptx_server.py create-admin admin@local.test
AP_DATA=server/data python server/autopptx_server.py serve
AUTOPPTX_SERVER=http://127.0.0.1:8780 python AutoPPTX_Studio_V2.py
```

Kiểm tra tự động: `python -m unittest tests.test_store_server` (chạy máy chủ thật trên cổng tạm
và dùng `CloudClient` của app nối vào).

## Đổi sang tên miền riêng

1. Trỏ DNS (vd `pptx.goldenasia.vn`) về `103.153.74.177`.
2. Thêm server{} nginx cho tên miền đó (copy các `location` trong `nginx-autopptx.conf`),
   chạy `certbot --nginx -d pptx.goldenasia.vn`.
3. Đổi `SERVER_URL` trong `autopptx2/constants.py`, build lại app. Bản cũ vẫn chạy qua
   `talent.goldenasia.vn/pptx` cho tới khi gỡ.

## Lưu ý

- `include /etc/nginx/snippets/talent-*.conf;` đã thêm vào `/etc/nginx/sites-available/ga360`
  và vào mẫu `E:\GOLDEN AI\360 VIEW\deploy\nginx-ga360.conf`, để deploy lại trang talent không
  làm mất đường `/pptx/`.
- Dịch vụ bị giới hạn 1,5 CPU / 1,5 GB RAM để không ảnh hưởng trang talent.
