# AutoPPTX Studio V2

Trình dựng báo cáo PowerPoint kéo-thả — bản viết lại của `AutoPPTX_Studio.py`
với trọng tâm: **kéo-thả mượt**, kiến trúc module hoá, xuất file WYSIWYG.

## Cài đặt & chạy

```bash
pip install -r requirements.txt
python AutoPPTX_Studio_V2.py
```

Yêu cầu Python 3.10+. Không cần pandas — Excel đọc trực tiếp bằng openpyxl
(app mở gần như tức thì).

## Quy trình sử dụng

1. **Đăng nhập** tài khoản công ty (cùng hệ thống RP SALES / bản AutoPPTX cũ).
   Mật khẩu nhớ được trên máy (Windows Credential Manager).
2. Tab **Nguồn**: đồng bộ Excel + Avatar trên đám mây, thêm **nhiều thư mục ảnh**
   (chọn thư mục cha → tick nhiều thư mục con, hoặc kéo-thả nhiều folder).
3. Tab **Slide**: bấm phần tử trên canvas — tiêu đề hiện font/cỡ/màu; ảnh hiện
   bo góc, đóng dấu ngày giờ, mini-map GPS. Double-click cũng mở inspector này
   (không còn cửa sổ phụ).
4. Tab **Xuất** → Xuất PPTX.

## Điều khiển trong khung xem trước

| Thao tác | Phím / chuột |
|---|---|
| Di chuyển phần tử | kéo chuột trái |
| Đổi kích thước | kéo 1 trong 8 tay nắm (Shift = giữ tỉ lệ) |
| Tắt hít (snap) tạm thời | giữ Alt khi kéo |
| Zoom tới con trỏ | Ctrl + lăn chuột (25–400%) |
| Pan | kéo vùng trống hoặc chuột giữa |
| Nhích 0.02″ / 0.2″ | phím mũi tên / Shift + mũi tên |
| Undo / Redo | Ctrl+Z / Ctrl+Y |
| Nhóm trước / sau | PageUp / PageDown |
| Mở inspector (font / bo góc / đóng dấu) | double-click phần tử |

Ngoài ra: khoá/ẩn từng phần tử, lưới 0.5″, căn nhanh theo slide
(trái/giữa/phải/trên/giữa/dưới), preset bố cục đặt tên tuỳ ý,
tự lưu cấu hình (mở lại là khôi phục).

## Vì sao bản này mượt hơn bản cũ

| Bản cũ | Bản V2 |
|---|---|
| Kéo cạnh gọi vẽ lại **toàn bộ** canvas mỗi sự kiện chuột (kèm ghi file PNG tạm ra đĩa) | Chỉ vẽ lại đúng phần tử đang kéo, giới hạn ~30fps |
| Di chuyển kéo theo rebuild media pool + timeline (widget CTk) | Move chỉ là `canvas.move()` — sidebar/timeline đứng yên |
| Ảnh decode trên luồng UI khi thiếu cache | Mọi decode ở thread nền; UI chỉ resize thumbnail RAM |
| 1 class ~10.700 dòng trộn UI + ảnh + mạng + PPTX | 7 module tách bạch, exporter thuần không đụng Tkinter |

## Cấu trúc mã

```
AutoPPTX_Studio_V2.py   # điểm chạy
autopptx2/
  constants.py    # hằng số, bố cục mặc định
  geometry.py     # toán bố cục DÙNG CHUNG preview & export (WYSIWYG)
  datasource.py   # Excel, nhóm ảnh, khớp mã (kể cả gộp block Building Premium)
  thumbs.py       # cache thumbnail chạy nền
  fonts.py        # quét Windows Fonts, mẫu Việt, resolve .ttf
  effects.py      # đóng dấu ngày/giờ, mini-map GPS, tự chỉnh sáng
  style.py        # vẽ chữ preview từ đúng file .ttf
  exporter.py     # sinh PPTX (thread-safe, progress, huỷ, chia part)
  editor.py       # canvas kéo-thả (snap, guides, zoom, undo)
  app.py          # cửa sổ chính, sidebar 3 tab, timeline
```

## Đóng gói exe

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name AutoPPTX_Studio_V2 AutoPPTX_Studio_V2.py
```

## Ghi chú

- Cấu hình + preset lưu tại `%USERPROFILE%\.autopptx_studio_v2\`.
  Tile bản đồ OSM cache tại `...\map_cache\` (cũng đọc cache bản 1 nếu có).
- Đăng nhập + đám mây dùng cùng project Supabase với bản 1 (Excel theo tài khoản,
  avatar kho chung, ảnh nền mẫu). Cần: `pip install supabase keyring`.
