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

## Đóng gói (Windows .exe & macOS .app)

### 1. Tự động build qua Git (GitHub Actions)
Dự án đã được cấu hình CI/CD tự động trong `.github/workflows/build.yml`.
- **Tự động build khi push code:** Mỗi khi push lên nhánh `master`/`main`, GitHub Actions dựng 4 bản:
  - `AutoPPTX_Studio_V2.exe` — Windows
  - `AutoPPTX_Studio_V2-macOS-Universal.zip` — **1 app dùng cho cả Mac Intel và Apple Silicon** (nên gửi bản này)
  - `AutoPPTX_Studio_V2-macOS-AppleSilicon.zip` — chỉ Mac M1/M2/M3+, nhẹ hơn
  - `AutoPPTX_Studio_V2-macOS-Intel.zip` — chỉ Mac Intel, nhẹ hơn
- **Tải bản build:** Vào tab **Actions** trên GitHub repository để tải bản Artifacts mới nhất.
- **Tự động tạo Release:** Push tag phiên bản (ví dụ: `git tag v1.0.0 && git push origin v1.0.0`) sẽ tự động tạo GitHub Release kèm file đính kèm cho cả Windows và macOS.

### 2. Build thủ công trên máy cục bộ
Dùng file `.spec` (đóng gói `LOGO`, `app_icon.ico`, `app_icon.icns`).

- **Trên Windows:**
  Chạy file `build_win.bat` hoặc lệnh:
  ```cmd
  build_win.bat
  ```
  File `.exe` sẽ được tạo tại `dist/AutoPPTX_Studio_V2.exe`.

- **Trên macOS:**
  Chạy file `build_mac.sh` hoặc lệnh:
  ```bash
  chmod +x build_mac.sh
  ./build_mac.sh
  ```
  File `.app` và `.zip` được tạo tại `dist/AutoPPTX_Studio_V2-macOS-<Intel|AppleSilicon>.zip`.

### 3. Một app cho cả Mac Intel & Apple Silicon
Bản `.app` do PyInstaller dựng chỉ chạy đúng chip của máy đã build. Bản arm64
**không** chạy trên Mac Intel — Rosetta chỉ dịch Intel → Apple Silicon, không dịch ngược.

Workflow tự gộp 2 bản thành **một** `AutoPPTX_Studio_V2-macOS-Universal.zip`:
`.app` ngoài chỉ chứa script `Contents/MacOS/launch`, script đọc `uname -m` rồi chạy
bản tương ứng trong `Contents/Resources/arm64/` hoặc `Contents/Resources/x86_64/`.
Người dùng chỉ thấy 1 app, double-click là chạy. Đổi lại dung lượng gấp đôi.

Gộp tay trên bất kỳ máy Mac nào (nếu đã có 2 file zip):
```bash
bash tools/make_mac_dual_app.sh <AppleSilicon>.zip <Intel>.zip release_out
```

Không dùng `lipo` để ghép 2 file PyInstaller **onefile**: bản ghép chỉ chạy được 1 chip,
vì PKG archive nhúng bên trong chỉ thuộc một slice.

Bản `universal2` thật (1 binary fat) hiện **không** dựng được trực tiếp: PyInstaller đòi
mọi thư viện nhị phân phải là fat binary, mà `Pillow` và `pydantic-core` không phát hành
wheel `universal2`. Muốn làm thì phải ghép wheel bằng `delocate-merge` trước mỗi lần build.

Hai cách khác cho Mac Intel:
- **Chạy trực tiếp từ source** (không cần đóng gói):
  ```bash
  python3 -m pip install -r requirements.txt
  python3 AutoPPTX_Studio_V2.py
  ```
  Cần Python 3.10+ bản x86_64 (tải ở python.org, đã kèm Tk 8.6).
- **Build ngay trên máy Mac Intel** bằng `./build_mac.sh`.

GitHub chỉ còn cấp runner Intel (`macos-15-intel`) đến 8/2027.

**Nếu job Mac lỗi:** nguyên nhân hay gặp nhất là Tcl/Tk. Python của `setup-python`
link `_tkinter` với Tcl/Tk 8.6 của Homebrew, mà runner Intel thường thiếu symlink
`tcl-tk` (Homebrew đã tách thành `tcl-tk@8`) — build vẫn xong nhưng mở app là tắt ngay.
Workflow đã có bước `Ensure Tcl/Tk 8.6` tạo symlink và chạy thử `import tkinter`
trước khi build, cùng bước `Verify macOS build` kiểm `lipo -archs` để chặn bản sai chip.
Xem log 2 bước này trước khi tìm chỗ khác.

App chưa ký số nên lần đầu mở: right-click → **Open**, hoặc chạy
`xattr -cr /Applications/AutoPPTX_Studio_V2.app`.

Lưu ý khi chạy trên macOS: kéo-thả thư mục vào cửa sổ (`windnd`) và xuất PDF kèm PPTX
(cần Microsoft PowerPoint qua COM) là tính năng chỉ có trên Windows — app vẫn chạy,
chỉ ghi log bỏ qua. Mở file PPTX rồi "Save as PDF" nếu cần.


## Ghi chú

- Cấu hình + preset lưu tại `%USERPROFILE%\.autopptx_studio_v2\`.
  Tile bản đồ OSM cache tại `...\map_cache\` (cũng đọc cache bản 1 nếu có).
- Đăng nhập + đám mây dùng cùng project Supabase với bản 1 (Excel theo tài khoản,
  avatar kho chung, ảnh nền mẫu). Cần: `pip install supabase keyring`.
