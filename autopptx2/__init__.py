"""AutoPPTX Studio V2 — trình dựng báo cáo PPTX kéo-thả.

Kiến trúc:
    constants   — hằng số, bố cục mặc định
    geometry    — toán bố cục DÙNG CHUNG preview & export (WYSIWYG)
    datasource  — đọc Excel (openpyxl), nhóm ảnh theo mã, khớp dữ liệu
    thumbs      — cache thumbnail nền (không chặn UI)
    exporter    — sinh PPTX thuần (không đụng Tkinter, chạy ở worker thread)
    effects     — đóng dấu ngày/giờ, mini-map GPS, tự chỉnh sáng
    qa          — đối chiếu list sales × ảnh lúc xuất (không ghi lên slide)
    editor      — canvas kéo-thả mượt (move/resize không vẽ lại toàn cục)
    app         — cửa sổ chính, sidebar, timeline, xuất file
"""

__version__ = "2.0.0"
