# KHBD Tháng – Ghép giáo án tuần thành kế hoạch bài dạy tháng

## Dùng phần mềm
1. Mở `KHBD_Thang.exe` (máy phải cài **Microsoft Word**).
2. **Thêm file...** → chọn các file giáo án tuần trong tháng (`.docx` hoặc `.wbk`).
   Phần mềm tự đọc số tuần và ngày, tự sắp xếp, tự đoán tháng và tên tổ trưởng.
3. Kiểm tra **Tháng** và **Tên file xuất** (sửa trực tiếp được; nút ↺ trả về tên theo mẫu).
4. Tab **Chữ ký duyệt**: tiêu đề BGH, chức danh (chọn sẵn hoặc gõ tuỳ ý; gõ `\n` để xuống dòng,
   VD `KT. HIỆU TRƯỞNG\nPHÓ HIỆU TRƯỞNG`), họ tên người ký, nơi + ngày ký, tổ trưởng.
   Tab **Trang bìa**: cơ quan, trường, lớp, giáo viên, năm học.
5. **Xem trước** → xem toàn bộ báo cáo ở khung bên phải.
6. **Xuất báo cáo tháng** → lưu `.docx` (và `.pdf` nếu bật trong Cài đặt).

Bố cục báo cáo: Trang bìa → bảng báo giảng các tuần (trang ngang) → bài dạy các tuần
(trang dọc) → bảng chữ ký BGH + Tổ trưởng ở cuối. File tuần gốc không bị thay đổi.

**Cài đặt mặc định**: lưu sẵn mọi ô (kể cả mẫu tên file, `{thang}` = số tháng, thư mục lưu)
vào `%APPDATA%\KHBD-Thang\settings.json`. Nút "Lưu thông tin đang nhập làm mặc định" lưu nhanh form hiện tại.

## Yêu cầu file tuần
- Trang đầu: bảng báo giảng có dòng `TUẦN THỨ: ...  TỪ NGÀY ... ĐẾN NGÀY ...`, sau đó là ngắt section.
- Phần bài dạy ở section sau.
- Cuối file: bảng `TỔ TRƯỞNG KÝ DUYỆT` (bị bỏ đi khi ghép, thay bằng bảng chữ ký tháng).

## Dành cho người phát triển
- `app.py` giao diện · `merger.py` ghép bằng Word COM · `docinfo.py` đọc số tuần/ngày ·
  `settings.py`, `settings_dialog.py` cài đặt.
- `template.docx` (bìa + bảng chữ ký có placeholder `{{...}}`) tạo bằng
  `python make_template.py "<file KHBD tháng mẫu.docx>"`; icon tạo bằng `python make_icon.py`.
- Build exe: chạy `build.bat` (Nuitka biên dịch sang mã máy, khó dịch ngược hơn PyInstaller).
