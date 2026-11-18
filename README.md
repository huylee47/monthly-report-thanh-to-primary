# KHBD Tháng – Ghép giáo án tuần thành Kế hoạch bài dạy tháng

Phần mềm Windows giúp giáo viên ghép các file **Kế hoạch bài dạy (KHBD) tuần** thành
**KHBD tháng** hoàn chỉnh: có trang bìa, báo giảng các tuần, bài dạy các tuần và phần
chữ ký duyệt của Ban giám hiệu + Tổ trưởng. Có xem trước toàn bộ trước khi xuất.

![Bố cục](https://img.shields.io/badge/Windows-10%2F11-blue) ![Word](https://img.shields.io/badge/c%E1%BA%A7n-Microsoft%20Word-2b579a)

## Tính năng

- Chọn nhiều file tuần (`.docx`, `.wbk`); tự đọc **số tuần, ngày bắt đầu/kết thúc**, tự sắp xếp,
  tự đoán **tháng** và **tên tổ trưởng**.
- Cảnh báo khi trùng tuần, tuần không liên tiếp, tuần không thuộc tháng, file sai mẫu.
- Điền thông tin **trang bìa** (cơ quan, trường, lớp, giáo viên, năm học).
- Điền **phần chữ ký**: tiêu đề BGH, chức danh tuỳ chỉnh (Phó Hiệu trưởng, Hiệu trưởng,
  `KT. HIỆU TRƯỞNG\nPHÓ HIỆU TRƯỞNG`...), họ tên người ký, nơi + ngày ký, tổ trưởng.
- **Xem trước** toàn bộ báo cáo ngay trong phần mềm (phóng to/thu nhỏ, cuộn trang).
- **Xuất** `.docx` và (tuỳ chọn) `.pdf`; tên file sửa trực tiếp hoặc sinh theo mẫu.
- **Cài đặt mặc định** lưu mọi giá trị hay dùng, lần sau mở là có sẵn.
- Giữ nguyên định dạng, bảng, hình ảnh, header/footer của file gốc; **không sửa file tuần gốc**.

## Bố cục báo cáo tháng

```
Trang bìa (ngang)
Báo giảng tuần 1 → tuần N   (trang ngang)
Bài dạy tuần 1 → tuần N     (trang dọc, mỗi tuần sang trang mới)
BAN GIÁM HIỆU KÍ DUYỆT            Hải An, ngày … tháng … năm …
PHÓ HIỆU TRƯỞNG                   TỔ TRƯỞNG KÍ DUYỆT
<Họ tên>                          <Họ tên>
```

## Sử dụng

1. Mở `KHBD_Thang.exe` (không cần cài Python; **máy phải có Microsoft Word**).
2. **Thêm file...** → chọn các file giáo án tuần của tháng.
3. Kiểm tra **Tháng** và **Tên file xuất** (nút ↺ trả về tên theo mẫu).
4. Điền tab **Chữ ký duyệt** và **Trang bìa**.
5. **Xem trước** → kiểm tra → **Xuất báo cáo tháng**.

Cấu hình mặc định lưu tại `%APPDATA%\KHBD-Thang\settings.json` (mở bằng nút **Cài đặt mặc định**;
trong mẫu tên file, `{thang}` được thay bằng số tháng).

### Yêu cầu với file tuần

| Phần | Nhận diện |
|---|---|
| Báo giảng | Section đầu tiên (trang ngang), có bảng `TUẦN THỨ: … TỪ NGÀY … ĐẾN NGÀY …` |
| Bài dạy | Các section sau |
| Chữ ký tuần | Bảng cuối cùng có chữ `KÝ DUYỆT` / `KÍ DUYỆT` – được bỏ khi ghép |

## Cách hoạt động

- `docinfo.py` đọc nhanh XML của file Word (không cần mở Word) để lấy số tuần, ngày, tên tổ trưởng.
- `merger.py` điều khiển Microsoft Word qua COM (pywin32):
  file tuần đầu làm bản gốc → nối báo giảng các tuần sau vào section ngang → nối bài dạy vào cuối →
  chèn trang bìa và bảng chữ ký từ `template.docx` → thay các placeholder `{{...}}` → lưu `.docx`/xuất `.pdf`.
  Số thứ tự tự động của các tuần sau được "đóng băng" thành chữ để không bị đánh số nối tiếp.
  File tuần có sẵn trang bìa (VD tuần 1) được tự bỏ bìa. Mỗi tuần sang trang bằng "page break before"
  (không sinh trang trắng); ô gộp "THỨ/BUỔI" vắt qua 2 trang được tách tại chỗ sang trang để có đường kẻ chân.
- `app.py` giao diện tkinter; xem trước PDF bằng PyMuPDF (chỉ vẽ các trang đang hiển thị).

## Cấu trúc mã nguồn

| File | Vai trò |
|---|---|
| `app.py` | Giao diện chính, xem trước, xuất file |
| `merger.py` | Lõi ghép bằng Word COM |
| `docinfo.py` | Đọc thông tin tuần từ file `.docx`/`.wbk` |
| `settings.py`, `settings_dialog.py` | Cấu hình mặc định và cửa sổ Cài đặt |
| `template.docx` | Mẫu trang bìa + bảng chữ ký (placeholder `{{CO_QUAN}}`, `{{BGH_TEN}}`, …) |
| `make_template.py` | Tạo lại `template.docx` từ một file KHBD tháng mẫu |
| `make_icon.py`, `icon.ico` | Icon phần mềm |
| `build.bat` | Build file exe bằng Nuitka |

## Phát triển

```bat
python -m pip install -r requirements.txt
python app.py
```

Tạo lại mẫu (khi đổi trang bìa):

```bat
python make_template.py "đường\dẫn\KẾ HOẠCH THÁNG mẫu.docx"
```

Placeholder trong mẫu: `CO_QUAN, TRUONG, LOP, GIAO_VIEN, NAM_HOC, BGH_TIEU_DE, BGH_CHUC_DANH,
BGH_TEN, NGAY_KY, TT_TIEU_DE, TT_TEN`.

## Build file exe

```bat
build.bat
```

Nuitka biên dịch mã Python sang C/mã máy và đóng gói thành một file `build\KHBD_Thang.exe`
(kèm Python runtime, thư viện, `template.docx`, icon). Cách này khó dịch ngược hơn nhiều so với
PyInstaller (không còn bytecode Python), dù không có biện pháp nào tuyệt đối. Lần build đầu
mất 10–30 phút do Nuitka tải trình biên dịch C.

## Thiết kế

Lê Khánh Huy – Zalo: 0938194069

## Yêu cầu

- Windows 10/11 64-bit
- Microsoft Word (2013 trở lên)
- Chỉ khi phát triển: Python 3.12+ , `pywin32`, `pymupdf`, `nuitka`
