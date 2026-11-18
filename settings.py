"""Lưu / đọc cấu hình mặc định của tool trong %APPDATA%\\KHBD-Thang\\settings.json."""
import json
import os

APP_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "KHBD-Thang")
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")

# (khoá, nhãn hiển thị, nhóm)
FIELDS = [
    ("co_quan", "Cơ quan chủ quản", "Trang bìa"),
    ("truong", "Tên trường", "Trang bìa"),
    ("lop", "Lớp", "Trang bìa"),
    ("giao_vien", "Họ tên giáo viên", "Trang bìa"),
    ("nam_hoc", "Năm học", "Trang bìa"),
    ("bgh_tieu_de", "Tiêu đề BGH", "Chữ ký"),
    ("bgh_chuc_danh", "Chức danh người ký BGH", "Chữ ký"),
    ("bgh_ten", "Họ tên người ký BGH", "Chữ ký"),
    ("dia_danh", "Địa danh (ngày ký)", "Chữ ký"),
    ("tt_tieu_de", "Tiêu đề tổ trưởng", "Chữ ký"),
    ("tt_ten", "Họ tên tổ trưởng (để trống = lấy từ file tuần)", "Chữ ký"),
    ("ten_file", "Mẫu tên file xuất ({thang} = số tháng)", "Xuất file"),
    ("thu_muc_xuat", "Thư mục lưu mặc định", "Xuất file"),
]

DEFAULTS = {
    "co_quan": "UBND PHƯỜNG HẢI AN",
    "truong": "TRƯỜNG TIỂU HỌC THÀNH TÔ",
    "lop": "2",
    "giao_vien": "LÊ THỊ VIN",
    "nam_hoc": "2026 – 2027",
    "bgh_tieu_de": "BAN GIÁM HIỆU KÍ DUYỆT",
    "bgh_chuc_danh": "PHÓ HIỆU TRƯỞNG",
    "bgh_ten": "Bùi Thị Hường",
    "dia_danh": "Hải An",
    "tt_tieu_de": "TỔ TRƯỞNG KÍ DUYỆT",
    "tt_ten": "",
    "ten_file": "Lê Vin - KẾ HOẠCH BÀI DẠY THÁNG {thang} - KHỐI 2 - 2026-2027",
    "thu_muc_xuat": "",
    "xuat_pdf": True,
}

# "\\n" hiển thị trong ô nhập, sẽ được đổi thành xuống dòng khi ghép
CHUC_DANH_GOI_Y = ["PHÓ HIỆU TRƯỞNG", "HIỆU TRƯỞNG", "KT. HIỆU TRƯỞNG\\nPHÓ HIỆU TRƯỞNG"]


def load():
    data = dict(DEFAULTS)
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            data.update(json.load(f))
    except (OSError, ValueError):
        pass
    return data


def save(data):
    os.makedirs(APP_DIR, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
