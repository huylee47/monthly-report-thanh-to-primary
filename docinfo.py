"""Đọc nhanh thông tin của file KHBD tuần (không cần mở Word)."""
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import date
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

RE_TUAN = re.compile(r"TUẦN\s*THỨ\s*:?\s*(\d+)", re.I)
RE_NGAY = re.compile(
    r"TỪ\s*NGÀY\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4}).*?"
    r"ĐẾN\s*NGÀY\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", re.I | re.S)


@dataclass
class WeekInfo:
    path: str
    week: int | None = None
    start: date | None = None
    end: date | None = None
    to_truong: str = ""
    errors: list = field(default_factory=list)

    @property
    def label(self):
        w = f"Tuần {self.week}" if self.week is not None else "Tuần ?"
        if self.start and self.end:
            w += f"  ({self.start:%d/%m} – {self.end:%d/%m/%Y})"
        return w


def _text(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def _paras(el):
    return [_text(p).strip() for p in el.iter(W + "p")]


def _safe_date(y, m, d):
    try:
        return date(int(y), int(m), int(d))
    except ValueError:
        return None


def read_week(path):
    info = WeekInfo(path)
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("word/document.xml"))
    except (zipfile.BadZipFile, KeyError, OSError, ET.ParseError):
        info.errors.append("Không đọc được file (không phải file Word .docx/.wbk hợp lệ).")
        return info
    body = root.find(W + "body")
    tables = [el for el in body if el.tag == W + "tbl"]
    bg = next((t for t in tables if RE_TUAN.search(_text(t))), None)
    if bg is None:
        info.errors.append("Không tìm thấy bảng báo giảng (dòng 'TUẦN THỨ: ...').")
    else:
        txt = _text(bg)
        info.week = int(RE_TUAN.search(txt).group(1))
        m = RE_NGAY.search(txt)
        if m:
            d1, m1, y1, d2, m2, y2 = m.groups()
            info.start = _safe_date(y1, m1, d1)
            info.end = _safe_date(y2, m2, d2)
    sig = next((t for t in reversed(tables) if "DUYỆT" in _text(t).upper()), None)
    if sig is None:
        info.errors.append("Không tìm thấy bảng chữ ký 'TỔ TRƯỞNG KÝ DUYỆT' ở cuối file.")
    else:
        names = [p for p in _paras(sig) if p and "DUYỆT" not in p.upper()]
        if names:
            info.to_truong = re.sub(r"\s+", " ", names[-1])
    return info


def guess_month(infos):
    """Tháng xuất hiện nhiều nhất trong các ngày (thứ 2 → thứ 6) của các tuần."""
    c = Counter()
    for i in infos:
        if i.start:
            c[(i.start.year, i.start.month)] += 1
        if i.end:
            c[(i.end.year, i.end.month)] += 1
    if not c:
        return None, None
    (y, m), _ = c.most_common(1)[0]
    return m, y
