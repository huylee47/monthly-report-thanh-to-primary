"""Ghép các file KHBD tuần thành KHBD tháng bằng Microsoft Word (COM)."""
import os
import shutil
import sys
import tempfile

import pythoncom
import pywintypes
import win32com.client as w32

WD_PAGE_BREAK = 7
WD_COLLAPSE_END = 0
WD_FORMAT_DOCX = 16
WD_EXPORT_PDF = 17
WD_REPLACE_ALL = 2
WD_FIND_STOP = 0
WD_CHARACTER = 1
WD_BACKWARD = -1073741823


class MergeError(Exception):
    pass


def resource_path(name):
    """Đường dẫn tới file đi kèm (chạy từ mã nguồn hoặc từ exe Nuitka/PyInstaller)."""
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, name)


def _sig_table(doc):
    """Bảng chữ ký cuối file tuần (bảng cuối cùng có chữ 'DUYỆT')."""
    for i in range(doc.Tables.Count, 0, -1):
        t = doc.Tables(i)
        if "DUYỆT" in t.Range.Text.upper():
            return t
    return None


def _check_week_doc(doc, name):
    if doc.Sections.Count < 2:
        raise MergeError(
            f"File '{name}' không đúng cấu trúc: cần phần báo giảng (trang ngang) "
            "và phần bài dạy (trang dọc) tách nhau bằng ngắt section.")
    if "TUẦN THỨ" not in doc.Sections(1).Range.Text.upper():
        raise MergeError(f"File '{name}': không thấy bảng báo giảng 'TUẦN THỨ' ở trang đầu.")
    sig = _sig_table(doc)
    if sig is None:
        raise MergeError(f"File '{name}': không thấy bảng 'TỔ TRƯỞNG KÝ DUYỆT' ở cuối.")
    return sig


def _escape_replace(text):
    text = (text or "").replace("^", "^^")
    return text.replace("\r\n", "\n").replace("\n", "^l")


def _replace_all(doc, key, value):
    rng = doc.Content
    f = rng.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    # Word COM bỏ qua tham số đặt tên khi gọi muộn -> phải truyền theo vị trí:
    # FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
    # MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace
    f.Execute(key, True, False, False, False, False, True, WD_FIND_STOP, False,
              _escape_replace(value), WD_REPLACE_ALL)


def _end_point(doc):
    r = doc.Content
    r.Collapse(WD_COLLAPSE_END)
    return r


def merge(week_paths, values, out_docx, out_pdf=None, progress=None):
    """week_paths: danh sách file tuần đã sắp xếp.
    values: dict placeholder -> chuỗi (CO_QUAN, TRUONG, LOP, GIAO_VIEN, NAM_HOC,
            BGH_TIEU_DE, BGH_CHUC_DANH, BGH_TEN, NGAY_KY, TT_TIEU_DE, TT_TEN).
    """
    def step(msg):
        if progress:
            progress(msg)

    if not week_paths:
        raise MergeError("Chưa chọn file tuần nào.")
    template = resource_path("template.docx")
    if not os.path.exists(template):
        raise MergeError("Thiếu file mẫu template.docx đi kèm phần mềm.")

    work = tempfile.mkdtemp(prefix="khbd_")
    copies = []
    for i, p in enumerate(week_paths):
        dst = os.path.join(work, f"tuan_{i}.docx")
        try:
            shutil.copyfile(p, dst)
        except OSError as e:
            raise MergeError(f"Không đọc được file '{os.path.basename(p)}': {e}")
        copies.append(dst)
    tpl_copy = os.path.join(work, "template.docx")
    shutil.copyfile(template, tpl_copy)
    tmp_out = os.path.join(work, "ket_qua.docx")

    pythoncom.CoInitialize()
    word = None
    try:
        step("Đang khởi động Microsoft Word...")
        try:
            word = w32.DispatchEx("Word.Application")
        except pywintypes.com_error:
            raise MergeError("Không khởi động được Microsoft Word. Máy cần cài Word để dùng phần mềm này.")
        word.Visible = False
        word.DisplayAlerts = 0
        word.ScreenUpdating = False

        names = [os.path.basename(p) for p in week_paths]
        step(f"Mở {names[0]} làm bản gốc...")
        base = word.Documents.Open(copies[0], AddToRecentFiles=False, Visible=False)
        _check_week_doc(base, names[0]).Delete()

        for i in range(1, len(copies)):
            step(f"Ghép {names[i]} ({i + 1}/{len(copies)})...")
            src = word.Documents.Open(copies[i], ReadOnly=True, AddToRecentFiles=False, Visible=False)
            try:
                sig = _check_week_doc(src, names[i])
                # Đóng băng số thứ tự tự động (1., 2., a)...) thành chữ như đang hiển thị,
                # tránh danh sách bị đánh số nối tiếp từ tuần trước khi ghép.
                if src.ListParagraphs.Count:
                    src.ConvertNumbersToText()
                s1 = src.Sections(1).Range
                bao_giang = src.Range(s1.Start, s1.End - 1)
                bai_day = src.Range(src.Sections(2).Range.Start, sig.Range.Start)

                # 1) báo giảng: nối vào cuối section ngang đầu tiên, sang trang mới
                pos = base.Sections(1).Range.End - 1
                base.Range(pos, pos).InsertBreak(WD_PAGE_BREAK)
                pos = base.Sections(1).Range.End - 1
                base.Range(pos, pos).FormattedText = bao_giang.FormattedText

                # 2) bài dạy: nối vào cuối tài liệu, sang trang mới
                _end_point(base).InsertBreak(WD_PAGE_BREAK)
                _end_point(base).FormattedText = bai_day.FormattedText
            finally:
                src.Close(False)

        step("Thêm trang bìa và phần chữ ký...")
        tpl = word.Documents.Open(tpl_copy, ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            tsig = tpl.Tables(tpl.Tables.Count)
            cover = tpl.Range(0, tsig.Range.Start)
            # bỏ các dòng trống thừa cuối bìa (tránh đẩy sang trang trắng), giữ 1 dấu đoạn
            cover.MoveEndWhile(chr(13), WD_BACKWARD)
            cover.MoveEnd(WD_CHARACTER, 1)
            cover_len = cover.End - cover.Start
            base.Range(0, 0).FormattedText = cover.FormattedText
            base.Range(cover_len, cover_len).InsertBreak(WD_PAGE_BREAK)

            end = _end_point(base)
            end.InsertParagraphAfter()
            _end_point(base).FormattedText = tsig.Range.FormattedText
        finally:
            tpl.Close(False)

        step("Điền thông tin...")
        for key, val in values.items():
            _replace_all(base, "{{" + key + "}}", val)

        step("Lưu file Word...")
        base.SaveAs2(tmp_out, FileFormat=WD_FORMAT_DOCX, AddToRecentFiles=False)
        if out_pdf:
            step("Xuất PDF...")
            base.ExportAsFixedFormat(out_pdf, WD_EXPORT_PDF)
        base.Close(False)
    except pywintypes.com_error as e:
        detail = ""
        try:
            detail = e.excepinfo[2] or ""
        except (AttributeError, IndexError, TypeError):
            pass
        raise MergeError(f"Word báo lỗi khi ghép: {detail or e}")
    finally:
        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except pywintypes.com_error:
                pass
        pythoncom.CoUninitialize()

    try:
        shutil.copyfile(tmp_out, out_docx)
    except OSError as e:
        raise MergeError(f"Không lưu được '{out_docx}' (file có đang mở trong Word?): {e}")
    finally:
        shutil.rmtree(work, ignore_errors=True)
    step("Hoàn tất.")
    return out_docx
