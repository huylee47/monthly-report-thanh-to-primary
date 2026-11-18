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
WD_WITHIN_TABLE = 12
WD_LINE_SPACE_EXACTLY = 4
WD_PAGE_NUMBER = 3   # wdActiveEndPageNumber
WD_BORDER_TOP = -1
WD_BORDER_BOTTOM = -3


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


def _bao_giang_start(doc):
    """Vị trí bắt đầu phần báo giảng trong section 1 (bỏ qua trang bìa nếu file tuần có bìa).
    = đầu bảng chứa 'TUẦN THỨ' (hoặc đầu đoạn chứa nó nếu nằm ngoài bảng)."""
    s1 = doc.Sections(1).Range
    r = doc.Range(s1.Start, s1.End)
    f = r.Find
    f.ClearFormatting()
    if not f.Execute("TUẦN THỨ", False, False, False, False, False, True, WD_FIND_STOP):
        return s1.Start
    if r.Information(WD_WITHIN_TABLE):
        return r.Tables(1).Range.Start
    return r.Paragraphs(1).Range.Start


def _para_before_first_table(doc):
    """Bảng nằm ngay đầu tài liệu -> tách trước dòng 1 để có 1 đoạn trống phía trước."""
    if not doc.Range(0, 0).Information(WD_WITHIN_TABLE):
        return
    try:
        doc.Tables(1).Split(1)
    except pywintypes.com_error:
        pass   # Word vẫn tách nhưng báo 'Object has been deleted' khi trả về bảng mới
    if doc.Range(0, 0).Information(WD_WITHIN_TABLE):
        raise MergeError("Không tách được bảng báo giảng ở đầu file tuần đầu tiên.")


def _strip_cover(doc):
    """Xoá trang bìa có sẵn trong file tuần (nội dung trước bảng báo giảng),
    chừa lại 1 đoạn trống trước bảng để sau đó chèn bìa tháng vào (không chèn lọt vào ô bảng)."""
    start = _bao_giang_start(doc)
    if start <= 0:
        _para_before_first_table(doc)
        return
    if start == 1:
        return   # chỉ có 1 dòng trống trước bảng: không có bìa
    # xoá đoạn bìa (khung/hình neo trong đoạn đó bị xoá theo)
    doc.Range(0, start - 1).Delete()   # chừa lại dấu đoạn ngay trước bảng
    _para_before_first_table(doc)
    # đoạn trống còn lại: định dạng gọn, không ngắt trang
    par = doc.Range(0, 0).Paragraphs(1).Range
    par.Style = doc.Styles(-1)  # wdStyleNormal
    _shrink_paragraph(par)


def _shrink_paragraph(par_range):
    """Đoạn trống ngăn cách giữa các bảng: thu nhỏ để không đẩy sang trang trắng."""
    pf = par_range.ParagraphFormat
    pf.SpaceBefore = 0
    pf.SpaceAfter = 0
    pf.LineSpacingRule = WD_LINE_SPACE_EXACTLY
    pf.LineSpacing = 1
    pf.PageBreakBefore = False
    par_range.Font.Size = 1


def _cell_trim_empty_paragraphs(doc, cell):
    """Xoá các đoạn trống thừa trong ô (sinh ra khi gộp ô)."""
    pars = cell.Range.Paragraphs
    for i in range(pars.Count, 1, -1):
        p = pars(i).Range
        if p.Text.strip(chr(13) + chr(7) + " ") == "":
            doc.Range(p.Start - 1, p.Start).Delete()


def _split_merged_at_page_breaks(doc, table):
    """Bảng báo giảng: ô gộp dọc (cột 'THỨ', 'BUỔI') vắt qua 2 trang thì Word không kẻ đường
    chân ở cuối trang. Tách ô gộp đúng tại chỗ sang trang thành 2 ô gộp riêng -> đủ đường kẻ."""
    cells = table.Range.Cells
    have = {}        # cột -> các dòng có ô bắt đầu
    row_page = {}
    last_row = 0
    for i in range(1, cells.Count + 1):
        c = cells(i)
        r, col = c.RowIndex, c.ColumnIndex
        last_row = max(last_row, r)
        if col in (1, 2):
            have.setdefault(col, []).append(r)
        elif col == 3 and r not in row_page:
            row_page[r] = c.Range.Information(WD_PAGE_NUMBER)
    jobs = []
    for col, rows in have.items():
        rows = sorted(rows) + [last_row + 1]
        for r0, nxt in zip(rows, rows[1:]):
            r1 = nxt - 1
            pages = [row_page.get(r) for r in range(r0, r1 + 1)]
            if r1 <= r0 or None in pages or pages[0] == pages[-1]:
                continue
            cuts = [r for r in range(r0, r1) if pages[r - r0] != pages[r - r0 + 1]]
            jobs.append((r0, col, r1, cuts))
    # làm từ dưới lên để chỉ số dòng/ô phía trên không bị ảnh hưởng
    for r0, col, r1, cuts in sorted(jobs, reverse=True):
        try:
            table.Cell(r0, col).Split(r1 - r0 + 1, 1)
            bounds = [r0] + [c + 1 for c in cuts] + [r1 + 1]
            for s, e in zip(bounds, bounds[1:]):
                # ô tách ra thừa hưởng viền của ô gốc (thường không có viền dưới) -> kẻ lại
                # trên từng ô TRƯỚC khi gộp (Word lấy viền dưới của ô gộp từ ô cuối nhóm)
                rb = table.Cell(e - 1, 3).Borders(WD_BORDER_BOTTOM)
                style = rb.LineStyle or 1
                width = rb.LineWidth if rb.LineStyle else 4
                for r in range(s, e):
                    for side in (WD_BORDER_TOP, WD_BORDER_BOTTOM):
                        b = table.Cell(r, col).Borders(side)
                        b.LineStyle = style
                        b.LineWidth = width
                if e - 1 > s:
                    table.Cell(s, col).Merge(table.Cell(e - 1, col))
                _cell_trim_empty_paragraphs(doc, table.Cell(s, col))
        except pywintypes.com_error:
            pass   # bảng lạ không tách được -> giữ nguyên như bản gốc


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
        # file tuần có sẵn trang bìa (thường là tuần 1) -> bỏ, bìa tháng sẽ được chèn sau
        _strip_cover(base)

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
                bao_giang = src.Range(_bao_giang_start(src), s1.End - 1)
                bai_day = src.Range(src.Sections(2).Range.Start, sig.Range.Start)

                # 1) báo giảng: thêm 1 đoạn ngăn cách (thu nhỏ) cuối section ngang đầu tiên,
                #    rồi nối báo giảng tuần này; "sang trang trước" thay cho ngắt trang
                #    để không sinh trang trắng khi bảng trước vừa kín trang.
                pos = base.Sections(1).Range.End - 1
                base.Range(pos, pos).InsertBefore("\r")
                _shrink_paragraph(base.Range(pos, pos + 1))
                pos += 1
                base.Range(pos, pos).FormattedText = bao_giang.FormattedText
                base.Range(pos, pos).Paragraphs(1).PageBreakBefore = True

                # 2) bài dạy: đoạn mới ở cuối tài liệu, đặt "sang trang trước"
                _end_point(base).InsertParagraphAfter()
                pos = _end_point(base).Start
                _end_point(base).FormattedText = bai_day.FormattedText
                base.Range(pos, pos).Paragraphs(1).PageBreakBefore = True
            finally:
                src.Close(False)

        s1 = base.Sections(1).Range
        _shrink_paragraph(base.Range(s1.End - 1, s1.End))   # đoạn chứa ngắt section

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
        base.Close(False)

        # Mở lại file vừa lưu (bố cục trang giống hệt khi người dùng mở bằng Word)
        # rồi mới tính chỗ sang trang để tách ô gộp của bảng báo giảng.
        step("Căn chỉnh bảng báo giảng...")
        # tắt ScreenUpdating làm Word dàn trang/xuất PDF sai (mất viền tiêu đề...) -> bật lại
        word.ScreenUpdating = True
        doc = word.Documents.Open(tmp_out, AddToRecentFiles=False, Visible=False)
        try:
            doc.Repaginate()
            tables = doc.Sections(1).Range.Tables   # section 1 = bìa + báo giảng (khổ ngang)
            for k in range(tables.Count, 0, -1):
                _split_merged_at_page_breaks(doc, tables(k))
            doc.Save()
        finally:
            doc.Close(False)
        if out_pdf:
            # Xuất PDF từ file Word VỪA LƯU (mở lại), không xuất từ bản đang ghép trong bộ nhớ:
            # bản trong bộ nhớ có bố cục cũ nên PDF thiếu đường kẻ tiêu đề, lệch hình...
            # Mở lại giống như người dùng mở file .docx -> PDF khớp với Word.
            step("Xuất PDF...")
            saved = word.Documents.Open(tmp_out, ReadOnly=True, AddToRecentFiles=False, Visible=False)
            try:
                saved.Repaginate()
                saved.ExportAsFixedFormat(out_pdf, WD_EXPORT_PDF)
            finally:
                saved.Close(False)
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
