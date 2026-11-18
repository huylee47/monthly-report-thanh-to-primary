"""Chỉ dùng khi phát triển: tạo template.docx (trang bìa + bảng chữ ký) từ KHBD tháng 10.

Chạy: python make_template.py "<đường dẫn file tháng mẫu .docx>"
"""
import os
import shutil
import sys
import tempfile

import win32com.client as w32

HERE = os.path.dirname(os.path.abspath(__file__))

# Thay đúng chuỗi (không ghi đè cả đoạn để giữ "neo" của khung trang trí bìa)
REPLACE = [
    ("UBND PHƯỜNG HẢI AN", "{{CO_QUAN}}"),
    ("TRƯỜNG TIỂU HỌC THÀNH TÔ", "{{TRUONG}}"),
    ("LỚP 2", "LỚP {{LOP}}"),
    ("LÊ THỊ VIN", "{{GIAO_VIEN}}"),
    ("2026 – 2027", "{{NAM_HOC}}"),
]


def replace_text(doc, find, repl, match_case=True):
    rng = doc.Content
    f = rng.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    # FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
    # MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace(1 = một lần)
    return f.Execute(find, match_case, False, False, False, False, True, 0, False, repl, 1)


def fix_fonts(doc, old=".VnTime", new="Times New Roman"):
    rng = doc.Content
    f = rng.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    f.Font.Name = old
    f.Replacement.Font.Name = new
    f.Execute("", False, False, False, False, False, True, 0, True, "", 2)


# Bảng chữ ký tháng: (ô trái, ô phải, in đậm)
SIG_ROWS = [
    ("", "{{NGAY_KY}}", False),
    ("{{BGH_TIEU_DE}}", "{{TT_TIEU_DE}}", True),
    ("{{BGH_CHUC_DANH}}", "", True),
    ("", "", False),
    ("", "", False),
    ("", "", False),
    ("{{BGH_TEN}}", "{{TT_TEN}}", True),
]


def build_sig_table(doc):
    end = doc.Content
    end.Collapse(0)
    end.InsertParagraphAfter()
    end = doc.Content
    end.Collapse(0)
    t = doc.Tables.Add(end, len(SIG_ROWS), 2)
    t.Borders.Enable = False
    t.PreferredWidthType = 2  # phần trăm
    t.PreferredWidth = 100
    for r, (left, right, bold) in enumerate(SIG_ROWS, start=1):
        for c, txt in ((1, left), (2, right)):
            cell = t.Cell(r, c).Range
            cell.Text = txt
            cell.ParagraphFormat.Alignment = 1
            cell.ParagraphFormat.SpaceBefore = 0
            cell.ParagraphFormat.SpaceAfter = 0
            cell.ParagraphFormat.FirstLineIndent = 0
            cell.ParagraphFormat.LeftIndent = 0
            cell.Font.Bold = bold
            cell.Font.Italic = False
            cell.Font.Name = "Times New Roman"
            cell.Font.Size = 14
    t.Cell(1, 2).Range.Font.Italic = True
    t.Rows.AllowBreakAcrossPages = False
    # giữ cả bảng trên cùng một trang
    for r in range(1, len(SIG_ROWS)):
        t.Rows(r).Range.ParagraphFormat.KeepWithNext = True


def main(src):
    tmp = os.path.join(tempfile.mkdtemp(), "src.docx")
    shutil.copy(src, tmp)
    out = os.path.join(HERE, "template.docx")
    wd = w32.DispatchEx("Word.Application")
    wd.Visible = False
    wd.DisplayAlerts = 0
    try:
        d = wd.Documents.Open(tmp, AddToRecentFiles=False)
        first_bg = next(d.Tables(i) for i in range(1, d.Tables.Count + 1)
                        if "TUẦN THỨ" in d.Tables(i).Range.Text)
        print("mở xong", flush=True)
        # Giữ nguyên tài liệu gốc (styles, khung trang trí bìa), chỉ xoá từ bảng báo giảng đầu tiên trở đi
        d.Range(first_bg.Range.Start, d.Content.End).Delete()
        print("đã xoá phần sau bìa", flush=True)
        build_sig_table(d)
        print("đã tạo bảng chữ ký", flush=True)
        for key, val in REPLACE:
            if not replace_text(d, key, val):
                print("  ! không thấy:", key)
        fix_fonts(d)
        print("đã đặt placeholder", flush=True)
        d.SaveAs2(out, FileFormat=16)
        d.Close(False)
        print("Đã tạo", out)
    finally:
        wd.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
