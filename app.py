"""KHBD Tháng – ghép giáo án tuần thành kế hoạch bài dạy tháng."""
import ctypes
import os
import queue
import shutil
import sys
import tempfile
import threading
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, ttk

# pymupdf tự nạp "pymupdf.layout" (AI phân tích bố cục, cần onnxruntime + file model) nếu có cài.
# Tool không dùng tới, và trong exe thiếu file model nên sẽ lỗi khi khởi động -> chặn hẳn.
sys.modules["pymupdf.layout"] = None
import pymupdf  # noqa: E402

import settings
from docinfo import guess_month, read_week
from merger import MergeError, merge, resource_path
from settings_dialog import SettingsDialog

APP_TITLE = "KHBD Tháng – Ghép giáo án tuần"
FILE_TYPES = [("Giáo án Word", "*.docx *.wbk *.docm"), ("Tất cả", "*.*")]

# (khoá settings, placeholder trong mẫu, nhãn, tab)
FORM_FIELDS = [
    ("co_quan", "CO_QUAN", "Cơ quan chủ quản", "bia"),
    ("truong", "TRUONG", "Tên trường", "bia"),
    ("lop", "LOP", "Lớp", "bia"),
    ("giao_vien", "GIAO_VIEN", "Họ tên giáo viên", "bia"),
    ("nam_hoc", "NAM_HOC", "Năm học", "bia"),
    ("bgh_tieu_de", "BGH_TIEU_DE", "Tiêu đề BGH", "ky"),
    ("bgh_chuc_danh", "BGH_CHUC_DANH", "Chức danh người ký", "ky"),
    ("bgh_ten", "BGH_TEN", "Họ tên người ký BGH", "ky"),
    ("tt_tieu_de", "TT_TIEU_DE", "Tiêu đề tổ trưởng", "ky"),
    ("tt_ten", "TT_TEN", "Họ tên tổ trưởng", "ky"),
]


class PreviewPane(ttk.Frame):
    """Xem trước PDF: chỉ vẽ những trang đang hiển thị để nhẹ máy."""

    GAP = 12

    def __init__(self, master):
        super().__init__(master)
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 4))
        ttk.Button(bar, text="−", width=3, command=lambda: self.set_zoom(self.zoom / 1.2)).pack(side="left")
        ttk.Button(bar, text="+", width=3, command=lambda: self.set_zoom(self.zoom * 1.2)).pack(side="left", padx=2)
        ttk.Button(bar, text="Vừa khung", command=self.fit_width).pack(side="left", padx=(2, 10))
        self.lbl = ttk.Label(bar, text="Chưa có bản xem trước")
        self.lbl.pack(side="left")
        self.btn_open = ttk.Button(bar, text="Mở bằng trình đọc PDF", command=self.open_external, state="disabled")
        self.btn_open.pack(side="right")

        wrap = ttk.Frame(self)
        wrap.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(wrap, background="#8a8f98", highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self._yview)
        self.canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self._layout(keep=True))
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
        self.doc = None
        self.path = None
        self.zoom = None
        self.tops = []
        self.items = {}

    # ---- tải / đóng ----
    def load(self, pdf_path):
        self.close()
        self.doc = pymupdf.open(pdf_path)
        self.path = pdf_path
        self.btn_open.config(state="normal")
        self.zoom = None
        self.canvas.yview_moveto(0)
        self._layout()

    def close(self):
        self.canvas.delete("all")
        self.items.clear()
        if self.doc is not None:
            self.doc.close()
        self.doc = None
        self.btn_open.config(state="disabled")
        self.lbl.config(text="Chưa có bản xem trước")

    def open_external(self):
        if self.path and os.path.exists(self.path):
            os.startfile(self.path)

    # ---- bố cục ----
    def fit_width(self):
        if self.doc is None:
            return
        w = max(200, self.canvas.winfo_width() - 2 * self.GAP)
        widest = max(p.rect.width for p in self.doc)
        self.set_zoom(w / widest)

    def set_zoom(self, z):
        if self.doc is None:
            return
        self.zoom = max(0.2, min(4.0, z))
        self._layout(keep=True)

    def _layout(self, keep=False):
        if self.doc is None:
            return
        if self.zoom is None:
            w = max(200, self.canvas.winfo_width() - 2 * self.GAP)
            self.zoom = w / max(p.rect.width for p in self.doc)
        frac = self.canvas.yview()[0] if keep else 0
        self.canvas.delete("all")
        self.items.clear()
        cw = self.canvas.winfo_width()
        y = self.GAP
        self.tops = []
        for p in self.doc:
            w, h = p.rect.width * self.zoom, p.rect.height * self.zoom
            x = max(self.GAP, (cw - w) / 2)
            self.canvas.create_rectangle(x + 3, y + 3, x + w + 3, y + h + 3, fill="#5d626b", outline="")
            self.canvas.create_rectangle(x, y, x + w, y + h, fill="white", outline="#444")
            self.tops.append((y, h, x))
            y += h + self.GAP
        self.canvas.configure(scrollregion=(0, 0, cw, y))
        self.canvas.yview_moveto(frac)
        self._render_visible()

    def _yview(self, *args):
        self.canvas.yview(*args)
        self._render_visible()

    def _wheel(self, e):
        self.canvas.yview_scroll(int(-e.delta / 40), "units")
        self._render_visible()

    def _render_visible(self):
        if self.doc is None or not self.tops:
            return
        y0 = self.canvas.canvasy(0)
        y1 = y0 + self.canvas.winfo_height()
        visible = [i for i, (t, h, _) in enumerate(self.tops) if t + h >= y0 - 400 and t <= y1 + 400]
        for i in list(self.items):
            if i not in visible:
                self.canvas.delete(self.items.pop(i)[0])
        for i in visible:
            if i in self.items:
                continue
            t, h, x = self.tops[i]
            pix = self.doc[i].get_pixmap(matrix=pymupdf.Matrix(self.zoom, self.zoom), alpha=False)
            img = tk.PhotoImage(data=pix.tobytes("ppm"))
            item = self.canvas.create_image(x, t, anchor="nw", image=img)
            self.items[i] = (item, img)
        cur = next((i for i, (t, h, _) in enumerate(self.tops) if t + h > y0 + 40), 0)
        self.lbl.config(text=f"Trang {cur + 1} / {len(self.tops)}   ·   {int(self.zoom * 100)}%")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x820")
        self.minsize(1000, 640)
        self._set_icon()
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Big.TButton", font=("Segoe UI", 10, "bold"), padding=(10, 6))

        self.cfg = settings.load()
        self.weeks = []          # list[WeekInfo]
        self.vars = {}
        self.tt_auto = True      # tên tổ trưởng đang lấy tự động từ file
        self.queue = queue.Queue()
        self.busy = False
        self.preview_dir = tempfile.mkdtemp(prefix="khbd_xem_")
        self.preview_key = None  # dữ liệu đầu vào của bản xem trước hiện tại

        self._build()
        self._apply_cfg(self.cfg)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll)

    # ------------------------------------------------------------------ UI
    def _set_icon(self):
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except tk.TclError:
            pass

    def _build(self):
        outer = ttk.Frame(self, padding=8)
        outer.pack(fill="both", expand=True)
        left = ttk.Frame(outer, padding=(0, 0, 10, 0))
        left.pack(side="left", fill="y")
        right = ttk.Frame(outer)
        right.pack(side="left", fill="both", expand=True)

        # 1. file tuần
        f1 = ttk.LabelFrame(left, text=" 1. Chọn các file giáo án tuần ", padding=8)
        f1.pack(fill="x")
        self.tree = ttk.Treeview(f1, columns=("tuan", "ngay", "file"), show="headings", height=6,
                                 selectmode="browse")
        for col, txt, w in (("tuan", "Tuần", 50), ("ngay", "Từ – đến", 120), ("file", "Tên file", 260)):
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w, stretch=(col == "file"), anchor="w")
        self.tree.pack(fill="x")
        self.tree.tag_configure("err", foreground="#c00")
        b = ttk.Frame(f1)
        b.pack(fill="x", pady=(6, 0))
        ttk.Button(b, text="Thêm file...", command=self.add_files).pack(side="left")
        ttk.Button(b, text="Xoá", command=self.remove_file).pack(side="left", padx=4)
        ttk.Button(b, text="▲", width=3, command=lambda: self.move(-1)).pack(side="left")
        ttk.Button(b, text="▼", width=3, command=lambda: self.move(1)).pack(side="left", padx=4)
        ttk.Button(b, text="Sắp xếp theo tuần", command=self.sort_weeks).pack(side="left")
        self.lbl_warn = ttk.Label(f1, text="", foreground="#b35c00", wraplength=420, justify="left")
        self.lbl_warn.pack(fill="x", pady=(4, 0))

        # 2. tháng
        f2 = ttk.LabelFrame(left, text=" 2. Tháng báo cáo ", padding=8)
        f2.pack(fill="x", pady=(8, 0))
        f2.columnconfigure(1, weight=1)
        self.var_thang = tk.StringVar(value=str(date.today().month))
        ttk.Label(f2, text="Tháng").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Spinbox(f2, from_=1, to=12, width=4, textvariable=self.var_thang,
                    command=self._month_changed).grid(row=0, column=1, sticky="w")
        self.var_thang.trace_add("write", lambda *a: self._month_changed())
        # tên file xuất: tự sinh theo mẫu trong Cài đặt, nhưng cho sửa trực tiếp
        ttk.Label(f2, text="Tên file xuất").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        fn = ttk.Frame(f2)
        fn.grid(row=1, column=1, sticky="ew", pady=(6, 0))
        self.var_tenfile = tk.StringVar()
        self.tenfile_auto = True
        self._setting_fn = False
        ttk.Entry(fn, textvariable=self.var_tenfile).pack(side="left", fill="x", expand=True)
        ttk.Label(fn, text=".docx").pack(side="left", padx=(2, 4))
        ttk.Button(fn, text="↺", width=3, command=self._reset_filename).pack(side="left")
        self.var_tenfile.trace_add("write", self._fn_edited)

        # 3. thông tin
        f3 = ttk.LabelFrame(left, text=" 3. Thông tin trang bìa & chữ ký ", padding=8)
        f3.pack(fill="x", pady=(8, 0))
        nb = ttk.Notebook(f3)
        nb.pack(fill="x")
        tabs = {"bia": ttk.Frame(nb, padding=8), "ky": ttk.Frame(nb, padding=8)}
        nb.add(tabs["ky"], text="Chữ ký duyệt")
        nb.add(tabs["bia"], text="Trang bìa")
        rows = {"bia": 0, "ky": 0}
        for key, _ph, label, tab in FORM_FIELDS:
            fr = tabs[tab]
            fr.columnconfigure(1, weight=1)
            r = rows[tab]
            ttk.Label(fr, text=label).grid(row=r, column=0, sticky="w", padx=(0, 8), pady=2)
            var = tk.StringVar()
            self.vars[key] = var
            if key == "bgh_chuc_danh":
                w = ttk.Combobox(fr, textvariable=var, values=settings.CHUC_DANH_GOI_Y, width=34)
            else:
                w = ttk.Entry(fr, textvariable=var, width=36)
            w.grid(row=r, column=1, sticky="ew", pady=2)
            rows[tab] += 1
            if key == "bgh_ten":
                # dòng ngày ký
                r = rows[tab]
                ttk.Label(fr, text="Nơi, ngày ký").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=2)
                d = ttk.Frame(fr)
                d.grid(row=r, column=1, sticky="ew", pady=2)
                self.vars["dia_danh"] = tk.StringVar()
                ttk.Entry(d, textvariable=self.vars["dia_danh"], width=12).pack(side="left")
                today = date.today()
                self.var_d = tk.StringVar(value=str(today.day))
                self.var_m = tk.StringVar(value=str(today.month))
                self.var_y = tk.StringVar(value=str(today.year))
                for txt, var_, lo, hi, wd in ((" ngày", self.var_d, 1, 31, 3), (" tháng", self.var_m, 1, 12, 3),
                                              (" năm", self.var_y, 2020, 2100, 5)):
                    ttk.Label(d, text=txt).pack(side="left")
                    ttk.Spinbox(d, from_=lo, to=hi, width=wd, textvariable=var_).pack(side="left", padx=(2, 0))
                rows[tab] += 1
        self.vars["tt_ten"].trace_add("write", self._tt_edited)
        ttk.Label(tabs["ky"], text="Chức danh có thể chọn sẵn hoặc gõ tuỳ ý (gõ \\n để xuống dòng).",
                  foreground="#666").grid(row=rows["ky"], column=0, columnspan=2, sticky="w", pady=(4, 0))

        # 4. nút
        f4 = ttk.Frame(left)
        f4.pack(fill="x", pady=(10, 0))
        self.btn_prev = ttk.Button(f4, text="👁  Xem trước", style="Big.TButton", command=self.preview)
        self.btn_prev.pack(side="left", fill="x", expand=True)
        self.btn_export = ttk.Button(f4, text="💾  Xuất báo cáo tháng", style="Big.TButton", command=self.export)
        self.btn_export.pack(side="left", fill="x", expand=True, padx=(6, 0))
        f5 = ttk.Frame(left)
        f5.pack(fill="x", pady=(6, 0))
        ttk.Button(f5, text="⚙  Cài đặt mặc định", command=self.open_settings).pack(side="left")
        ttk.Button(f5, text="Lưu thông tin đang nhập làm mặc định",
                   command=self.save_as_default).pack(side="left", padx=6)

        self.progress = ttk.Progressbar(left, mode="determinate", value=0)
        self.progress.pack(fill="x", pady=(10, 2))
        self.lbl_status = ttk.Label(left, text="Sẵn sàng. Yêu cầu máy có Microsoft Word.", foreground="#333",
                                    wraplength=420)
        self.lbl_status.pack(fill="x")

        self.preview_pane = PreviewPane(right)
        self.preview_pane.pack(fill="both", expand=True)

    # ------------------------------------------------------------ settings
    def _apply_cfg(self, cfg):
        for key, var in self.vars.items():
            if key == "tt_ten":
                continue
            var.set(str(cfg.get(key, "")).replace("\n", "\\n"))
        self.tt_auto = not cfg.get("tt_ten")
        self._set_tt(cfg.get("tt_ten") or self._auto_tt())
        self._month_changed()

    def open_settings(self):
        def saved(cfg):
            self.cfg = cfg
            if messagebox.askyesno("Cài đặt", "Đã lưu. Áp dụng các giá trị mặc định mới vào form hiện tại?"):
                self._apply_cfg(cfg)
            self._month_changed()
        SettingsDialog(self, on_saved=saved)

    def save_as_default(self):
        for key, var in self.vars.items():
            if key == "tt_ten" and self.tt_auto:
                continue
            self.cfg[key] = var.get().replace("\\n", "\n")
        settings.save(self.cfg)
        self.status("Đã lưu thông tin đang nhập làm mặc định.")

    # ------------------------------------------------------------- tổ trưởng
    def _auto_tt(self):
        return next((w.to_truong for w in self.weeks if w.to_truong), "")

    def _set_tt(self, value):
        self._setting_tt = True
        self.vars["tt_ten"].set(value)
        self._setting_tt = False

    def _tt_edited(self, *a):
        if not getattr(self, "_setting_tt", False):
            self.tt_auto = False

    # ---------------------------------------------------------- file tuần
    def add_files(self):
        init = os.path.dirname(self.weeks[-1].path) if self.weeks else os.getcwd()
        paths = filedialog.askopenfilenames(title="Chọn file giáo án tuần", filetypes=FILE_TYPES, initialdir=init)
        if not paths:
            return
        have = {os.path.normcase(w.path) for w in self.weeks}
        for p in paths:
            if os.path.normcase(p) in have:
                continue
            self.weeks.append(read_week(p))
        self.sort_weeks()
        if self.tt_auto:
            self._set_tt(self._auto_tt())
        m, _y = guess_month(self.weeks)
        if m:
            self.var_thang.set(str(m))

    def remove_file(self):
        sel = self.tree.selection()
        if sel:
            del self.weeks[int(sel[0])]
            self._refresh()

    def move(self, d):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        j = i + d
        if 0 <= j < len(self.weeks):
            self.weeks[i], self.weeks[j] = self.weeks[j], self.weeks[i]
            self._refresh(select=j)

    def sort_weeks(self):
        self.weeks.sort(key=lambda w: (w.week is None, w.week or 0, w.path))
        self._refresh()

    def _refresh(self, select=None):
        self.tree.delete(*self.tree.get_children())
        for i, w in enumerate(self.weeks):
            ngay = f"{w.start:%d/%m} – {w.end:%d/%m}" if w.start and w.end else "?"
            self.tree.insert("", "end", iid=str(i), values=(w.week if w.week is not None else "?", ngay,
                                                             os.path.basename(w.path)),
                             tags=("err",) if w.errors else ())
        if select is not None:
            self.tree.selection_set(str(select))
        self.lbl_warn.config(text="\n".join(self._warnings()))

    def _warnings(self):
        out = []
        for w in self.weeks:
            for e in w.errors:
                out.append(f"✖ {os.path.basename(w.path)}: {e}")
        nums = [w.week for w in self.weeks if w.week is not None]
        dup = sorted({n for n in nums if nums.count(n) > 1})
        if dup:
            out.append("⚠ Trùng tuần: " + ", ".join(map(str, dup)))
        if nums and sorted(nums) != list(range(min(nums), min(nums) + len(nums))) and not dup:
            out.append("⚠ Các tuần không liên tiếp: " + ", ".join(map(str, sorted(nums))))
        if nums and nums != sorted(nums):
            out.append("⚠ Thứ tự file không theo số tuần (bấm 'Sắp xếp theo tuần' nếu cần).")
        try:
            m = int(self.var_thang.get())
            lech = [str(w.week) for w in self.weeks
                    if w.start and w.end and m not in (w.start.month, w.end.month)]
            if lech:
                out.append(f"⚠ Tuần {', '.join(lech)} không có ngày nào thuộc tháng {m}.")
        except ValueError:
            pass
        return out

    def _month_changed(self):
        if self.tenfile_auto:
            self._set_filename(self._pattern_name())
        if self.weeks:
            self.lbl_warn.config(text="\n".join(self._warnings()))

    def _set_filename(self, value):
        self._setting_fn = True
        self.var_tenfile.set(value)
        self._setting_fn = False

    def _fn_edited(self, *a):
        if not self._setting_fn:
            self.tenfile_auto = False

    def _reset_filename(self):
        self.tenfile_auto = True
        self._set_filename(self._pattern_name())

    def _file_name(self):
        name = self.var_tenfile.get().strip()
        if name.lower().endswith(".docx"):
            name = name[:-5]
        name = "".join("_" if c in '<>:"/\\|?*' else c for c in name).strip()
        return name or self._pattern_name()

    def _pattern_name(self):
        try:
            m = int(self.var_thang.get())
        except ValueError:
            return ""
        try:
            name = self.cfg.get("ten_file") or settings.DEFAULTS["ten_file"]
            name = name.format(thang=m)
        except (KeyError, IndexError, ValueError):
            name = f"KE HOACH BAI DAY THANG {m}"
        return "".join("_" if c in '<>:"/\\|?*' else c for c in name).strip()

    # -------------------------------------------------------------- ghép
    def _values(self):
        vals = {}
        for key, ph, _label, _tab in FORM_FIELDS:
            vals[ph] = self.vars[key].get().replace("\\n", "\n").strip()
        dd = self.vars["dia_danh"].get().strip()
        ngay = f"ngày {self.var_d.get().strip()} tháng {self.var_m.get().strip()} năm {self.var_y.get().strip()}"
        vals["NGAY_KY"] = f"{dd}, {ngay}" if dd else ngay.capitalize()
        return vals

    def _input_key(self):
        files = tuple((w.path, os.path.getmtime(w.path) if os.path.exists(w.path) else 0) for w in self.weeks)
        return files, tuple(sorted(self._values().items()))

    def _validate(self):
        if not self.weeks:
            messagebox.showwarning(APP_TITLE, "Hãy thêm ít nhất 1 file giáo án tuần.")
            return False
        bad = [w for w in self.weeks if w.errors]
        if bad:
            messagebox.showerror(APP_TITLE, "Có file không đúng mẫu giáo án tuần:\n\n" +
                                 "\n".join(f"• {os.path.basename(w.path)}: {'; '.join(w.errors)}" for w in bad))
            return False
        missing = [w.path for w in self.weeks if not os.path.exists(w.path)]
        if missing:
            messagebox.showerror(APP_TITLE, "Không tìm thấy file:\n" + "\n".join(missing))
            return False
        warns = [x for x in self._warnings() if x.startswith("⚠")]
        if warns and not messagebox.askyesno(APP_TITLE, "\n".join(warns) + "\n\nVẫn tiếp tục?"):
            return False
        return True

    def _run(self, out_docx, out_pdf, on_done):
        self.busy = True
        for b in (self.btn_prev, self.btn_export):
            b.state(["disabled"])
        self.progress.config(mode="indeterminate")
        self.progress.start(12)
        paths = [w.path for w in self.weeks]
        values = self._values()

        def work():
            try:
                merge(paths, values, out_docx, out_pdf, progress=lambda m: self.queue.put(("status", m)))
                self.queue.put(("done", on_done))
            except MergeError as e:
                self.queue.put(("error", str(e)))
            except Exception as e:  # noqa: BLE001 – báo mọi lỗi cho người dùng
                self.queue.put(("error", f"Lỗi không mong muốn: {e!r}"))

        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        try:
            while True:
                kind, data = self.queue.get_nowait()
                if kind == "status":
                    self.status(data)
                else:
                    self.busy = False
                    self.progress.stop()
                    self.progress.config(mode="determinate", value=0)
                    for b in (self.btn_prev, self.btn_export):
                        b.state(["!disabled"])
                    if kind == "done":
                        data()
                    else:
                        self.status("Có lỗi: " + data)
                        messagebox.showerror(APP_TITLE, data)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def status(self, msg):
        self.lbl_status.config(text=msg)

    def preview(self):
        if self.busy or not self._validate():
            return
        key = self._input_key()
        docx = os.path.join(self.preview_dir, "xem_truoc.docx")
        pdf = os.path.join(self.preview_dir, "xem_truoc.pdf")
        self.preview_pane.close()

        def done():
            self.preview_key = key
            self.preview_pane.load(pdf)
            self.status(f"Đã tạo bản xem trước ({len(self.preview_pane.doc)} trang). "
                        "Kiểm tra rồi bấm 'Xuất báo cáo tháng'.")
        self._run(docx, pdf, done)

    def export(self):
        if self.busy or not self._validate():
            return
        init_dir = self.cfg.get("thu_muc_xuat") or os.path.dirname(self.weeks[0].path)
        out = filedialog.asksaveasfilename(title="Lưu báo cáo tháng", defaultextension=".docx",
                                           initialdir=init_dir, initialfile=self._file_name() + ".docx",
                                           filetypes=[("Word", "*.docx")])
        if not out:
            return
        want_pdf = bool(self.cfg.get("xuat_pdf", True))
        out_pdf = os.path.splitext(out)[0] + ".pdf" if want_pdf else None

        def finished():
            msg = f"Đã xuất:\n{out}" + (f"\n{out_pdf}" if out_pdf else "")
            self.status(msg.replace("\n", "  "))
            if messagebox.askyesno(APP_TITLE, msg + "\n\nMở thư mục chứa file?"):
                os.startfile(os.path.dirname(out))

        # Dùng lại bản xem trước nếu dữ liệu không đổi -> xuất ngay
        if self.preview_key is not None and self.preview_key == self._input_key():
            try:
                shutil.copyfile(os.path.join(self.preview_dir, "xem_truoc.docx"), out)
                if out_pdf:
                    shutil.copyfile(os.path.join(self.preview_dir, "xem_truoc.pdf"), out_pdf)
                finished()
                return
            except OSError as e:
                messagebox.showerror(APP_TITLE, f"Không lưu được file (đang mở trong Word?):\n{e}")
                return
        self._run(out, out_pdf, finished)

    def _on_close(self):
        if self.busy and not messagebox.askyesno(APP_TITLE, "Đang ghép file, vẫn thoát?"):
            return
        self.preview_pane.close()
        shutil.rmtree(self.preview_dir, ignore_errors=True)
        self.destroy()


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass
    App().mainloop()


if __name__ == "__main__":
    main()
