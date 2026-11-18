"""Cửa sổ Cài đặt: chỉnh giá trị mặc định cho các ô nhập."""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import settings


class SettingsDialog(tk.Toplevel):
    def __init__(self, master, on_saved=None):
        super().__init__(master)
        self.title("Cài đặt mặc định")
        self.transient(master)
        self.resizable(True, False)
        self.on_saved = on_saved
        self.data = settings.load()
        self.vars = {}

        body = ttk.Frame(self, padding=12)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)

        row = 0
        group = None
        for key, label, grp in settings.FIELDS:
            if grp != group:
                group = grp
                ttk.Label(body, text=grp, font=("Segoe UI", 10, "bold")).grid(
                    row=row, column=0, columnspan=3, sticky="w", pady=(10 if row else 0, 4))
                row += 1
            ttk.Label(body, text=label).grid(row=row, column=0, sticky="w", padx=(8, 8), pady=2)
            var = tk.StringVar(value=str(self.data.get(key, "")).replace("\n", "\\n"))
            self.vars[key] = var
            if key == "bgh_chuc_danh":
                w = ttk.Combobox(body, textvariable=var, values=settings.CHUC_DANH_GOI_Y)
            else:
                w = ttk.Entry(body, textvariable=var, width=48)
            w.grid(row=row, column=1, sticky="ew", pady=2)
            if key == "thu_muc_xuat":
                ttk.Button(body, text="Chọn...", command=self._pick_dir).grid(row=row, column=2, padx=(4, 0))
            row += 1

        self.var_pdf = tk.BooleanVar(value=bool(self.data.get("xuat_pdf", True)))
        ttk.Checkbutton(body, text="Khi xuất báo cáo, tạo kèm file PDF",
                        variable=self.var_pdf).grid(row=row, column=0, columnspan=3, sticky="w", pady=(8, 0))
        row += 1
        ttk.Label(body, text="Mẹo: chức danh có thể nhiều dòng, gõ \\n để xuống dòng "
                             "(VD: KT. HIỆU TRƯỞNG\\nPHÓ HIỆU TRƯỞNG).",
                  foreground="#666").grid(row=row, column=0, columnspan=3, sticky="w", pady=(4, 0))

        btns = ttk.Frame(self, padding=(12, 0, 12, 12))
        btns.pack(fill="x")
        ttk.Button(btns, text="Khôi phục mặc định gốc", command=self._reset).pack(side="left")
        ttk.Button(btns, text="Huỷ", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Lưu", command=self._save).pack(side="right", padx=6)

        self.grab_set()
        self.wait_visibility()
        self.focus_set()

    def _pick_dir(self):
        d = filedialog.askdirectory(parent=self, title="Chọn thư mục lưu mặc định")
        if d:
            self.vars["thu_muc_xuat"].set(d)

    def _reset(self):
        if not messagebox.askyesno("Khôi phục", "Đưa mọi ô về giá trị mặc định ban đầu?", parent=self):
            return
        for k, v in self.vars.items():
            v.set(settings.DEFAULTS.get(k, ""))
        self.var_pdf.set(settings.DEFAULTS["xuat_pdf"])

    def _save(self):
        for k, v in self.vars.items():
            self.data[k] = v.get().replace("\\n", "\n")
        self.data["xuat_pdf"] = self.var_pdf.get()
        if "{thang}" not in self.data["ten_file"]:
            if not messagebox.askyesno(
                    "Tên file", "Mẫu tên file không có {thang} nên mọi tháng sẽ trùng tên. Vẫn lưu?",
                    parent=self):
                return
        settings.save(self.data)
        if self.on_saved:
            self.on_saved(self.data)
        self.destroy()
