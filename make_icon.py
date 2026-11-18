"""Tạo icon.ico (PNG nhúng trong ICO) – chỉ dùng khi phát triển."""
import struct
import pymupdf

def png(size):
    doc = pymupdf.open()
    p = doc.new_page(width=64, height=64)
    p.draw_rect(pymupdf.Rect(4, 2, 60, 62), color=None, fill=(0.09, 0.36, 0.67), radius=0.15)
    p.draw_rect(pymupdf.Rect(14, 10, 50, 54), color=None, fill=(1, 1, 1), radius=0.08)
    for i, y in enumerate((20, 28, 36, 44)):
        p.draw_line((19, y), (45 - (8 if i == 3 else 0), y), color=(0.09, 0.36, 0.67), width=3)
    p.draw_circle((48, 50), 9, color=None, fill=(0.95, 0.6, 0.1))
    p.insert_text((43.5, 54.5), "+", fontsize=15, color=(1, 1, 1), fontname="helv")
    z = size / 64
    return p.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=True).tobytes("png")

sizes = [16, 32, 48, 256]
imgs = [png(s) for s in sizes]
head = struct.pack("<HHH", 0, 1, len(sizes))
off = 6 + 16 * len(sizes)
dirs = b""
for s, d in zip(sizes, imgs):
    dirs += struct.pack("<BBBBHHII", s % 256, s % 256, 0, 0, 1, 32, len(d), off)
    off += len(d)
open("icon.ico", "wb").write(head + dirs + b"".join(imgs))
open("icon_preview.png", "wb").write(imgs[-1])
print("ok")
