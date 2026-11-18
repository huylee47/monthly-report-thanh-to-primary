@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === Cai thu vien can thiet ===
python -m pip install -r requirements.txt
echo === Bien dich KHBD_Thang.exe bang Nuitka (lan dau mat 10-30 phut) ===
python -m nuitka --onefile --assume-yes-for-downloads ^
  --windows-console-mode=disable ^
  --enable-plugin=tk-inter ^
  --nofollow-import-to=pymupdf4llm,pymupdf_layout,pymupdf.layout,onnxruntime,pandas,numpy,numba,scipy,sklearn,matplotlib,PIL,IPython,pytest ^
  --include-data-files=template.docx=template.docx ^
  --include-data-files=icon.ico=icon.ico ^
  --windows-icon-from-ico=icon.ico ^
  --product-name="KHBD Thang" --file-description="Ghep giao an tuan thanh KHBD thang" ^
  --product-version=1.0.0 --file-version=1.0.0 ^
  --company-name="Le Khanh Huy - Zalo 0938194069" --copyright="Thiet ke: Le Khanh Huy - Zalo 0938194069" ^
  --output-dir=build --output-filename=KHBD_Thang.exe ^
  --remove-output ^
  app.py
if errorlevel 1 (echo LOI khi build & pause & exit /b 1)
echo === Xong: build\KHBD_Thang.exe ===
pause
