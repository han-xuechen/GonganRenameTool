# -*- coding: utf-8 -*-
# 结论性文书合并移动工具V3.3.12.py
#
# 在 V3.3.11 基础上新增：
# - 运行结束弹窗后，自动生成并打开 “核查清单.xlsx”（Excel 2007 兼容 .xlsx）
# - 清单包含列：类别(JPG/PDF)｜档号｜原因｜详情/路径
# - 所有“跳过/失败”的场景均会记录一条，便于后续核对

import os, re, sys, tempfile, shutil, threading, time, subprocess
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import pandas as pd
from PIL import Image, ImageTk
import pytesseract
from PyPDF2 import PdfMerger

# ================== 主题 / 常量 ==================
THEME_PRIMARY   = "#14b8a6"
THEME_PRIMARY_D = "#0f766e"
THEME_BG        = "#f7f9fb"
THEME_FG        = "#111827"
THEME_MUTED     = "#6b7280"
ENTRY_BG        = "#ffffff"
TEXT_BG         = "#ffffff"
BORDER          = "#e5e7eb"

DEFAULT_LANG    = "chi_sim"   # 固定中文
PSM_FIXED       = 6           # 固定 PSM=6
ALLOWED_EXTS    = (".jpg", ".jpeg")
LOGO_MAX_PX     = 160

SYS_TESS_EXE    = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# 运行时全局
CUR_TESS_EXE   = None
CUR_TESSDATA   = None
CUR_TESSCFG    = f"--psm {PSM_FIXED}"

# ================== 基础工具 ==================
def _base_dir() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

def resource_path(name: str) -> str:
    return str((_base_dir() / name).resolve())

def _norm(p: str | Path) -> str:
    return os.path.normpath(str(p)).strip().strip(' "\'')

def natural_keys(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]

def list_images_sorted(folder: str):
    p = Path(folder)
    files = [x for x in p.iterdir() if x.suffix.lower() in ALLOWED_EXTS]
    files.sort(key=lambda x: natural_keys(x.name))
    return [str(x) for x in files]

def parse_ranges(rng_str):
    if pd.isna(rng_str):
        return []
    s = str(rng_str).strip()
    if not s:
        return []
    parts = re.split(r'[；;，,、\s]+', s)
    pages = []
    for token in (p.strip() for p in parts if p):
        if re.search(r'[-~—～至]', token):
            a, b = re.split(r'[-~—～至]', token, maxsplit=1)
            try:
                start = int(re.sub(r'\D', '', a))
                end = int(re.sub(r'\D', '', b))
                pages.extend(range(min(start, end), max(start, end) + 1))
            except:
                pass
        else:
            try:
                pages.append(int(re.sub(r'\D', '', token)))
            except:
                pass
    seen, uniq = set(), []
    for x in pages:
        if x not in seen:
            seen.add(x); uniq.append(x)
    return uniq

def prepare_log_file():
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    target_dir = Path("D:/") if Path("D:/").exists() else (Path.home() / "Documents")
    target_dir = target_dir / "OCR_Logs"
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        target_dir = _base_dir() / "OCR_Logs"
        target_dir.mkdir(parents=True, exist_ok=True)
    return _norm(target_dir / f"log_{ts}.txt")

# ================== Tesseract ==================
def _apply_tesseract(exe_path: str | Path) -> bool:
    global CUR_TESS_EXE, CUR_TESSDATA
    exe = Path(_norm(exe_path))
    if not exe.is_file():
        return False
    tdata = exe.parent / "tessdata"
    if not tdata.is_dir():
        return False
    CUR_TESS_EXE  = _norm(exe)
    CUR_TESSDATA  = _norm(tdata)
    pytesseract.pytesseract.tesseract_cmd = CUR_TESS_EXE
    os.environ["TESSDATA_PREFIX"] = CUR_TESSDATA
    return True

def _tess_ready() -> bool:
    if not (CUR_TESS_EXE and os.path.isfile(CUR_TESS_EXE) and CUR_TESSDATA and os.path.isdir(CUR_TESSDATA)):
        return False
    try:
        r = subprocess.run(
            [CUR_TESS_EXE, "--list-langs", "--tessdata-dir", CUR_TESSDATA],
            capture_output=True, text=True, timeout=8, env=os.environ
        )
        return (r.returncode == 0) and (("chi_sim" in (r.stdout+r.stderr)) or ("eng" in (r.stdout+r.stderr)))
    except Exception:
        return False

# ================== 应用 ==================
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("结论性文书合并移动工具 V3.3.12")
        root.geometry("980x760")
        root.configure(bg=THEME_BG)
        self._apply_theme()

        # 记录核查项（跳过/失败）
        self.check_items = []   # 每一项：{"类别": "JPG/PDF", "档号": str, "原因": str, "详情/路径": str}

        # 窗口图标
        try:
            ico = resource_path("logo.ico")
            if os.path.exists(ico):
                root.iconbitmap(ico)
        except Exception:
            pass

        # 顶部：左LOGO + 右标题区域（居中）
        head = tk.Frame(root, bg=THEME_BG)
        head.pack(fill="x", padx=12, pady=(10, 6))
        head.grid_columnconfigure(0, weight=0)
        head.grid_columnconfigure(1, weight=1)

        self.logo_label = tk.Label(head, bg=THEME_BG)
        self.logo_label.grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.logo_img = None
        self.root.after(10, self._load_logo_async, self.logo_label)

        title_area = tk.Frame(head, bg=THEME_BG)
        title_area.grid(row=0, column=1, sticky="nsew")
        tk.Label(
            title_area, text="结论性文书 合并移动工具",
            font=("Microsoft YaHei", 18, "bold"), bg=THEME_BG, fg=THEME_FG
        ).pack(pady=2)

        # 变量
        self.excel_path      = tk.StringVar()
        self.image_root      = tk.StringVar()
        self.output_pdf_dir  = tk.StringVar()
        self.copy_target_dir = tk.StringVar()
        self.tesseract_path  = tk.StringVar(value=SYS_TESS_EXE if os.path.exists(SYS_TESS_EXE) else "")

        # 表单
        form = tk.Frame(root, bg=THEME_BG, highlightbackground=BORDER, highlightthickness=1, bd=0)
        form.pack(fill="x", padx=12, pady=8)
        for i in (0,1,2): form.columnconfigure(i, weight=(0,1,0)[i])
        ROW_PADY = 10

        def add_row(row, label, var, btn_text, cmd):
            tk.Label(form, text=label, width=14, anchor="e", bg=THEME_BG, fg=THEME_FG)\
                .grid(row=row, column=0, padx=10, pady=ROW_PADY, sticky="e")
            tk.Entry(form, textvariable=var, bg=ENTRY_BG, fg=THEME_FG, relief="solid", bd=1)\
                .grid(row=row, column=1, padx=6, pady=ROW_PADY, sticky="we")
            ttk.Button(form, text=btn_text, command=cmd)\
                .grid(row=row, column=2, padx=10, pady=ROW_PADY, sticky="w")

        add_row(0, "Excel文件：",    self.excel_path,     "打开Excel", self.choose_excel)
        add_row(1, "原图像根目录：",  self.image_root,     "选择目录",   self.choose_img_root)
        add_row(2, "PDF输出目录：",   self.output_pdf_dir, "选择目录",   self.choose_pdf_out)
        add_row(3, "图片复制到：",    self.copy_target_dir,"选择目录",   self.choose_copy_target)

        tk.Label(form, text="Tesseract路径：", width=14, anchor="e", bg=THEME_BG, fg=THEME_FG)\
            .grid(row=4, column=0, padx=10, pady=ROW_PADY, sticky="e")
        tk.Entry(form, textvariable=self.tesseract_path, bg=ENTRY_BG, fg=THEME_FG, relief="solid", bd=1)\
            .grid(row=4, column=1, padx=6,  pady=ROW_PADY, sticky="we")
        ttk.Button(form, text="浏览", command=self.choose_tesseract)\
            .grid(row=4, column=2, padx=10, pady=ROW_PADY, sticky="w")

        # 操作按钮
        bar = tk.Frame(root, bg=THEME_BG); bar.pack(fill="x", padx=12, pady=(6, 8))
        for col, w in enumerate((2,1,1,1,1)): bar.grid_columnconfigure(col, weight=w)
        self.btn_both = ttk.Button(bar, text="复制 + 生成PDF",
                                   command=lambda: self.run(do_copy=True, do_pdf=True),
                                   style="Primary.TButton")
        self.btn_both.grid(row=0, column=0, padx=6, sticky="we")
        self.btn_copy = ttk.Button(bar, text="只复制图片",
                                   command=lambda: self.run(do_copy=True, do_pdf=False))
        self.btn_copy.grid(row=0, column=1, padx=6, sticky="we")
        self.btn_pdf  = ttk.Button(bar, text="只生成PDF",
                                   command=lambda: self.run(do_copy=False, do_pdf=True))
        self.btn_pdf.grid(row=0, column=2, padx=6, sticky="we")
        ttk.Button(bar, text="打开PDF目录", command=lambda: self.open_dir(self.output_pdf_dir.get()))\
            .grid(row=0, column=3, padx=6, sticky="we")
        ttk.Button(bar, text="打开复制目录", command=lambda: self.open_dir(self.copy_target_dir.get()))\
            .grid(row=0, column=4, padx=6, sticky="we")

        # 进度
        prog = tk.Frame(root, bg=THEME_BG); prog.pack(fill="x", padx=12, pady=(4,2))
        tk.Label(prog, text="总进度：", width=10, anchor="e", bg=THEME_BG, fg=THEME_FG).pack(side="left")
        self.pb_total = ttk.Progressbar(prog, length=650, mode="determinate",
                                        style="Primary.Horizontal.TProgressbar")
        self.pb_total.pack(side="left", fill="x", expand=True, padx=6)
        self.pb_total_val = tk.StringVar(value="0/0")
        tk.Label(prog, textvariable=self.pb_total_val, width=10, anchor="w",
                 bg=THEME_BG, fg=THEME_MUTED).pack(side="left")

        prog2 = tk.Frame(root, bg=THEME_BG); prog2.pack(fill="x", padx=12, pady=(0,6))
        tk.Label(prog2, text="当前档号：", width=10, anchor="e", bg=THEME_BG, fg=THEME_FG).pack(side="left")
        self.pb_item = ttk.Progressbar(prog2, length=650, mode="determinate",
                                       style="Primary.Horizontal.TProgressbar")
        self.pb_item.pack(side="left", fill="x", expand=True, padx=6)
        self.pb_item_val = tk.StringVar(value="0/0")
        tk.Label(prog2, textvariable=self.pb_item_val, width=10, anchor="w",
                 bg=THEME_BG, fg=THEME_MUTED).pack(side="left")

        # 日志
        log_frame = tk.Frame(root, bg=THEME_BG)
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        tk.Label(log_frame, text="日志：", bg=THEME_BG, fg=THEME_FG).pack(anchor="w")
        self.log_path = prepare_log_file()
        path_bar = tk.Frame(log_frame, bg=THEME_BG); path_bar.pack(fill="x", pady=(0,6))
        tk.Label(path_bar, text="当前日志文件：", bg=THEME_BG, fg=THEME_MUTED).pack(side="left")
        self.log_path_var = tk.StringVar(value=self.log_path)
        tk.Entry(path_bar, textvariable=self.log_path_var, bg=ENTRY_BG, fg=THEME_MUTED, bd=1, relief="solid")\
            .pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(path_bar, text="打开日志", command=self.open_log_file).pack(side="left")
        self.log = ScrolledText(log_frame, height=14, bg=TEXT_BG, fg=THEME_FG, insertbackground=THEME_FG)
        self.log.pack(fill="both", expand=True)

        # 启动提示
        self._log("准备就绪：依次选择 Excel、原图像根目录、PDF 输出目录、图片复制目录…")
        self._log(f"日志已启动，自动保存到：{self.log_path}")

    # 样式
    def _apply_theme(self):
        style = ttk.Style()
        try: style.theme_use("clam")
        except: pass
        style.configure("TButton", padding=(10,6))
        style.configure("Primary.TButton", background=THEME_PRIMARY, foreground="white",
                        padding=(12,8), borderwidth=0)
        style.map("Primary.TButton", background=[("active", THEME_PRIMARY_D)])
        style.configure("Primary.Horizontal.TProgressbar",
                        troughcolor="#e5e7eb", bordercolor="#e5e7eb",
                        background=THEME_PRIMARY, lightcolor=THEME_PRIMARY, darkcolor=THEME_PRIMARY)

    # Logo
    def _load_logo_async(self, label: tk.Label):
        try:
            p = resource_path("logo.png")
            if not os.path.exists(p): return
            img = Image.open(p).convert("RGBA")
            w, h = img.size
            scale = min(1.0, LOGO_MAX_PX / max(w, h))
            if scale < 1.0:
                img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)
            label.configure(image=self.logo_img)
        except Exception:
            pass

    # 日志
    def _log(self, msg):
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        self.log.insert("end", line + "\n")
        self.log.see("end")
        try:
            Path(self.log_path).parent.mkdir(parents=True, exist_ok=True)
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass

    def _warn(self, msg, kind=None, danghao=None, detail=None):
        """高亮日志，并可顺便把该条写入核查清单。"""
        self._log(f"!!! {msg}")
        if kind and danghao:
            self.check_items.append({"类别": kind, "档号": danghao, "原因": msg, "详情/路径": detail or ""})

    def open_log_file(self):
        p = self.log_path
        if p and os.path.exists(p): os.startfile(p)
        else: messagebox.showinfo("提示", "日志文件不存在。")

    # 选择器
    def choose_excel(self):
        p = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if p: self.excel_path.set(_norm(p)); self._log(f"已选择Excel：{self.excel_path.get()}")

    def choose_img_root(self):
        p = filedialog.askdirectory(title="选择原图像根目录")
        if p: self.image_root.set(_norm(p)); self._log(f"已选择原图像根目录：{self.image_root.get()}")

    def choose_pdf_out(self):
        p = filedialog.askdirectory(title="选择PDF输出目录")
        if p: self.output_pdf_dir.set(_norm(p)); self._log(f"已选择PDF输出目录：{self.output_pdf_dir.get()}")

    def choose_copy_target(self):
        p = filedialog.askdirectory(title="选择图片复制目录")
        if p: self.copy_target_dir.set(_norm(p)); self._log(f"已选择图片复制目录：{self.copy_target_dir.get()}")

    def choose_tesseract(self):
        p = filedialog.askopenfilename(title="选择 tesseract.exe",
                                       filetypes=[("tesseract.exe", "tesseract*.exe"), ("所有文件", "*.*")])
        if p:
            if _apply_tesseract(p):
                self.tesseract_path.set(_norm(p))
                self._log(f"已设定 Tesseract：{self.tesseract_path.get()}")
                self._log(f"当前 TESSDATA：{CUR_TESSDATA}")
            else:
                self.tesseract_path.set(_norm(p))
                self._warn("tesseract 同级未发现 tessdata 目录，可能无法 OCR。")

    # 打开目录
    def open_dir(self, d):
        d = (d or "").strip()
        if d and os.path.isdir(d): os.startfile(d)
        else: messagebox.showinfo("提示", "请先选择有效目录。")

    # 执行
    def run(self, do_copy: bool, do_pdf: bool):
        if not self.excel_path.get().strip():  return messagebox.showwarning("提示", "请先选择 Excel。")
        if not self.image_root.get().strip():  return messagebox.showwarning("提示", "请先选择 原图像根目录。")
        if do_pdf and not self.output_pdf_dir.get().strip():  return messagebox.showwarning("提示", "请选择 PDF 输出目录。")
        if do_copy and not self.copy_target_dir.get().strip(): return messagebox.showwarning("提示", "请选择 图片复制目录。")
        if not do_copy and not do_pdf:         return messagebox.showwarning("提示", "请至少选择一项操作。")

        if self.tesseract_path.get().strip(): _apply_tesseract(self.tesseract_path.get().strip())
        if do_pdf and not _tess_ready():
            return messagebox.showerror("错误", "未检测到可用的 Tesseract 或 tessdata。\n请确认安装并选择正确的 tesseract.exe（同级需有 tessdata）。")

        for b in (self.btn_both, self.btn_copy, self.btn_pdf): b.config(state="disabled")
        self.log_path = prepare_log_file(); self.log_path_var.set(self.log_path)
        self._log("=== 新任务开始 ===")
        self._set_total(0, 1); self._set_item(0, 1)

        threading.Thread(target=self._worker, args=(do_copy, do_pdf), daemon=True).start()

    # 进度条
    def _set_total(self, cur, total):
        self.pb_total["maximum"] = max(total, 1)
        self.pb_total["value"]   = min(cur, total)
        self.pb_total_val.set(f"{cur}/{total}")

    def _set_item(self, cur, total):
        self.pb_item["maximum"] = max(total, 1)
        self.pb_item["value"]   = min(cur, total)
        self.pb_item_val.set(f"{cur}/{total}")

    # 核心工作线程
    def _worker(self, do_copy: bool, do_pdf: bool):
        jpg_success = jpg_skipped = jpg_failed = 0
        pdf_success = pdf_skipped = pdf_failed = 0

        try:
            df = pd.read_excel(self.excel_path.get().strip(), engine="openpyxl", dtype=str)

            rng_col = None
            for cand in ("结论文书的页码范围", "法律结论文书的页码范围"):
                if cand in df.columns: rng_col = cand; break
            if ("档号" not in df.columns) or (rng_col is None):
                messagebox.showerror("错误", "Excel需包含列：‘档号’ 与 ‘结论文书的页码范围’（或旧名‘法律结论文书的页码范围’）")
                return

            df = df[["档号", rng_col]].dropna(subset=["档号"]).copy()
            df = df.sort_values(by="档号", key=lambda s: s.map(natural_keys)).reset_index(drop=True)

            img_root = self.image_root.get().strip()
            pdf_out  = self.output_pdf_dir.get().strip()
            copy_out = self.copy_target_dir.get().strip()
            if do_pdf: Path(pdf_out).mkdir(parents=True, exist_ok=True)
            if do_copy: Path(copy_out).mkdir(parents=True, exist_ok=True)

            total = len(df); done = 0
            self._set_total(0, total)
            self._log(f"开始处理（{'复制+PDF' if (do_copy and do_pdf) else ('仅复制' if do_copy else '仅PDF')}），共 {total} 个档号…")

            for _, row in df.iterrows():
                danghao = str(row["档号"]).strip()
                rng_str = row[rng_col]
                folder  = _norm(Path(img_root) / danghao)
                if not os.path.isdir(folder):
                    self._warn(f"档号目录不存在：{folder}", kind="JPG", danghao=danghao, detail=folder)
                    if do_pdf:
                        self._warn(f"档号目录不存在：{folder}", kind="PDF", danghao=danghao, detail=folder)
                    jpg_failed += int(do_copy); pdf_failed += int(do_pdf)
                    done += 1; self._set_total(done, total); continue

                all_imgs = list_images_sorted(folder)
                if not all_imgs:
                    self._warn(f"无 JPG 图片：{folder}", kind="JPG", danghao=danghao, detail=folder)
                    if do_pdf:
                        self._warn(f"无 JPG 图片：{folder}", kind="PDF", danghao=danghao, detail=folder)
                    jpg_failed += int(do_copy); pdf_failed += int(do_pdf)
                    done += 1; self._set_total(done, total); continue

                picks = parse_ranges(rng_str)
                if not picks:
                    self._warn(f"页码范围为空", kind="JPG", danghao=danghao, detail=str(rng_str))
                    if do_pdf: self._warn(f"页码范围为空", kind="PDF", danghao=danghao, detail=str(rng_str))
                    jpg_failed += int(do_copy); pdf_failed += int(do_pdf)
                    done += 1; self._set_total(done, total); continue

                valid_pages = [p for p in picks if 1 <= p <= len(all_imgs)]
                if not valid_pages:
                    self._warn(f"页码越界（总 {len(all_imgs)} 张）", kind="JPG", danghao=danghao, detail=str(picks))
                    if do_pdf: self._warn(f"页码越界（总 {len(all_imgs)} 张）", kind="PDF", danghao=danghao, detail=str(picks))
                    jpg_failed += int(do_copy); pdf_failed += int(do_pdf)
                    done += 1; self._set_total(done, total); continue

                targets = [all_imgs[p-1] for p in valid_pages]
                self._log(f"▶ 处理：{danghao}  选页 {valid_pages}")

                self._set_item(0, len(targets))
                item_done = 0

                # ---------- JPG：保留原文件名，不加序号 ----------
                if do_copy:
                    copy_dir = Path(copy_out) / danghao
                    try:
                        copy_dir.mkdir(parents=True, exist_ok=True)
                        copied, skipped, errors = 0, 0, 0
                        for src in targets:
                            try:
                                base = os.path.basename(src)
                                dst = copy_dir / base
                                if dst.exists():
                                    skipped += 1
                                    self._warn(f"JPG已存在，跳过：{dst}", kind="JPG", danghao=danghao, detail=str(dst))
                                else:
                                    shutil.copy2(src, dst)
                                    copied += 1
                            except Exception as e:
                                errors += 1
                                self._warn(f"复制失败：{src} ({e})", kind="JPG", danghao=danghao, detail=str(src))
                        if copied > 0:
                            jpg_success += 1
                            self._log(f"📷 复制完成：新增 {copied} 张，跳过 {skipped} 张，失败 {errors} 张 -> {copy_dir}")
                        elif skipped > 0 and copied == 0:
                            jpg_skipped += 1
                            self._warn(f"本卷 JPG 全部已存在，未新增：{copy_dir}", kind="JPG", danghao=danghao, detail=str(copy_dir))
                        else:
                            jpg_failed += 1
                            self._warn(f"本卷 JPG 复制失败", kind="JPG", danghao=danghao, detail=str(copy_dir))
                    except Exception as e:
                        jpg_failed += 1
                        self._warn(f"创建JPG子目录失败：{copy_dir} ({e})", kind="JPG", danghao=danghao, detail=str(copy_dir))

                # ---------- PDF：按档号建子目录；同名PDF跳过 ----------
                if do_pdf and _tess_ready():
                    workdir = Path(tempfile.mkdtemp(prefix="ocrpdf_"))
                    part_pdfs = []
                    try:
                        for p, img_path in zip(valid_pages, targets):
                            try:
                                with Image.open(img_path) as im:
                                    if im.mode not in ("RGB", "L"):
                                        im = im.convert("RGB")
                                    pdf_bytes = pytesseract.image_to_pdf_or_hocr(
                                        im, extension="pdf", lang=DEFAULT_LANG, config=CUR_TESSCFG
                                    )
                                out_page = workdir / f"p_{p}.pdf"
                                with open(out_page, "wb") as f:
                                    f.write(pdf_bytes)
                                part_pdfs.append(str(out_page))
                            except Exception as e:
                                self._warn(f"OCR失败：{img_path} ({e})", kind="PDF", danghao=danghao, detail=str(img_path))
                            item_done += 1; self._set_item(item_done, len(targets))

                        if part_pdfs:
                            out_dir = Path(pdf_out) / danghao
                            if not out_dir.exists():
                                try:
                                    out_dir.mkdir(parents=True, exist_ok=True)
                                    self._log(f"📁 已创建PDF子目录：{out_dir}")
                                except Exception as ce:
                                    pdf_failed += 1
                                    self._warn(f"创建PDF子目录失败：{out_dir} ({ce})", kind="PDF", danghao=danghao, detail=str(out_dir))
                                    done += 1; self._set_total(done, total)
                                    shutil.rmtree(workdir, ignore_errors=True)
                                    continue

                            out_path = out_dir / f"{danghao}.pdf"
                            if out_path.exists():
                                pdf_skipped += 1
                                self._warn(f"PDF已存在，跳过生成：{out_path}（请核对检查）", kind="PDF", danghao=danghao, detail=str(out_path))
                            else:
                                merger = PdfMerger()
                                for pth in part_pdfs: merger.append(pth)
                                try:
                                    with open(out_path, "wb") as f: merger.write(f)
                                    pdf_success += 1
                                    self._log(f"✅ 生成PDF：{out_path}")
                                except Exception as we:
                                    pdf_failed += 1
                                    self._warn(f"写入PDF失败：{out_path} ({we})", kind="PDF", danghao=danghao, detail=str(out_path))
                                finally:
                                    merger.close()
                        else:
                            pdf_failed += 1
                            self._warn(f"没有成功的页可合并", kind="PDF", danghao=danghao, detail=str(valid_pages))
                    finally:
                        shutil.rmtree(workdir, ignore_errors=True)
                elif do_pdf:
                    pdf_failed += 1
                    self._warn("Tesseract 未就绪，无法生成PDF。", kind="PDF", danghao=danghao, detail="Tesseract not ready")

                done += 1
                self._set_total(done, total)

            # ----------- 任务汇总 -----------
            summary = (
                "=== 任务汇总（卷级） ===\n"
                f"JPG：成功 {jpg_success} 卷；跳过 {jpg_skipped} 卷；失败 {jpg_failed} 卷\n"
                f"PDF：成功 {pdf_success} 卷；跳过 {pdf_skipped} 卷；失败 {pdf_failed} 卷\n"
            )
            self._log(summary)
            messagebox.showinfo(
                "运行结果",
                f"JPG：成功 {jpg_success} 卷；跳过 {jpg_skipped} 卷；失败 {jpg_failed} 卷\n"
                f"PDF：成功 {pdf_success} 卷；跳过 {pdf_skipped} 卷；失败 {pdf_failed} 卷\n\n"
                f"详情见日志：\n{self.log_path}"
            )

            # ----------- 生成并打开“核查清单.xlsx” -----------
            if self.check_items:
                df_check = pd.DataFrame(self.check_items, columns=["类别", "档号", "原因", "详情/路径"])
                check_path = self._make_checklist_path()
                try:
                    # Excel 2007 兼容 .xlsx
                    df_check.to_excel(check_path, index=False, engine="openpyxl", sheet_name="核查清单")
                    self._log(f"已生成核查清单：{check_path}")
                    try:
                        os.startfile(check_path)  # 弹窗后自动打开
                    except Exception:
                        pass
                except Exception as e:
                    self._warn(f"生成核查清单失败：{check_path} ({e})")

        except Exception as e:
            self._warn(f"异常：{e}")
            messagebox.showerror("异常", str(e))
        finally:
            for b in (self.btn_both, self.btn_copy, self.btn_pdf): b.config(state="normal")

    def _make_checklist_path(self) -> str:
        """核查清单与日志放一起，命名 check_时间.xlsx"""
        base = Path(self.log_path).parent
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return _norm(base / f"check_{ts}.xlsx")

# ================== 入口 ==================
if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("com.jiangxun.judocmerge.v3312")
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    App(root)
    root.mainloop()
