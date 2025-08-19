# -*- coding: utf-8 -*-
r"""
南京匠勋科技有限公司 — 公安系统自动改名工具 v8.1.5
更新要点：
  · 基本设置区：采用 Entry 可拉伸 + 弹性占位列 + 右侧按钮，消除右侧大空白
  · 运行日志：白底深色字 + 细边框，视觉更协调
其余沿用 v8.1.4：
  · D列(页数) ↔ J列(正文范围) 一致性=错误
  · K列(备考表的图像位置) = 文件夹图像总数=错误
  · 预检无问题不导出；有问题导出 xlsx + csv(GBK) 到 D:\公安改名工具\reports\
  · 撤销日志 txt 用 utf-8-sig
  · 任务栏图标 ico（优先）+ 界面 LOGO png（兜底），高分屏 DPI 感知
"""

import os, sys, json, shutil, time, ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# 可选：natsort、Pillow
try:
    from natsort import natsorted
    NAT_OK = True
except Exception:
    NAT_OK = False

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except Exception:
    PIL_OK = False

def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

LOGO_PNG = resource_path("logo.png")
LOGO_ICO = resource_path("logo.ico")

# 固定目录
import os, sys
from pathlib import Path
# === MacFix: 跨平台路径 & 目录创建开始 ===
if sys.platform.startswith("win"):
    BASE_DIR = r"D:\公安改名工具"
else:
    BASE_DIR = os.path.join(os.path.expanduser("~"), "公安改名工具")
REPORT_DIR = os.path.join(BASE_DIR, "reports")
LOG_DIR    = os.path.join(BASE_DIR, "logs")
UNDO_DIR   = os.path.join(BASE_DIR, "undo_logs")
for d in (BASE_DIR, REPORT_DIR, LOG_DIR, UNDO_DIR):
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
# === MacFix: 跨平台路径 & 目录创建结束 ===


# 业务常量
ALLOWED_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
REQUIRED_COLS = ["档号","封面图像位置","目录的图像位置","结论文书的页码范围","正文的图像范围","备考表的图像位置"]
OPTIONAL_COLS = ["页数"]

# 配色（与 v8.1.4 一致，日志为白底）
COLOR_PRIMARY   = "#00B4A0"
COLOR_PRIMARY_2 = "#009784"
COLOR_ACCENTBAR = "#15C1AE"
COLOR_TEXT      = "#202124"
COLOR_MUTED     = "#5F6368"
COLOR_BG        = "#FFFFFF"
COLOR_PANEL     = "#FFFFFF"
COLOR_BORDER    = "#E5E7EB"
COLOR_LOG_BG    = "#FFFFFF"   # 日志白底
COLOR_LOG_FG    = "#202124"   # 日志深色字

def now_ts(): return time.strftime("%Y%m%d_%H%M%S")
def ensure_dir(p): os.makedirs(p, exist_ok=True)
def natural_sort(lst): return natsorted(lst) if NAT_OK else sorted(lst)

def list_pages(folder):
    files=[]
    for n in os.listdir(folder):
        p=os.path.join(folder,n)
        if os.path.isfile(p) and os.path.splitext(n)[1].lower() in ALLOWED_EXTS:
            files.append(n)
    return natural_sort(files)

def parse_single(v):
    if pd.isna(v): return None
    s=str(v).strip()
    if not s: return None
    try: return int(float(s))
    except: return None

def parse_range(v):
    if pd.isna(v): return None
    s=str(v).replace(" ","")
    if not s: return None
    if "-" in s:
        a,b=s.split("-",1)
        try:
            a=int(float(a)); b=int(float(b))
            if a>b: a,b=b,a
            return (a,b)
        except: return None
    try:
        x=int(float(s)); return (x,x)
    except: return None

def collect_pages(rr):
    if rr is None: return []
    if isinstance(rr, tuple):
        a,b=rr; return list(range(a,b+1))
    return [rr]

def count_range_pages(rr):
    if rr is None: return None
    if isinstance(rr, tuple):
        a,b=rr; return abs(b-a)+1
    try: return 1 if int(rr) else None
    except: return None

def build_name(base, part, ext): return f"{base}N-{part}{ext}"

def suggest(k):
    m={
        "no_folder":"请在图像根目录下创建与“档号”同名的子文件夹，或修正 Excel。",
        "no_images":"将该档号对应的图片拷入同名子文件夹后再预检。",
        "bkb_mismatch":"核对该档号图片总量；K列应等于实际张数。",
        "oor":"检查页码是否超过 1..总页数。",
        "dup":"避免同一页多处引用，调整目录/正文区间。",
        "conflict":"同名目标将被覆盖，请检查页码是否重复/顺序有误。",
        "law_oor":"结论文书范围包含越界页。",
        "body_missing":"补充正文范围（J列）后再核对 D列。",
        "body_count_mismatch":"核对 D列与 J列的数量是否一致（区间是否含端点）。",
    }
    return m.get(k,"核对 Excel 页码与实际图像后修正，再预检。")

# ---------------- 主程序 ----------------
class App(tk.Tk):
    def __init__(self):
        # DPI 感知
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

        super().__init__()

        # 任务栏图标：优先 ico，其次 png 兜底
        try:
            if os.path.isfile(LOGO_ICO):
                self.iconbitmap(LOGO_ICO)
        except Exception:
            pass
        if PIL_OK and os.path.isfile(LOGO_PNG):
            try:
                self._icon_photo = ImageTk.PhotoImage(Image.open(LOGO_PNG))
                self.iconphoto(True, self._icon_photo)
            except Exception:
                pass

        self.title("南京匠勋科技有限公司 — 公安系统自动改名工具 v8.1.5")
        self.configure(bg=COLOR_BG)
        self.geometry("1080x680")
        self.minsize(980, 620)
        self.grid_columnconfigure(0, weight=1)  # 主列可拉伸

        # 主题
        self._init_style()

        # 变量
        self.img_root = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_root = tk.StringVar(value="")
        self.direct_rename = tk.BooleanVar(value=False)

        # 菜单
        self._build_menubar()

        # 主体
        self._build_layout()

    # ---- 主题样式 ----
    def _init_style(self):
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except: pass

        style.configure("TLabel", background=COLOR_PANEL, foreground=COLOR_TEXT,
                        font=("Microsoft YaHei UI", 10))
        style.configure("TButton", font=("Microsoft YaHei UI", 10))
        style.configure("Primary.TButton", background=COLOR_PRIMARY, foreground="#FFFFFF",
                        bordercolor=COLOR_PRIMARY, focusthickness=3, focuscolor=COLOR_PRIMARY, padding=(14,8))
        style.map("Primary.TButton",
                  background=[("active", COLOR_PRIMARY_2), ("pressed", COLOR_PRIMARY_2)])
        style.configure("Outline.TButton", background="#FFFFFF", foreground=COLOR_PRIMARY,
                        bordercolor=COLOR_PRIMARY, relief="solid", padding=(10,6))

    def _build_menubar(self):
        menubar = tk.Menu(self)
        m = tk.Menu(menubar, tearoff=0)
        m.add_command(label="预运行检测（不改名）", command=self.precheck)
        m.add_command(label="执行重命名", command=self.execute)
        m.add_separator()
        m.add_command(label="撤销上一次执行…", command=self.undo_last)
        menubar.add_cascade(label="操作", menu=m)
        self.config(menu=menubar)

    # ---- 分区容器（自绘标题条） ----
    def _section(self, parent, title):
        outer = tk.Frame(parent, bg=COLOR_PANEL, bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER)
        outer.grid_columnconfigure(0, weight=1)
        # 标题条
        bar = tk.Frame(outer, bg=COLOR_ACCENTBAR, height=36)
        bar.grid(row=0, column=0, sticky="ew")
        tk.Label(bar, text=title, bg=COLOR_ACCENTBAR, fg="#FFFFFF",
                 font=("Microsoft YaHei UI", 10, "bold")).pack(side="left", padx=12)
        # 内容
        body = tk.Frame(outer, bg=COLOR_PANEL, bd=0)
        body.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        outer.grid_rowconfigure(1, weight=1)
        return outer, body

    def _build_layout(self):
        # 顶部：Logo + 标题
        top = tk.Frame(self, bg=COLOR_PANEL, bd=0, highlightthickness=1, highlightbackground=COLOR_BORDER)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 8))
        top.grid_columnconfigure(1, weight=1)

        logo_box = tk.Frame(top, bg=COLOR_PANEL)
        logo_box.grid(row=0, column=0, sticky="w", padx=12, pady=12)
        self._load_logo(logo_box)

        tk.Label(top, text="南京匠勋科技有限公司  |  公安系统自动改名工具 v8.1.5",
                 bg=COLOR_PANEL, fg=COLOR_TEXT,
                 font=("Microsoft YaHei UI", 16, "bold")).grid(row=0, column=1, sticky="w", padx=8)

        # 基本设置
        base_sec, base = self._section(self, "基本设置")
        base_sec.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        self.grid_rowconfigure(1, weight=1)

        # 列配置：Entry 列可拉伸 + 第2列弹性占位 + 第3列按钮列
        base.grid_columnconfigure(1, weight=1)  # Entry
        base.grid_columnconfigure(2, weight=1)  # 弹性占位（新增）
        base.grid_columnconfigure(3, weight=0)  # 右侧按钮列

        def add_row(r, label, var, btn_text, btn_cmd):
            ttk.Label(base, text=label).grid(row=r, column=0, sticky="e", padx=8, pady=10)

            e = ttk.Entry(base, textvariable=var)
            e.grid(row=r, column=1, sticky="ew", padx=8, pady=10)

            spacer = tk.Frame(base, bg=COLOR_PANEL)
            spacer.grid(row=r, column=2, sticky="ew")  # 把按钮推到最右

            ttk.Button(base, text=btn_text, style="Outline.TButton", command=btn_cmd)\
                .grid(row=r, column=3, sticky="e", padx=8, pady=10)

        add_row(0, "图像根目录：", self.img_root, "选择…", self.pick_img_root)
        add_row(1, "读取 Excel：", self.excel_path, "选择…", self.pick_excel)
        add_row(2, "输出目录（复制模式）：", self.output_root, "浏览…", self.pick_output_root)

        ttk.Checkbutton(base, text="直接在原目录内重命名（可撤销）",
                        variable=self.direct_rename).grid(row=3, column=1, sticky="w", padx=8, pady=6)

        # 运行日志
        log_sec, log_body = self._section(self, "运行日志（只读）")
        log_sec.grid(row=2, column=0, sticky="nsew", padx=10, pady=(8, 6))
        self.grid_rowconfigure(2, weight=1)

        self.log = tk.Text(
            log_body, height=8,
            bg=COLOR_LOG_BG, fg=COLOR_LOG_FG, insertbackground=COLOR_LOG_FG,
            relief="solid", bd=1, padx=10, pady=8, highlightthickness=0
        )
        self.log.pack(fill="both", expand=True)

        # 底部按钮
        btn_row = tk.Frame(self, bg=COLOR_BG)
        btn_row.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))
        btn_row.grid_columnconfigure(0, weight=0)
        btn_row.grid_columnconfigure(1, weight=0)
        btn_row.grid_columnconfigure(2, weight=1)  # 右侧留白弹性

        ttk.Button(btn_row, text="预运行检测（不改名）", style="Primary.TButton",
                   command=self.precheck).grid(row=0, column=0, padx=(0,10), ipadx=12, ipady=2)
        ttk.Button(btn_row, text="执行重命名", style="Primary.TButton",
                   command=self.execute).grid(row=0, column=1, padx=(0,10), ipadx=12, ipady=2)

    def _load_logo(self, box: tk.Frame):
        if not PIL_OK:
            tk.Label(box, text="(Pillow 未安装)", bg=COLOR_PANEL, fg=COLOR_MUTED).pack()
            return
        try:
            if os.path.isfile(LOGO_PNG):
                img = Image.open(LOGO_PNG)
                target_h = 80
                w, h = img.size
                ratio = target_h / float(h) if h else 1.0
                new_w = min(int(w * ratio), 240)
                img = img.resize((new_w, target_h), Image.LANCZOS)
                self._logo_imgtk = ImageTk.PhotoImage(img)
                tk.Label(box, image=self._logo_imgtk, bg=COLOR_PANEL, bd=0).pack()
            else:
                tk.Label(box, text="(未找到 logo.png)", bg=COLOR_PANEL, fg=COLOR_MUTED).pack()
        except Exception as e:
            tk.Label(box, text=f"(LOGO加载失败：{e})", bg=COLOR_PANEL, fg=COLOR_MUTED).pack()

    # ===== 基础交互 =====
    def pick_img_root(self):
        p = filedialog.askdirectory(title="选择图像根目录")
        if p: self.img_root.set(p)

    def pick_excel(self):
        p = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if p: self.excel_path.set(p)

    def pick_output_root(self):
        p = filedialog.askdirectory(title="选择输出目录（复制模式）")
        if p: self.output_root.set(p)

    def log_write(self, s: str):
        self.log.insert("end", s + "\n"); self.log.see("end")

    # ===== 预检 / 执行 / 撤销（与 v8.1.4 相同业务） =====
    def precheck(self):
        try: self._run(dry_run=True)
        except Exception as e: messagebox.showerror("错误", f"预运行检测失败：{e}")

    def execute(self):
        try: self._run(dry_run=False)
        except Exception as e: messagebox.showerror("错误", f"执行失败：{e}")

    def undo_last(self):
        if not os.path.isdir(UNDO_DIR):
            messagebox.showinfo("提示","未找到撤销日志目录。"); return
        logs=[os.path.join(UNDO_DIR,f) for f in os.listdir(UNDO_DIR) if f.startswith("执行_") and f.endswith(".json")]
        if not logs:
            messagebox.showinfo("提示","没有可撤销的执行日志。"); return
        if not messagebox.askyesno("确认撤销","确定要撤销上一次执行吗？"):
            return
        logs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        log_path=logs[0]
        try:
            with open(log_path,"r",encoding="utf-8") as f: rec=json.load(f)
        except Exception as e:
            messagebox.showerror("错误",f"读取撤销日志失败：{e}"); return
        ops=rec.get("operations",[]); mode=rec.get("mode")
        self.log_write(f"开始撤销：{log_path}（模式：{mode}，操作数：{len(ops)}）")
        ok=fail=0
        for op in reversed(ops):
            act=op.get("action"); src=op.get("src"); dst=op.get("dst")
            try:
                if act=="copy":
                    if dst and os.path.exists(dst): os.remove(dst)
                elif act=="move":
                    if dst and os.path.exists(dst):
                        os.makedirs(os.path.dirname(src), exist_ok=True)
                        if os.path.exists(src):
                            base,ext=os.path.splitext(src); i=1
                            new_src=f"{base}_undo{i}{ext}"
                            while os.path.exists(new_src):
                                i+=1; new_src=f"{base}_undo{i}{ext}"
                            src=new_src
                        os.replace(dst, src)
                ok+=1
            except Exception as e:
                self.log_write(f"[撤销失败] {src} <-> {dst}：{e}"); fail+=1
        try:
            ensure_dir(LOG_DIR)
            rp=os.path.join(LOG_DIR,f"撤销结果_{now_ts()}.txt")
            with open(rp,"w",encoding="utf-8-sig") as f:
                f.write(f"撤销日志：{log_path}\n模式：{mode}\n成功：{ok}\n失败：{fail}\n")
        except Exception:
            rp="(写入失败)"
        self.log_write(f"撤销完成。成功：{ok}，失败：{fail}。")
        messagebox.showinfo("撤销结果",f"撤销完成。\n成功：{ok}\n失败：{fail}\n明细日志：{rp}")

    def _run(self, dry_run: bool):
        img_root=self.img_root.get().strip()
        excel=self.excel_path.get().strip()
        out_root=self.output_root.get().strip()
        direct=self.direct_rename.get()

        if not img_root or not os.path.isdir(img_root): raise RuntimeError("请正确选择图像根目录。")
        if not excel or not os.path.isfile(excel): raise RuntimeError("请正确选择 Excel 文件。")
        if not dry_run and (not direct) and (not out_root): raise RuntimeError("复制模式需要选择“输出目录”。")
        if not dry_run and (not direct) and out_root: ensure_dir(out_root)

        try: df=pd.read_excel(excel)
        except Exception as e: raise RuntimeError(f"无法读取 Excel：{e}")
        missing=[c for c in REQUIRED_COLS if c not in df.columns]
        if missing: raise RuntimeError(f"Excel 缺少必要列：{', '.join(missing)}")

        self.log_write("开始预检……" if dry_run else "开始执行……")
        issues=[]; ok=fail=0
        ops_log={"time":now_ts(),"mode":"move" if direct else "copy","operations":[]}

        for _,row in df.iterrows():
            case_id=str(row["档号"]).strip() if not pd.isna(row["档号"]) else ""
            if not case_id:
                issues.append({"档号":"(空)","等级":"错误","问题":"档号缺失","建议":"补全Excel中的档号"}); fail+=1; continue
            folder=os.path.join(img_root,case_id)
            if not os.path.isdir(folder):
                issues.append({"档号":case_id,"等级":"错误","问题":"未找到档号子文件夹","建议":suggest("no_folder")}); fail+=1; continue
            files=list_pages(folder)
            if not files:
                issues.append({"档号":case_id,"等级":"错误","问题":"文件夹内无图像","建议":suggest("no_images")}); fail+=1; continue
            total=len(files)
            def file_by(ix): return os.path.join(folder, files[ix-1]) if 1<=ix<=total else None

            p_cover=parse_single(row["封面图像位置"])
            p_catalog=parse_range(row["目录的图像位置"])
            p_body=parse_range(row["正文的图像范围"])
            p_bkb=parse_single(row["备考表的图像位置"])
            law_raw=str(row["结论文书的页码范围"]).strip() if not pd.isna(row["结论文书的页码范围"]) else ""

            excel_total=None
            if "页数" in df.columns and not pd.isna(row["页数"]):
                try: excel_total=int(float(row["页数"]))
                except: excel_total=None
            body_cnt=count_range_pages(p_body)
            if excel_total is not None:
                if body_cnt is None:
                    issues.append({"档号":case_id,"等级":"错误","问题":"正文范围缺失或无效，无法与页数核对","建议":suggest("body_missing")})
                elif excel_total!=body_cnt:
                    issues.append({"档号":case_id,"等级":"错误","问题":f"页数与正文范围不一致：页数={excel_total}，正文范围页数={body_cnt}","建议":suggest("body_count_mismatch")})

            if p_bkb is not None and p_bkb!=total:
                issues.append({"档号":case_id,"等级":"错误","问题":f"备考表页号不等于图像总数：K={p_bkb}，实际={total}","建议":suggest("bkb_mismatch")})
                if not dry_run: fail+=1; continue

            used=set(); oor=[]; dup=[]
            def add(pages, tag):
                for p in pages:
                    if p<1 or p>total: oor.append((p,tag)); continue
                    if p in used: dup.append((p,tag))
                    else: used.add(p)

            cover_pages=collect_pages(p_cover); add(cover_pages,"封面")
            catalog_pages=[]
            if p_catalog:
                a,b=p_catalog; catalog_pages=list(range(a,b+1)); add(catalog_pages,"目录")
            body_pages=[]
            if p_body:
                a,b=p_body; body_pages=list(range(a,b+1)); add(body_pages,"正文")
            bkb_pages=collect_pages(p_bkb); add(bkb_pages,"备考表")

            if law_raw:
                parts=[x for x in law_raw.replace("；",";").split(";") if x.strip()]
                for seg in parts:
                    rr=parse_range(seg)
                    if rr:
                        a,b=rr
                        for p in range(a,b+1):
                            if p<1 or p>total:
                                issues.append({"档号":case_id,"等级":"警告","问题":"结论文书越界","建议":suggest("law_oor")})
                                break

            if oor:
                txt="；".join([f"{p}({t})" for p,t in oor])
                issues.append({"档号":case_id,"等级":"错误","问题":f"页号越界：{txt}","建议":suggest("oor")})
            if dup:
                txt="；".join([f"{p}({t})" for p,t in dup])
                issues.append({"档号":case_id,"等级":"错误","问题":f"页号重复引用：{txt}","建议":suggest("dup")})

            if dry_run: ok+=1; continue

            if any(x for x in issues if x["档号"]==case_id and x["等级"]=="错误"):
                self.log_write(f"[{case_id}] 存在错误问题，已跳过执行。"); fail+=1; continue

            def plan(ix, part):
                if not (1<=ix<=total): return None
                src=file_by(ix); ext=os.path.splitext(src)[1].lower()
                dst_dir = folder if self.direct_rename.get() else os.path.join(self.output_root.get().strip(), case_id)
                ensure_dir(dst_dir)
                return src, os.path.join(dst_dir, build_name(case_id, part, ext))

            ops=[]
            for p in cover_pages:
                r=plan(p,"F"); 
                if r: ops.append(r)
            if catalog_pages:
                m=1
                for p in catalog_pages:
                    r=plan(p,f"M{m}"); 
                    if r: ops.append(r); m+=1
            if body_pages:
                seq=1
                for p in body_pages:
                    r=plan(p,f"{seq:03d}"); 
                    if r: ops.append(r); seq+=1
            for p in bkb_pages:
                r=plan(p,"B"); 
                if r: ops.append(r)

            try:
                if not self.direct_rename.get():
                    for src,dst in ops:
                        shutil.copy2(src,dst)
                        ops_log["operations"].append({"action":"copy","src":src,"dst":dst})
                    ok+=1; self.log_write(f"[{case_id}] 输出 {len(ops)} 个文件。")
                else:
                    tmp_pairs=[]
                    for src,dst in ops:
                        base=os.path.dirname(src); tmp=os.path.join(base,f"__tmp__{os.path.basename(src)}")
                        i=1; t0,ext=os.path.splitext(tmp)
                        while os.path.exists(tmp):
                            tmp=f"{t0}_{i}{ext}"; i+=1
                        if os.path.abspath(src)!=os.path.abspath(dst): os.replace(src,tmp)
                        tmp_pairs.append((tmp,dst,src))
                    for tmp,dst,src0 in tmp_pairs:
                        if os.path.exists(dst): os.remove(dst)
                        os.replace(tmp,dst)
                        ops_log["operations"].append({"action":"move","src":src0,"dst":dst})
                    ok+=1; self.log_write(f"[{case_id}] 原地重命名完成：{len(ops)} 个文件。")
            except Exception as e:
                fail+=1; self.log_write(f"[{case_id}] 处理失败：{e}")

        if dry_run:
            err=sum(1 for x in issues if x["等级"]=="错误")
            warn=sum(1 for x in issues if x["等级"]=="警告")
            self.log_write(f"预检结束。错误 {err}，警告 {warn}。")
            if err==0 and warn==0:
                messagebox.showinfo("预检完成","未发现问题。")
            else:
                try:
                    ensure_dir(REPORT_DIR)
                    xlsx=os.path.join(REPORT_DIR,f"预运行检测报告_{now_ts()}.xlsx")
                    pd.DataFrame(issues).to_excel(xlsx, index=False, engine="openpyxl")
                    csv =os.path.join(REPORT_DIR,f"预运行检测报告_{now_ts()}.csv")
                    pd.DataFrame(issues).to_csv(csv, index=False, encoding="gbk")
                    messagebox.showinfo("预检完成", f"错误：{err}，警告：{warn}\n已导出报告：\n{xlsx}\n{csv}")
                except Exception as e:
                    messagebox.showwarning("预检完成（报告写入失败）", f"错误：{err}，警告：{warn}\n报告写入失败：{e}")
        else:
            try:
                ensure_dir(UNDO_DIR)
                up=os.path.join(UNDO_DIR,f"执行_{now_ts()}.json")
                with open(up,"w",encoding="utf-8") as f:
                    json.dump(ops_log,f,ensure_ascii=False,indent=2)
                self.log_write(f"已保存事务日志（用于撤销）：{up}")
            except Exception as e:
                self.log_write(f"保存事务日志失败：{e}")
            messagebox.showinfo("完成", f"处理结束。\n成功：{ok}\n失败：{fail}")

# 入口
if __name__ == "__main__":
    App().mainloop()
