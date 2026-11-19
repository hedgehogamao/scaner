"""Scaner: local receipt review and assisted entry into 易吧."""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

ROOT = Path(__file__).resolve().parent
SETTINGS = ROOT / "settings.json"
DEFAULT = (0.0, .25, .50, .75, 1.0, 0.0, 1.0)
DEMO = (.07, .34, .53, .72, .94, .30, .77)
PACKAGE_PATH = ROOT / ".packages"
if PACKAGE_PATH.is_dir():
    sys.path.insert(0, str(PACKAGE_PATH))
from PIL import Image, ImageOps, ImageTk
from core import FIELDS, LABELS, ReceiptRow, align_boxes, decimal_input, export_review, load_catalog, plain, validate

BG, NAVY, TEAL, INK, MUTED = "#eef3f6", "#132a37", "#007d73", "#152e3b", "#4a6776"
SOURCE_FIELDS = ("编码列", "名称列", "规格列", "单位列", "参考价列")


def good_bounds(values):
    if len(values) != 7 or not all(0 <= v <= 1 for v in values):
        raise ValueError("标定坐标超出图片。")
    if not all(b-a >= .012 for a,b in zip(values[:4], values[1:5])) or values[6]-values[5] < .035:
        raise ValueError("请按从左到右顺序点选，四列和框选区域都需有足够空间。")
    return tuple(values)


class App:
    def __init__(self, root, initial_demo=False):
        self.root = root
        self.photo = None
        self.photo_path = None
        self.bounds = DEFAULT
        self.preset = None
        self.calibration = None
        self.draw_state = None
        self.srcs = []
        self.real_catalog = {}
        self.catalog = {}
        self.counts = []
        self.rows = []
        self.i = None
        self.syncing = False
        self.demo = False
        self.job = None
        self.results = queue.Queue()
        self.status = tk.StringVar(value="先配置商品资料，或点击「打开演示」试用。")
        self.catalog_text = tk.StringVar(value="尚未配置商品表")
        self.detail_text = tk.StringVar(value="请选择明细行。")
        self.total_input = tk.StringVar()
        self.total_text = tk.StringVar()
        self.values = [tk.StringVar() for _ in LABELS]
        self.build()
        self.load_settings()
        if initial_demo:
            self.open_demo()
        self.refresh()
        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def build(self):
        r = self.root
        r.title("Scaner · 手写单据复核")
        r.geometry("1320x830")
        r.minsize(1080, 680)
        r.configure(bg=BG)
        r.option_add("*Font", ("Microsoft YaHei UI", 10))
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=(10, 6), foreground=INK)
        style.configure("Accent.TButton", background=TEAL, foreground="white")
        style.map("Accent.TButton", background=[("disabled", "#bbc8cb"), ("active", "#00655b")])
        style.configure("Treeview", rowheight=29, foreground=INK, background="white", fieldbackground="white")
        style.configure("Treeview.Heading", background="#e9f1f4", foreground=INK, font=("Microsoft YaHei UI", 10, "bold"))
        style.map("Treeview", background=[("selected", "#bde3df")], foreground=[("selected", INK)])
        head = tk.Frame(r, bg=NAVY, padx=21, pady=13)
        head.pack(fill="x")
        tk.Label(head, text="SCANER  /  手写单据复核", font=("Microsoft YaHei UI", 21, "bold"),
                 bg=NAVY, fg="white").pack(anchor="w")
        tk.Label(head, text="本机识别  →  商品匹配  →  人工确认  →  易吧逐条录入",
                 bg=NAVY, fg="#c1d9e0").pack(anchor="w")
        bar = tk.Frame(r, bg=BG, padx=12, pady=8)
        bar.pack(fill="x")
        self.action_buttons = []
        for caption, fn, accent in (("1  商品 Excel", self.sources, False), ("2  打开照片", self.open_photo, False),
                                    ("标定四列", self.begin_calibration, False), ("旋转照片", self.rotate, False),
                                    ("3  本地 OCR", self.run_ocr, True), ("打开演示", self.open_demo, False)):
            button = ttk.Button(bar, text=caption, command=fn, style="Accent.TButton" if accent else "TButton")
            button.pack(side="left", padx=3)
            self.action_buttons.append(button)
        tk.Label(bar, textvariable=self.catalog_text, bg=BG, fg=MUTED).pack(side="right", padx=5)

        split = ttk.PanedWindow(r, orient="horizontal")
        split.pack(fill="both", expand=True, padx=14, pady=(0, 5))
        left = tk.Frame(split, bg=BG)
        right = tk.Frame(split, bg=BG)
        split.add(left, weight=4)
        split.add(right, weight=6)
        tk.Label(left, text="原图与识别范围", bg=BG, fg=INK, font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(2,4))
        self.canvas = tk.Canvas(left, width=420, height=380, bg="#dbe5ea", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self.redraw())
        self.canvas.bind("<Button-1>", self.photo_click)
        self.canvas.bind("<Escape>", self.cancel_calibration)
        ttk.Button(left, text="放大当前行", command=self.zoom_row).pack(anchor="w", pady=6)
        tk.Label(right, text="明细复核", bg=BG, fg=INK, font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(2,4))
        pane = tk.Frame(right, bg=BG)
        pane.pack(fill="both", expand=True)
        cols = ("code", "name", "qty", "price", "amount", "state")
        self.table = ttk.Treeview(pane, columns=cols, show="headings", selectmode="browse", height=6)
        for key, label, width in zip(cols, ("编码", "商品", "数量", "单价", "票面金额", "状态"), (96, 145, 54, 66, 81, 95)):
            self.table.heading(key, text=label)
            self.table.column(key, width=width, minwidth=45, stretch=key in {"name","state"})
        self.table.tag_configure("error", foreground="#b22920")
        self.table.tag_configure("warn", foreground="#915211")
        self.table.tag_configure("good", foreground="#007160")
        self.table.bind("<<TreeviewSelect>>", self.select_row)
        scroll = ttk.Scrollbar(pane, command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.table.pack(side="left", fill="both", expand=True)

        editor = ttk.LabelFrame(right, text="核对并编辑所选行", padding=(9, 5))
        editor.pack(fill="x", pady=(6, 3))
        self.fields = []
        for i, (label, var) in enumerate(zip(LABELS, self.values)):
            ttk.Label(editor, text=label).grid(row=0, column=i, sticky="w", padx=3)
            entry = ttk.Entry(editor, textvariable=var, width=13)
            entry.grid(row=1, column=i, sticky="ew", padx=3, pady=3)
            editor.columnconfigure(i, weight=1)
            var.trace_add("write", lambda *_a, idx=i: self.edit_row(idx))
            self.fields.append(entry)
        self.product = ttk.Combobox(editor, state="disabled")
        self.product.grid(row=2, column=0, columnspan=4, sticky="ew", padx=3, pady=3)
        self.product.bind("<<ComboboxSelected>>", self.choose_product)
        self.detail = ttk.Label(editor, textvariable=self.detail_text, wraplength=570, foreground=MUTED)
        self.detail.grid(row=3, column=0, columnspan=4, sticky="w", padx=3, pady=(3,1))
        right.bind("<Configure>", lambda e: self.detail.configure(wraplength=max(380, e.width-32)))
        line = ttk.Frame(editor)
        line.grid(row=4, column=0, columnspan=4, sticky="ew", pady=3)
        ttk.Button(line, text="新增一行", command=self.add_row).pack(side="left", padx=3)
        ttk.Button(line, text="移除此行", command=self.remove_row).pack(side="left", padx=3)
        self.confirm_button = ttk.Button(line, text="确认此行", command=self.confirm, style="Accent.TButton")
        self.confirm_button.pack(side="right", padx=3)

        footer = ttk.Frame(right)
        footer.pack(fill="x", pady=4)
        ttk.Label(footer, text="票面合计（可选）").pack(side="left", padx=(3,5))
        ttk.Entry(footer, textvariable=self.total_input, width=12).pack(side="left")
        self.total_input.trace_add("write", lambda *_a: self.totals())
        ttk.Label(footer, textvariable=self.total_text).pack(side="right", padx=4)
        out = ttk.Frame(right)
        out.pack(fill="x", pady=(2, 5))
        self.export_button = ttk.Button(out, text="导出复核 Excel", command=self.export)
        self.export_button.pack(side="left", padx=3)
        self.assist_button = ttk.Button(out, text="易吧逐条录入助手", command=self.assistant, style="Accent.TButton")
        self.assist_button.pack(side="right", padx=3)
        tk.Label(r, textvariable=self.status, anchor="w", bg=NAVY, fg="white", padx=16, pady=8).pack(fill="x")

    def load_settings(self):
        if SETTINGS.exists():
            try:
                obj = json.loads(SETTINGS.read_text(encoding="utf-8"))
                if obj.get("layout"):
                    self.preset = good_bounds(obj["layout"])
                self.srcs = obj.get("sources", [])
                self.real_catalog, self.counts = load_catalog(self.srcs)
                self.catalog = self.real_catalog
            except Exception as e:
                self.status.set(f"配置未能完整加载，请重新选择商品 Excel：{e}")
        self.update_catalog_text()

    def save_settings(self):
        temp = SETTINGS.with_suffix(".tmp")
        temp.write_text(json.dumps({"sources":self.srcs, "layout": self.preset}, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temp, SETTINGS)

    def update_catalog_text(self):
        if self.demo:
            self.catalog_text.set("演示模式 · 4 份样例商品表")
        else:
            self.catalog_text.set(f"{len(self.counts)} / 4 份商品表 · {len(self.real_catalog)} 个编号")

    def sources(self):
        if self.job:
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("配置最多四份只读商品 Excel")
        dialog.transient(self.root)
        dialog.grab_set()
        form = ttk.Frame(dialog, padding=13)
        form.pack(fill="both", expand=True)
        ttk.Label(form, text="支持 .xls / .xlsx / .xlsm。映射留空会自动匹配常见表头；也可写 #A、#B 等列号。").pack(anchor="w", pady=(0,6))
        ttk.Label(form, text="工作表留空读第一张；表头行默认 1，没有表头填 0 并指定列号。").pack(anchor="w", pady=(0,6))
        all_fields = []
        for idx in range(4):
            part = ttk.LabelFrame(form, text=f"商品表 {idx+1}", padding=6)
            part.pack(fill="x", pady=3)
            cfg = self.srcs[idx] if idx < len(self.srcs) else {}
            v = {k: tk.StringVar(value=cfg.get(k, "1" if k == "header" else "")) for k in ("path","sheet","header",*FIELDS)}
            all_fields.append(v)
            ttk.Entry(part, textvariable=v["path"], width=73).grid(row=0, column=0, columnspan=6, sticky="ew", padx=2)
            def pick(var=v["path"]):
                file = filedialog.askopenfilename(parent=dialog, title="选择商品 Excel", filetypes=[("Excel", "*.xls *.xlsx *.xlsm")])
                if file:
                    var.set(file)
            ttk.Button(part, text="浏览…", command=pick).grid(row=0, column=6, padx=3)
            for c, (k, label) in enumerate(zip(("sheet","header",*FIELDS), ("工作表","表头行",*SOURCE_FIELDS))):
                ttk.Label(part, text=label).grid(row=1,column=c,sticky="w",padx=2)
                ttk.Entry(part, textvariable=v[k], width=13).grid(row=2,column=c,sticky="ew",padx=2)
        def commit():
            new = [{k: s.get().strip() for k,s in v.items()} for v in all_fields if v["path"].get().strip()]
            try:
                cat, counts = load_catalog(new)
                old_sources = self.srcs
                self.srcs = new
                try:
                    self.save_settings()
                except Exception:
                    self.srcs = old_sources
                    raise
            except Exception as e:
                messagebox.showerror("商品表未载入", str(e), parent=dialog)
                return
            self.real_catalog, self.counts = cat, counts
            if not self.demo:
                self.catalog = cat
                for row in self.rows:
                    row.confirmed = False
                    row.product_index = None
                self.refresh(self.i)
            self.update_catalog_text()
            conflicts = sum(len(items)>1 for items in cat.values())
            self.status.set(f"已只读加载 {len(counts)} 份商品表；{conflicts} 个编码有资料冲突。")
            dialog.destroy()
        ttk.Button(form, text="加载并保存配置", style="Accent.TButton", command=commit).pack(anchor="e", pady=(8, 0))

    def replace_ok(self):
        return not self.rows or messagebox.askokcancel("清空当前复核", "这会清空窗口里的当前明细。已导出的文件不变。", parent=self.root)

    def open_photo(self):
        path = filedialog.askopenfilename(parent=self.root, title="打开单据照片", filetypes=[("单据图片","*.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff")])
        if path and self.replace_ok():
            self.load_photo(Path(path), False)

    def open_demo(self):
        if self.replace_ok():
            self.load_photo(ROOT / "demo" / "receipt.png", True)

    def load_photo(self, path, demo):
        try:
            with Image.open(path) as f:
                im = ImageOps.exif_transpose(f).convert("RGB")
            if min(im.size) < 40:
                raise ValueError("图片太小。")
            im.thumbnail((4200, 4200), Image.Resampling.LANCZOS)
            cat = load_catalog([{"path": p} for p in sorted((ROOT / "demo").glob("catalog-*.xlsx"))])[0] if demo else self.real_catalog
        except Exception as e:
            messagebox.showerror("图片未打开", str(e), parent=self.root)
            return
        self.photo, self.photo_path, self.demo, self.catalog = im, path, demo, cat
        self.bounds = DEMO if demo else (self.preset or DEFAULT)
        self.calibration = None
        self.rows = []
        self.total_input.set("64.75" if demo else "")
        self.update_catalog_text()
        self.refresh()
        self.status.set("已打开合成示例。点击「本地 OCR」试跑；演示数据不会送往易吧。" if demo else
                        "已打开照片。请确认四列范围，必要时先标定；框内只保留明细，不含表头和整单合计。")

    def rotate(self):
        if self.photo and self.replace_ok():
            self.photo = self.photo.transpose(Image.Transpose.ROTATE_270)
            self.bounds, self.calibration, self.rows = DEFAULT, None, []
            self.refresh()
            self.status.set("已顺时针旋转，请重新标定四列。")

    def begin_calibration(self):
        if self.photo:
            self.calibration = []
            self.canvas.focus_set()
            self.status.set("标定 1 / 5：点击明细区左上角。按 Esc 取消。")
            self.redraw()

    def cancel_calibration(self, _event=None):
        self.calibration = None
        self.redraw()
        self.status.set("已取消标定，仍使用原有范围。")

    def redraw(self):
        c = self.canvas
        c.delete("all")
        w, h = max(c.winfo_width(), 5), max(c.winfo_height(), 5)
        if not self.photo:
            c.create_text(w/2, h/2, text="打开照片后在这里核对原图", fill=MUTED)
            return
        pw, ph = self.photo.size
        scale = min((w-14)/pw, (h-14)/ph)
        if scale <= 0:
            return
        nw, nh = max(1, round(pw*scale)), max(1, round(ph*scale))
        x, y = (w-nw)/2, (h-nh)/2
        self.draw_state = (x,y,nw,nh)
        self.tkimg = ImageTk.PhotoImage(self.photo.resize((nw,nh),Image.Resampling.LANCZOS))
        c.create_image(x,y,image=self.tkimg,anchor="nw")
        def pt(nx,ny):
            return (x+nx*nw,y+ny*nh)
        l,a,b,d,r,t,bo = self.bounds
        c.create_rectangle(*pt(l,t),*pt(r,bo),outline="#008d80",width=2)
        for pos in (a,b,d):
            c.create_line(*pt(pos,t),*pt(pos,bo),fill="#008d80",width=2,dash=(5,3))
        if self.i is not None and self.i < len(self.rows) and self.rows[self.i].box:
            p,q = self.rows[self.i].box
            c.create_rectangle(*pt(l,max(0,p-.008)),*pt(r,min(1,q+.008)),outline="#b45c05",width=3)
        for nx,ny in self.calibration or ():
            px,py = pt(nx,ny)
            c.create_oval(px-4,py-4,px+4,py+4,fill="#d56c15",outline="")

    def photo_click(self, event):
        if self.calibration is None or not self.draw_state:
            return
        x,y,w,h = self.draw_state
        nx,ny = (event.x-x)/w, (event.y-y)/h
        if not (0 <= nx <= 1 and 0 <= ny <= 1):
            return
        self.calibration.append((nx,ny))
        prompts = ("点明细区右下角。", "点编码 / 数量之间的分隔线。", "点数量 / 单价之间的分隔线。", "点单价 / 金额之间的分隔线。")
        if len(self.calibration)<5:
            self.status.set(f"标定 {len(self.calibration)+1} / 5：{prompts[len(self.calibration)-1]}")
        else:
            p,q,a,b,d = self.calibration
            self.calibration = None
            try:
                self.bounds = good_bounds((p[0],a[0],b[0],d[0],q[0],p[1],q[1]))
                if not self.demo:
                    self.preset = self.bounds
                    self.save_settings()
                self.status.set("标定完成。点击「本地 OCR」使用新范围；已有明细不会被自动替换。")
            except Exception as e:
                self.status.set(f"标定或保存失败：{e}")
        self.redraw()

    def zoom_row(self):
        if not self.photo or self.i is None or not self.rows[self.i].box:
            return
        top, bottom = self.rows[self.i].box
        w,h = self.photo.size
        im = self.photo.crop((int(self.bounds[0]*w), max(0,int((top-.025)*h)), int(self.bounds[4]*w),min(h,int((bottom+.025)*h))))
        im.thumbnail((1020,500),Image.Resampling.LANCZOS)
        dialog = tk.Toplevel(self.root)
        dialog.title(f"原图特写 · 第 {self.i+1} 行")
        label = ttk.Label(dialog)
        label.image = ImageTk.PhotoImage(im)
        label.configure(image=label.image)
        label.pack(padx=10,pady=10)

    def run_ocr(self):
        if self.job or not self.photo or not self.replace_ok():
            return
        photo, bounds = self.photo, self.bounds
        self.job = threading.Thread(target=self.ocr_job, args=(photo,bounds), daemon=True)
        self.status.set("正在本机识别。照片不会上传，请稍候…")
        for button in self.action_buttons:
            button.state(["disabled"])
        self.job.start()
        self.root.after(100, self.poll_ocr)

    def ocr_job(self, photo, bounds):
        try:
            w,h = photo.size
            x0,x1,y0,y1 = int(bounds[0]*w),round(bounds[4]*w),int(bounds[5]*h),round(bounds[6]*h)
            with tempfile.TemporaryDirectory(dir=ROOT,prefix=".ocr-") as temp:
                src, dst = Path(temp)/"input.png", Path(temp)/"output.json"
                photo.crop((x0,y0,x1,y1)).save(src)
                env = dict(os.environ)
                if PACKAGE_PATH.is_dir():
                    env["PYTHONPATH"] = str(PACKAGE_PATH) + (os.pathsep+env["PYTHONPATH"] if env.get("PYTHONPATH") else "")
                proc = subprocess.run([sys.executable, str(ROOT/"ocr_worker.py"), str(src), str(dst)], env=env,
                                      capture_output=True, encoding="utf-8", errors="replace", timeout=240,
                                      creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                if proc.returncode or not dst.exists():
                    logs = (proc.stderr or proc.stdout).strip().splitlines()
                    raise RuntimeError(f"OCR 进程退出码 {proc.returncode}："+"\n".join(logs[-7:]))
                boxes = json.loads(dst.read_text(encoding="utf-8"))
            for b in boxes:
                l,t,r,d = b["rect"]
                b["rect"] = (l+x0,t+y0,r+x0,d+y0)
            self.results.put((align_boxes(boxes, bounds, photo.size), None))
        except Exception as e:
            self.results.put((None, str(e)))

    def poll_ocr(self):
        try:
            result, error = self.results.get_nowait()
        except queue.Empty:
            self.root.after(110,self.poll_ocr)
            return
        self.job = None
        for button in self.action_buttons:
            button.state(["!disabled"])
        if error:
            self.status.set("OCR 未完成；已保留原有明细。")
            messagebox.showerror("本地 OCR 失败", error, parent=self.root)
        elif result:
            self.rows = result
            self.refresh(0)
            self.status.set(f"识别得到 {len(result)} 行。请逐行对照原图并确认，尤其注意漏行与小数点。")
        else:
            self.status.set("没识别到明细。请检查标定、照片方向和清晰度；原有明细未覆盖。")

    def row_item(self, i):
        row = self.rows[i]
        product, _, errors, warnings = validate(row,self.catalog)
        tag, status = ("error","需修正") if errors else ("good","已确认") if row.confirmed else ("warn","需看原图" if warnings else "待确认")
        return (row.values[0], product.name if product else "", *row.values[1:], status), (tag,)

    def refresh(self, selected=None):
        self.table.delete(*self.table.get_children())
        for i in range(len(self.rows)):
            values, tags = self.row_item(i)
            self.table.insert("", "end", iid=str(i), values=values, tags=tags)
        if selected is None or not self.rows:
            self.i = None
        else:
            self.i = min(max(0,selected),len(self.rows)-1)
            self.table.selection_set(str(self.i))
            self.table.see(str(self.i))
        self.sync_editor()
        self.totals()
        self.redraw()

    def select_row(self, _event=None):
        iid = self.table.selection()
        self.i = int(iid[0]) if iid else None
        self.sync_editor()
        self.redraw()

    def sync_editor(self):
        self.syncing = True
        active = self.i is not None and self.i < len(self.rows)
        row = self.rows[self.i] if active else ReceiptRow()
        for entry,var,v in zip(self.fields,self.values,row.values):
            entry.configure(state="normal" if active else "disabled")
            var.set(v)
        self.syncing = False
        self.details()

    def details(self):
        if self.i is None or not self.rows:
            self.product.configure(state="disabled", values=())
            self.product.set("")
            self.confirm_button.state(["disabled"])
            self.detail_text.set("选择一行核对，或点击「新增一行」手工录入。")
            return
        row = self.rows[self.i]
        candidates = self.catalog.get(row.values[0].strip(),[])
        choices = [f"{p.name or '（未填写名称）'} / {p.spec or '—'} / {p.unit or '—'} / 参考 {p.price or '—'} / {p.source}" for p in candidates]
        self.product.configure(state="readonly" if len(choices)>1 else "disabled", values=choices)
        chosen = 0 if len(choices)==1 else row.product_index
        if choices and chosen is not None and 0 <= chosen < len(choices):
            self.product.current(chosen)
        else:
            self.product.set("请选择商品资料来源" if choices else "未匹配商品资料")
        _, _, errors, warnings = validate(row,self.catalog)
        self.confirm_button.state(["disabled"] if errors or row.confirmed else ["!disabled"])
        prefix = "已确认。" if row.confirmed else "请对照原图后确认。"
        self.detail_text.set("；".join(errors+warnings) if (errors or warnings) else prefix)

    def update_row(self):
        values, tags = self.row_item(self.i)
        self.table.item(str(self.i),values=values,tags=tags)
        self.details()
        self.totals()

    def edit_row(self, idx):
        if self.syncing or self.i is None or self.i >= len(self.rows):
            return
        row = self.rows[self.i]
        row.values[idx] = self.values[idx].get()
        row.confirmed = False
        if idx == 0:
            row.product_index = None
        self.update_row()

    def choose_product(self, _event=None):
        if self.i is not None:
            row = self.rows[self.i]
            row.product_index = self.product.current()
            row.confirmed = False
            self.update_row()

    def add_row(self):
        if not self.job:
            self.rows.append(ReceiptRow())
            self.refresh(len(self.rows)-1)
            self.fields[0].focus_set()

    def remove_row(self):
        if self.i is not None and not self.job and messagebox.askokcancel("移除复核行","只移除当前窗口这行，不修改商品表或易吧。",parent=self.root):
            index = self.i
            del self.rows[index]
            self.refresh(index)

    def confirm(self):
        if self.i is None:
            return
        row = self.rows[self.i]
        _, _, issues, _ = validate(row,self.catalog)
        if not issues:
            row.confirmed = True
            self.update_row()
            next_row = self.i+1
            if next_row < len(self.rows):
                self.table.selection_set(str(next_row))
                self.table.see(str(next_row))

    def totals(self):
        amounts = [decimal_input(row.values[3]) for row in self.rows]
        summed = sum((n for n in amounts if n is not None),Decimal("0"))
        extra = ""
        total_ok = True
        if any(n is None for n in amounts):
            extra = "（含空缺）"
        typed = self.total_input.get().strip()
        if typed:
            entered = decimal_input(typed)
            total_ok = entered is not None and entered == summed and not extra
            extra += " · 合计一致" if total_ok else " · 合计待核对"
        confirmed = sum(row.confirmed for row in self.rows)
        ready = bool(self.rows) and confirmed == len(self.rows) and total_ok
        self.total_text.set(f"确认 {confirmed}/{len(self.rows)} · 明细 {plain(summed)}{extra}")
        self.export_button.state(["!disabled"] if ready else ["disabled"])
        self.assist_button.state(["!disabled"] if ready and not self.demo else ["disabled"])

    def export(self):
        if self.export_button.instate(["disabled"]):
            return
        path = filedialog.asksaveasfilename(parent=self.root,title="导出已确认明细（不是易吧导入模板）",
            defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")],initialfile="演示复核.xlsx" if self.demo else "已确认单据.xlsx")
        if path:
            try:
                export_review(path,self.rows,self.catalog,[s["path"] for s in self.srcs])
                self.status.set(f"已导出：{Path(path)}（仅供复核，不是易吧导入模板）")
            except Exception as e:
                messagebox.showerror("未导出",str(e),parent=self.root)

    def assistant(self):
        if self.assist_button.instate(["disabled"]):
            return
        rows = []
        for row in self.rows:
            product, values, errors, _ = validate(row,self.catalog)
            if errors or not row.confirmed:
                return
            rows.append((product,values))
        dialog = tk.Toplevel(self.root)
        dialog.title("易吧逐条录入 · 不自动保存")
        dialog.geometry("590x350+25+140")
        dialog.transient(self.root)
        form = ttk.Frame(dialog,padding=14)
        form.pack(fill="both",expand=True)
        ttk.Label(form,text="在易吧打开空白销售单，再按下列步骤逐条录入。",wraplength=530).pack(anchor="w")
        ttk.Label(form,text="① 粘贴编码，等易吧筛查刷新；核对完整编码和品名后再回车。\n② 粘贴数量和票面售价，不沿用默认值。核对该行金额后再点易吧「新行」。",wraplength=530).pack(anchor="w",pady=6)
        summary, feedback = tk.StringVar(), tk.StringVar()
        ttk.Label(form,textvariable=summary,foreground=INK,wraplength=530,font=("Microsoft YaHei UI",13,"bold")).pack(anchor="w",pady=9)
        buttons = ttk.Frame(form)
        buttons.pack(fill="x",pady=5)
        cursor = 0
        def display():
            p,(q,s,a,_) = rows[cursor]
            summary.set(f"{cursor+1} / {len(rows)}   {p.code} · {p.name}\n数量 {plain(q)}｜售价 {plain(s)}｜应核对金额 {plain(a)}")
            previous.state(["!disabled"] if cursor>0 else ["disabled"])
            forward.state(["!disabled"] if cursor+1<len(rows) else ["disabled"])
            feedback.set("复制后切换到易吧，粘贴到对应的单个输入框。")
        def shift(delta):
            nonlocal cursor
            cursor = max(0,min(len(rows)-1,cursor+delta))
            display()
        def copy(field):
            p,(q,s,_,_) = rows[cursor]
            v = (p.code,plain(q),plain(s))[field]
            self.root.clipboard_clear()
            self.root.clipboard_append(v)
            feedback.set(f"已复制{('编码','数量','售价')[field]}：{v}。不会自动点击易吧。")
        for i,label in enumerate(("复制编码","复制数量","复制售价")):
            ttk.Button(buttons,text=label,command=lambda n=i:copy(n)).pack(side="left",padx=3)
        ttk.Label(form,textvariable=feedback,wraplength=530).pack(anchor="w",pady=6)
        nav = ttk.Frame(form)
        nav.pack(fill="x",pady=5)
        previous = ttk.Button(nav,text="上一条",command=lambda:shift(-1))
        previous.pack(side="left",padx=3)
        forward = ttk.Button(nav,text="下一条",command=lambda:shift(1))
        forward.pack(side="left",padx=3)
        ttk.Button(nav,text="结束录入助手",command=dialog.destroy).pack(side="right")
        display()

    def close(self):
        if self.job:
            messagebox.showinfo("正在识别","请等本次 OCR 结束后再关闭。",parent=self.root)
        elif self.replace_ok():
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    App(root, initial_demo="--demo" in sys.argv[1:])
    root.mainloop()
