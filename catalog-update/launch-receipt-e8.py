"""Scaner Desktop catalogs and standard printed receipt recognition."""
from pathlib import Path
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from itertools import product
import json
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import scaner
from core import (CENT, ReceiptRow, canonical_header, catalog_sheet, cell_text, decimal_input,
                  excel_code, export_review, load_catalog, validate)

scaner.SETTINGS = HERE / "settings.json"


def signature(sources):
    """Only stat the configured files; never copy or write a source workbook."""
    files = []
    for entry in sources:
        try:
            st = Path(entry["path"]).stat()
            files.append((st.st_dev, st.st_ino, st.st_size, st.st_mtime_ns, st.st_ctime_ns))
        except (OSError, KeyError):
            files.append(None)
    return json.dumps(sources, sort_keys=True, ensure_ascii=False), tuple(files)


class CatalogApp(scaner.App):
    POLL_MS = 1000
    RETRY_SECONDS = 5

    def __init__(self, root):
        self.catalog_waiting = False
        self.catalog_error = None
        self._applied = None
        self._observed = None
        self._seen_at = 0.
        self._retry_at = 0.
        self._loading = False
        self._result = queue.Queue()
        self._assistants = set()
        super().__init__(root)
        self.root.after(self.POLL_MS, self._watch)

    def load_settings(self):
        try:
            data = json.loads(scaner.SETTINGS.read_text(encoding="utf-8"))
            before = signature(data.get("sources", []))
        except (OSError, ValueError):
            before = None
        super().load_settings()
        now = signature(self.srcs)
        if before == now and len(self.counts) == len(self.srcs):
            self._applied = now
        elif self.srcs:
            self.catalog_waiting = True
        self.update_catalog_text()

    def update_catalog_text(self):
        super().update_catalog_text()
        if self.srcs and not self.demo:
            label = " · 正在刷新" if self.catalog_waiting else " · 自动刷新已开启"
            if self.catalog_error:
                label = " · 等待 Excel 保存完成"
            self.catalog_text.set(self.catalog_text.get() + label)

    def totals(self):
        super().totals()
        if self.catalog_waiting and not self.demo:
            self.export_button.state(["disabled"])
            self.assist_button.state(["disabled"])

    def _close_assistants(self):
        count = 0
        for dialog in self._assistants:
            if dialog.winfo_exists():
                dialog.destroy()
                count += 1
        self._assistants.clear()
        return count

    def _notice_change(self, current, now):
        if current != self._observed:
            self._observed, self._seen_at, self._retry_at = current, now, 0.
            self.catalog_error = None
        if current != self._applied and self.srcs and not self.catalog_waiting:
            self.catalog_waiting = True
            closed = self._close_assistants()
            self.update_catalog_text()
            self.totals()
            if not self.demo:
                message = "检测到 Excel 有更新，正等保存完成。当前单据不会清空。"
                if closed:
                    message += " 已关闭旧数据的录入助手。"
                self.status.set(message)

    def _read_in_worker(self, sources, requested):
        try:
            catalog, counts = load_catalog(sources)
            self._result.put((requested, signature(sources), catalog, counts, None))
        except Exception as exc:
            self._result.put((requested, signature(sources), None, None, str(exc)))

    def _accept(self, stamp, catalog, counts):
        previous = self.real_catalog
        invalidated = 0
        self.real_catalog, self.counts = catalog, counts
        self._applied = stamp
        self.catalog_waiting = False
        self.catalog_error = None
        if not self.demo:
            self.catalog = catalog
            for row in self.rows:
                key = row.values[0].strip()
                if previous.get(key) != catalog.get(key):
                    invalidated += int(row.confirmed)
                    row.confirmed = False
                    row.product_index = None
            self.refresh(self.i)
            message = f"商品 Excel 已刷新：{len(counts)} 份表，{len(catalog)} 个编码。"
            if invalidated:
                message += f" {invalidated} 行商品资料已变动，请重新确认。"
            self.status.set(message)
        self.update_catalog_text()

    def _watch(self):
        if self.srcs:
            now = time.monotonic()
            current = signature(self.srcs)
            self._notice_change(current, now)
            try:
                requested, after, catalog, counts, error = self._result.get_nowait()
            except queue.Empty:
                pass
            else:
                self._loading = False
                if requested == after == current:
                    if error is None:
                        self._accept(current, catalog, counts)
                    else:
                        self.catalog_error = error
                        self._retry_at = now + self.RETRY_SECONDS
                        self.update_catalog_text()
                        if not self.demo:
                            self.status.set("暂时读不到最新商品表，正等待或重试。明细保留，导出和录入助手暂停。")
            if current == self._applied and self.catalog_waiting and not self._loading:
                self.catalog_waiting = False
                self.catalog_error = None
                self.update_catalog_text()
                self.totals()
            if (current != self._applied and not self._loading and now >= self._retry_at
                    and now - self._seen_at >= self.POLL_MS / 1000 * .95):
                self._loading = True
                sources = [dict(item) for item in self.srcs]
                threading.Thread(target=self._read_in_worker,args=(sources,current),daemon=True).start()
        else:
            try:
                self._result.get_nowait()
                self._loading = False
            except queue.Empty:
                pass
            if self.catalog_waiting:
                self.catalog_waiting = False
                self.catalog_error = None
                self.update_catalog_text()
                self.totals()
        self.root.after(self.POLL_MS, self._watch)

    def _ready(self):
        if not self.demo and self.srcs:
            current = signature(self.srcs)
            self._notice_change(current, time.monotonic())
            if self.catalog_waiting:
                return False
        return True

    def assistant(self):
        if not self._ready():
            return
        before = set(self.root.winfo_children())
        super().assistant()
        self._assistants.update(w for w in set(self.root.winfo_children()) - before if isinstance(w, tk.Toplevel))

    def export(self):
        if not self._ready() or self.export_button.instate(["disabled"]):
            return
        path = filedialog.asksaveasfilename(parent=self.root,title="导出已确认明细（不是易吧导入模板）",
            defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")],
            initialfile="演示复核.xlsx" if self.demo else "已确认单据.xlsx")
        if not path or not self._ready() or self.export_button.instate(["disabled"]):
            return
        try:
            export_review(path,self.rows,self.catalog,[s["path"] for s in self.srcs])
            self.status.set(f"已导出：{Path(path)}（仅供复核，不是易吧导入模板）")
        except Exception as exc:
            messagebox.showerror("未导出",str(exc),parent=self.root)


ORDER_OPTIONS = ("单价 → 数量（手写样票）", "数量 → 单价（印刷顺序）")
PROPORTIONS = (0., .273, .520, .758, 1.)
GRID_LINES = 12
MODEL = "PP-OCRv5_mobile_rec"


@dataclass
class Line:
    # y = m*x+b (horizontal) or x = m*y+b (vertical)
    m: float
    b: float
    at: float
    rank: int


def _clusters(items, gap):
    groups = []
    for line in sorted(items, key=lambda x: x.at):
        if groups and abs(line.at - groups[-1][0].at) <= gap:
            groups[-1].append(line)
        else:
            groups.append([line])
    return [min(g, key=lambda x:x.rank) for g in groups]


def _regular_run(lines, count, min_span, max_span, ratios=None, min_match=None):
    if ratios is None:
        ratios = [i/(count-1) for i in range(count)]
    best = None
    for i, first in enumerate(lines):
        for last in lines[i+1:]:
            span = last.at - first.at
            if not min_span <= span <= max_span:
                continue
            pitch = span/(count-1)
            targets = [first.at + r*span for r in ratios]
            matches = []
            cost = 0.
            for j, x in enumerate(targets):
                match = first if j == 0 else last if j == count-1 else min(lines,key=lambda l: abs(l.at-x))
                delta = abs(match.at-x)/pitch
                if delta > .23:
                    matches.append(None)
                else:
                    matches.append(match)
                    cost += delta + .25 * min(match.rank,350)/350
            hit = sum(m is not None for m in matches)
            if hit >= (count if min_match is None else min_match):
                merit = 10*hit-cost
                if best is None or merit > best[0]:
                    best = (merit, targets, matches)
    if best is None:
        raise ValueError("未找到标准票据的完整表格线；可用「点选四角」框出整个表格（含表头、合计）。")
    _, targets, matches = best
    for i, line in enumerate(matches):
        if line is None:
            a = max(j for j in range(i) if matches[j] is not None)
            b = min(j for j in range(i+1,count) if matches[j] is not None)
            r = (i-a)/(b-a)
            old, new = matches[a],matches[b]
            matches[i] = Line(old.m+(new.m-old.m)*r,old.b+(new.b-old.b)*r,targets[i],5000)
    return matches


def _intersect(horiz, vert):
    d = 1-horiz.m*vert.m
    x = (vert.m*horiz.b+vert.b)/d
    return (x,horiz.m*x+horiz.b)


def detect_grid(photo):
    import cv2
    import numpy as np
    from PIL import Image
    pixels = np.asarray(photo)
    H,W = pixels.shape[:2]
    work = photo.copy()
    work.thumbnail((1200,1600),Image.Resampling.LANCZOS)
    small = np.asarray(work)
    h,w = small.shape[:2]
    gray = cv2.cvtColor(small,cv2.COLOR_RGB2GRAY)
    edges = cv2.Canny(cv2.GaussianBlur(gray,(3,3),0),35,110)
    margin = max(3,int(min(h,w)*.022))
    edges[:margin]=0; edges[-margin:]=0
    edges[:,:margin]=0; edges[:,-margin:]=0
    raw = cv2.HoughLines(edges,1,np.pi/1800,max(95,int(min(h,w)*.1375)))
    if raw is None:
        raise ValueError("没有检测到足够清晰的表格线；请点选四角或重新拍照。")
    hs, vs = [], []
    for rank, (rho, theta) in enumerate(raw[:1700,0,:2]):
        c,s = float(np.cos(theta)),float(np.sin(theta))
        if abs(s) > .966:
            m,b = -c/s,float(rho)/s
            z = b+m*.5*w
            if .06*h < z < .975*h:
                hs.append(Line(m,b,z,rank))
        if abs(c) > .966:
            m,b = -s/c,float(rho)/c
            z = b+m*.62*h
            if .02*w < z < .98*w:
                vs.append(Line(m,b,z,rank))
    hs = _regular_run(_clusters(hs,max(4,h*.0046)),GRID_LINES,.26*h,.88*h,min_match=11)
    vs = _regular_run(_clusters(vs,max(4,w*.006)),5,.35*w,.95*w,PROPORTIONS)
    def to_image(point):
        return (point[0]*W/w,point[1]*H/h)
    quad = [to_image(_intersect(hs[a],vs[b])) for a,b in ((0,0),(0,-1),(-1,-1),(-1,0))]
    row_points = [to_image(_intersect(line,vs[2])) for line in hs]
    column_points = [to_image(_intersect(hs[5],line)) for line in vs]
    padx = min(H,W)*.025
    if any(not(-padx<=x<=W+padx and -padx<=y<=H+padx) for x,y in quad):
        raise ValueError("检测的表格超出照片边界，请完整拍下表格。")
    return quad,row_points,column_points


def rectify(photo, corners=None):
    import cv2
    import numpy as np
    if corners is None:
        corners,rp,cp = detect_grid(photo)
    else:
        rp = cp = None
    src = np.float32(corners)
    if src.shape != (4,2) or not np.isfinite(src).all() or not cv2.isContourConvex(src):
        raise ValueError("四角必须按左上、右上、右下、左下依次点击。")
    area = cv2.contourArea(src,oriented=True)
    if area < photo.width*photo.height*.045 or src[0,0]>=src[1,0] or src[0,1]>=src[3,1]:
        raise ValueError("框选区域太小或四角顺序不正确。")
    left,top,width,height,right = 285,28,1420,1518,190
    dst = np.float32(((left,top),(left+width,top),(left+width,top+height),(left,top+height)))
    t = cv2.getPerspectiveTransform(src,dst)
    warped = cv2.warpPerspective(np.asarray(photo),t,(left+width+right,top+height+24),
                                 borderMode=cv2.BORDER_CONSTANT,borderValue=(255,255,255))
    def map_coords(points,dim):
        return list(map(float,cv2.perspectiveTransform(np.float32([points]),t)[0,:,dim]))
    ys = map_coords(rp,1) if rp is not None else [top+height*i/(GRID_LINES-1) for i in range(GRID_LINES)]
    xs = map_coords(cp,0) if cp is not None else [left+width*r for r in PROPORTIONS]
    if min(b-a for a,b in zip(ys,ys[1:])) < height*.050 or min(b-a for a,b in zip(xs,xs[1:])) < width*.13:
        raise ValueError("表格线间距异常，请重新标定四角。")
    return warped,xs,ys


def pen_mask(warped,xs,ys):
    import cv2
    import numpy as np
    rgb = warped.astype(np.int16)
    r,g,b = cv2.split(rgb)
    gray = cv2.cvtColor(warped,cv2.COLOR_RGB2GRAY)
    blue = (b-g>=8)&(b-r>=4)&(gray<177)
    sample = blue[int(ys[1]):int(ys[-2]),int(xs[0]):int(xs[-1])]
    use_blue = np.count_nonzero(sample) > sample.size*.0015
    grayscale = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,61,15)
    grayscale[gray>166] = 0
    # Keep faded pixels close to blue strokes; the pure color threshold splits
    # handwritten zeroes on this paper. The alternate grey mask reads light ink.
    halo = cv2.dilate(blue.astype(np.uint8),cv2.getStructuringElement(cv2.MORPH_ELLIPSE,(23,23)))
    colored = np.where(halo,grayscale,0).astype(np.uint8)
    def remove_rules(mask):
        horiz = cv2.morphologyEx(mask,cv2.MORPH_OPEN,cv2.getStructuringElement(cv2.MORPH_RECT,(125,1)))
        vert = cv2.morphologyEx(mask,cv2.MORPH_OPEN,cv2.getStructuringElement(cv2.MORPH_RECT,(1,105)))
        return cv2.bitwise_and(mask,cv2.bitwise_not(cv2.bitwise_or(horiz,vert)))
    return remove_rules(grayscale),remove_rules(colored),use_blue


def ink_variants(mask, split=False):
    import cv2
    import numpy as np
    h,w = mask.shape
    n,labels,stats,_ = cv2.connectedComponentsWithStats(mask,8)
    clean = np.zeros_like(mask)
    for k in range(1,n):
        x,y,a,b,area = stats[k]
        if area >= max(10,h*w*.00017) and (b >= h*.055 or a >= w*.045):
            clean[labels==k]=255
    yy,xx = np.where(clean!=0)
    if len(xx) < 32 or max(yy)-min(yy) < h*.19:
        return []
    lo,hi = int(min(xx)),int(max(xx))+1
    ymin,ymax = int(min(yy)),int(max(yy))+1
    images = []
    def pack(start,finish,mode):
        crop = clean[ymin:ymax,start:finish]
        my,mx = np.where(crop!=0)
        if len(my):
            crop = crop[int(min(my)):int(max(my))+1,:]
        canvas = cv2.copyMakeBorder(255-crop,8,8,9,9,cv2.BORDER_CONSTANT,value=255)
        if canvas.shape[0] < 64:
            canvas = cv2.resize(canvas,None,fx=64/canvas.shape[0],fy=64/canvas.shape[0],interpolation=cv2.INTER_CUBIC)
        return (mode,cv2.cvtColor(canvas,cv2.COLOR_GRAY2BGR))
    images.append(pack(lo,hi,"full"))
    if split:
        columns = np.any(clean[ymin:ymax],axis=0)
        empties = np.where(~columns[lo:hi])[0] + lo
        for gap in reversed(np.split(empties,np.where(np.diff(empties)!=1)[0]+1)):
            if len(gap) >= max(12,int(h*.19)) and hi-gap[-1] > h*.19:
                images.append(pack(int(gap[-1])+1,hi,"right"))
                break
    return images


def code_gutter(mask,left,divider,table_width):
    """Plan a second code crop from separated ink runs on the same row."""
    import cv2
    h,w = mask.shape
    l0 = max(0,round(left-.104*table_width))
    r0 = min(w,round(divider+.041*table_width))
    start = max(0,round(left-.193*table_width))
    stop = min(w,round(divider+.23*table_width))
    pad=max(2,round(h*.025))
    n,_,stats,_ = cv2.connectedComponentsWithStats(mask[pad:h-pad,start:stop],8)
    runs=[]
    wide_touch=False
    for x,y,a,b,area in stats[1:n]:
        if (b<h*.19 or b>h*.96 or area<max(17,h*.21)
                or y+b*.5<h*.18 or y+b*.5>h*.85):
            continue
        if a>table_width*.17:
            wide_touch |= int(x)+start<divider and int(x+a)+start>r0
            continue
        runs.append((int(x)+start,int(x+a)+start))
    merged=[]
    for a,b in sorted(runs):
        if merged and a<=merged[-1][1]:
            merged[-1]=(merged[-1][0],max(b,merged[-1][1]))
        else:
            merged.append((a,b))

    l1,r1=l0,r0
    left_group=next((i for i,(a,b) in enumerate(merged) if b>=l0+5 and a<left+(divider-left)*.36),None)
    if left_group is not None:
        i=left_group
        while i>0 and merged[i][0]-merged[i-1][1]<=max(15,h*.30):
            i-=1
        if merged[i][0]<l0-4:
            l1=max(start,merged[i][0]-6)

    right_cap=min(w-1,round(divider+.106*table_width))
    ambiguous=wide_touch
    anchors=[i for i,(a,_) in enumerate(merged) if a<divider-max(6,h*.065)]
    if anchors:
        i=anchors[-1]
        a,b=merged[i]
        if b>=divider-1 or (i+1<len(merged) and merged[i+1][0]>=divider and merged[i+1][0]-b<h*.14):
            j=i
            while j+1<len(merged) and merged[j+1][0]-merged[j][1]<h*.14:
                j+=1
            end=merged[j][1]
            next_start=merged[j+1][0] if j+1<len(merged) else stop
            if end>r0-3:
                # Extend only to an actual visible gap near the divider.
                # If the price and code are connected we keep the old crop.
                if not ambiguous and end+4<=right_cap and next_start-end>=max(15,h*.20):
                    r1=min(right_cap,end+max(4,round(min(7,(next_start-end)/3))))
                    r1=max(r0,r1)
                else:
                    ambiguous=True
    return {"fixed":[l0,r0],"crop":[l1,r1],"left":l1<l0,"right":r1>r0,
            "right_ambiguous":ambiguous}


def _rec_model():
    cache = scaner.ROOT/".cache"
    profile = cache/"profile"
    profile.mkdir(parents=True,exist_ok=True)
    os.environ.update(PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK="1", HF_HUB_OFFLINE="1",
                      PADDLE_PDX_CACHE_HOME=str(cache/"paddlex"), USERPROFILE=str(profile), HOME=str(profile))
    from paddleocr import TextRecognition
    return TextRecognition(model_name=MODEL,model_dir=str(scaner.ROOT/"models"/MODEL),
                           device="cpu",enable_mkldnn=False,cpu_threads=4)


def numeric(text):
    s = re.sub(r"\s+","",text.strip())
    repaired = s.strip(".,，").translate(str.maketrans({"O":"0","o":"0","I":"1","l":"1","|":"1","/":"1","z":"2","Z":"2","J":"7"}))
    return (repaired, repaired!=s) if decimal_input(repaired) is not None else (None,False)


def apply_total_candidates(rows,total_choices):
    """Propose one unreadable digit field per affected line only when a
    unique reconstruction agrees with a separately recognized printed total.
    No proposal is marked confirmed; raw OCR evidence stays on the record.
    """
    targets = {decimal_input(text):text for _,_,text in reversed(total_choices)}
    if not targets:
        return None
    all_choices = []
    for row in rows:
        q,p,a = map(decimal_input,row["values"][1:])
        if q is not None and p is not None and a is not None and (q*p).quantize(CENT,rounding=ROUND_HALF_UP)==a:
            all_choices.append([(a,None,None)])
            continue
        options = set()
        strong = [(s or 0)>=.46 for s in row["scores"][1:]]
        if q is not None and p is not None and q > 0 and p >= 0 and strong[0] and strong[1]:
            new = (q*p).quantize(CENT,rounding=ROUND_HALF_UP)
            options.add((new,3,new))
        if q is not None and a is not None and q > 0 and a >= 0 and strong[0] and strong[2]:
            new = a/q
            if new == new.quantize(Decimal(".000001")):
                options.add((a,2,new))
        if p is not None and a is not None and p > 0 and a > 0 and strong[1] and strong[2]:
            new = a/p
            if new == new.quantize(Decimal(".000001")):
                options.add((a,1,new))
        all_choices.append([entry for entry in sorted(options) if decimal_input(format(entry[2],"f")) is not None])
    if not all(all_choices) or len(all_choices)>GRID_LINES-3:
        return None
    winners = []
    for choice in product(*all_choices):
        total = sum(line[0] for line in choice)
        if total in targets:
            winners.append((total,choice))
            if len(winners)>1:
                return None
    if len(winners)!=1:
        return None
    total,choice = winners[0]
    for item,(_,field,value) in zip(rows,choice,strict=True):
        if field is None:
            continue
        candidate = format(value,"f")
        old = item["values"][field]
        label = ("数量","单价","金额")[field-1]
        item["values"][field] = candidate
        item["scores"][field] = None
        item["flags"] = [f"{label}原 OCR「{old or '未读清'}」；{candidate} 是用该行另两格和票面合计算出的候选，不是直接读到，请核对",
                         *[f for f in item["flags"] if not f.startswith("按当前数量")]]
    return targets[total]


def recognize_receipt(source, preview, swapped, corners=None):
    import cv2
    from PIL import Image, ImageOps
    with Image.open(source) as photo:
        photo = ImageOps.exif_transpose(photo).convert("RGB")
    photo.thumbnail((4200,4200),Image.Resampling.LANCZOS)
    warped,xs,ys = rectify(photo,corners)
    gray,bw,blue = pen_mask(warped,xs,ys)
    H,W = bw.shape
    plans,images,data = [],[],{}
    extra_plans, extra_images, gutters = [],[],{}
    # Let codes extend outside their printed cell; read numeric fields first
    # away from printed lines. The blue-ink mask can include corrected totals
    # written beyond the outer border without looking at the printed side text.
    offsets = ((-.104,.041),(.011,-.002),(.004,-.005),(.006,-.007))
    # Only the nine detail bands and the amount box in the last printed band.
    for row in range(1,GRID_LINES-1):
        seq = (3,) if row==GRID_LINES-2 else range(4)
        if row<GRID_LINES-2:
            y0,y1 = ys[row],ys[row+1]
            band=(bw if blue else gray)[max(0,round(y0)):min(H,round(y1))]
            info=code_gutter(band,xs[0],xs[1],xs[-1]-xs[0])
            gutters[row]=info
            if info["left"] or info["right"]:
                l,r=info["crop"]
                for kind,mask in (("ink",bw),("gray",gray)) if blue else (("gray",gray),):
                    trim=(y1-y0)*(.017 if kind=="ink" else .055)
                    crop=mask[max(0,round(y0+trim)):min(H,round(y1-trim)),l:r]
                    for _,im in ink_variants(crop):
                        extra_plans.extend((row,row))
                        extra_images.extend((im,cv2.resize(im,None,fx=1.49,fy=1,interpolation=cv2.INTER_CUBIC)))
        for col in seq:
            w = xs[-1]-xs[0]
            y0,y1 = ys[row],ys[row+1]
            l,r = offsets[col]
            choices = [("gray",gray,l,r)]
            if blue:
                choices.insert(0,("ink",bw,l,(.130 if col==3 else r)))
                if col==1:
                    choices.append(("inner",bw,.052,-.002))
            if row==GRID_LINES-2 and not blue:
                choices = [("gray",gray,.012,.07)]
            if row==GRID_LINES-2 and blue:
                choices = [("ink",bw,.007,.130),("gray",gray,.009,-.002)]
            if row==GRID_LINES-2:
                y0 -= (y1-y0)*.20
            data[row,col] = []
            for kind,mask,l,r in choices:
                pad = (y1-y0)*(.017 if kind!="gray" else .055)
                x0,x1 = max(0,round(xs[col]+l*w)),min(W,round(xs[col+1]+r*w))
                ink = mask[max(0,round(y0+pad)):min(H,round(y1-pad)),x0:x1]
                for mode,img in ink_variants(ink,split=(col!=0 and col!=1 and row!=GRID_LINES-2)):
                    plans.append((row,col,mode+":"+kind))
                    images.append(img)
                    plans.append((row,col,mode+":"+kind+"-wide"))
                    images.append(cv2.resize(img,None,fx=1.49,fy=1,interpolation=cv2.INTER_CUBIC))
    if not images:
        raise ValueError("表格已定位，但没有读到墨迹。请用清楚的照片再试。")
    rec = _rec_model()
    for (row,col,mode), result in zip(plans,rec.predict(images,batch_size=8),strict=True):
        res = result.json["res"]
        data[row,col].append((str(res["rec_text"]).strip(), float(res["rec_score"]), mode))
    extra_data={}
    if extra_images:
        # Separate batch: widening a SKU crop cannot change the primary
        # numeric fields' existing OCR batch or their assignment.
        for row,result in zip(extra_plans,rec.predict(extra_images,batch_size=8),strict=True):
            res=result.json["res"]
            extra_data.setdefault(row,[]).append((str(res["rec_text"]).strip(),float(res["rec_score"])))
    answers = []
    for row in range(1,GRID_LINES-2):
        cells = [data.get((row,c),[]) for c in range(4)]
        if sum(any(s>=.34 and any(c.isalnum() for c in t) for t,s,_ in cell) for cell in cells)<2:
            continue
        cells = [sorted(cell,key=lambda item:item[1],reverse=True) for cell in cells]
        values = [items[0][0] if items else "" for items in cells]
        scores = [items[0][1] if items else None for items in cells]
        flags = []
        info=gutters[row]
        if info["left"] or info["right"]:
            sides = ("左侧" if info["left"] else "") + ("右侧" if info["right"] else "")
            flags.append(f"货号{sides}越过原识别范围，已补读；扩展结果仅作候选")
        if info["right_ambiguous"]:
            flags.append("货号和下一列之间未找到可靠空隙，未向右多读；请核对两列")
        op = [1,2,3]
        labels = {1:"单价" if swapped else "数量",2:"数量" if swapped else "单价",3:"金额"}
        nums = []
        for c in op:
            choices = []
            for text,score,mode in cells[c]:
                num,fixed = numeric(text)
                if num is not None:
                    choices.append((num,score,mode,fixed))
            nums.append(choices)
        matched = []
        for a,b,c in product(*nums):
            expected = (decimal_input(a[0])*decimal_input(b[0])).quantize(CENT,rounding=ROUND_HALF_UP)
            if expected == decimal_input(c[0]) and all(x[1]>=.40 for x in (a,b,c)):
                matched.append((a,b,c))
        if matched:
            chosen = max(matched,key=lambda p:sum(x[1]-(.22 if x[2].startswith("right") else 0)-(.10 if x[3] else 0) for x in p))
        else:
            chosen = [next((x for x in opts if x[2].startswith("full")),None) for opts in nums]
            # Distinct writing on the right of a crossed-out digit is useful
            # evidence, but not proof. Show the best separated read and flag
            # the cell rather than quietly joining crossed-out and new digits.
            for n in range(3):
                if chosen[n] is None or chosen[n][1]<.75:
                    alternate = next((x for x in nums[n] if x[2].startswith("right") and x[1]>=.73),None)
                    if alternate is not None:
                        chosen[n] = alternate
        for j,c in enumerate(op):
            val = chosen[j]
            if val is None:
                continue
            values[c],scores[c],mode,fixed=val
            if mode.startswith("right"):
                flags.append(f"{labels[c]}有涂改或多段，这是分段识别候选，须对照原票")
            elif fixed:
                flags.append(f"{labels[c]}有易混字符，须对照原票")
        if swapped:
            values[1],values[2] = values[2],values[1]
            scores[1],scores[2] = scores[2],scores[1]
        q,p,a = map(decimal_input,values[1:])
        if q is not None and p is not None and (a is None or (q*p).quantize(CENT,rounding=ROUND_HALF_UP)!=a):
            flags.append(f"按当前数量×单价算得 {(q*p).quantize(CENT,rounding=ROUND_HALF_UP)}，请核对原票金额")
        answers.append({"row":row,"values":values,"scores":scores,"flags":flags,
                        "extra_codes":extra_data.get(row,[]),"gutter":info,
                        "box":[ys[row]/H,ys[row+1]/H],"raw":[[list(x) for x in c] for c in cells]})
    total_cells=data.get((GRID_LINES-2,3),[])
    total_cells = sorted(total_cells,key=lambda item:item[1],reverse=True)
    total_choices = [(t,s,numeric(t)[0]) for t,s,_ in total_cells if s>=.36 and numeric(t)[0] is not None]
    total = apply_total_candidates(answers,total_choices)
    reliable = next((number for _,score,number in total_choices if score>=.80),None)
    reported = total_cells[0][0] if total_cells else ""
    total = total or reliable
    Image.fromarray(warped).save(preview)
    return {"rows":answers,"total":total,"total_raw":reported,"method":"four-click" if corners else "auto",
            "bounds":[xs[0]/W,xs[1]/W,xs[2]/W,xs[3]/W,xs[4]/W,ys[1]/H,ys[-2]/H],"blue":bool(blue)}


def _code_key(value):
    return re.sub(r"[^A-Z0-9]", "", value.upper())


CODE_CONFUSIONS = {frozenset(pair) for pair in ("O0","I1","L1","Y4","S5","Z2","B8","B3","G6","06","46","T7","J7","D0")}


def _distance(a,b):
    previous = [float(n) for n in range(len(b)+1)]
    for i,x in enumerate(a,1):
        now = [float(i)]
        for j,y in enumerate(b,1):
            cost = 0 if x==y else .38 if frozenset((x,y)) in CODE_CONFUSIONS else 1
            now.append(min(now[-1]+1,previous[j]+1,previous[j-1]+cost))
        previous = now
    return previous[-1]


def receipt_rows(payload,catalog):
    normalized = {}
    for code in catalog:
        normalized.setdefault(_code_key(code),[]).append(code)
    rows = []
    for item in payload["rows"]:
        values = list(item["values"])
        scores = list(item["scores"])
        flags = list(item["flags"])
        suggestions = []
        samples = [(t,s) for t,s,_ in item["raw"][0] if t and s>=.36]
        leading_key = _code_key(samples[0][0]) if samples else ""
        recognized = {c:s for t,s in reversed(samples) if re.fullmatch(r"[A-Z0-9 ._/-]+",t.upper())
                      for c in normalized.get(_code_key(t),[])}
        guess = normalized.get(leading_key,[])
        top_filtered = bool(samples and re.fullmatch(r"[A-Z0-9 ._/-]+",samples[0][0].upper()))
        if len(recognized)==1 and len(guess)==1 and top_filtered and guess[0] in recognized:
            code,score = guess[0],recognized[guess[0]]
            if values[0]!=code:
                flags.append("货号已按商品表匹配大小写、空格或分隔符，请核对")
            values[0],scores[0] = code,score
        elif len(recognized)>1:
            suggestions = sorted(recognized)[:4]
            flags.append("不同识别结果对应多个货号："+ "、".join(suggestions))
            # Different OCR variants match different real products. Do not
            # enable confirmation until the operator enters the code explicitly.
            values[0] = ""
        elif samples:
            # The exact match must come from the leading OCR hypothesis.
            # Never replace a code by the lower-ranked or fuzzy alternative.
            probes = {}
            for text,confidence in samples:
                key = _code_key(text)
                if len(key)>=4:
                    probes[key] = max(probes.get(key,0.), confidence)
                if re.search(r"[0-9][=+]$",text) and len(key)>=4:
                    probes[key+"7"] = max(probes.get(key+"7",0.),confidence-.12)
            distances = {}
            for key,codes in normalized.items():
                for probe,conf in probes.items():
                    if abs(len(key)-len(probe))<=2:
                        d = _distance(probe,key)
                        if d<=2.65:
                            for code in codes:
                                rank = d + .48*(1-conf)
                                distances[code] = min(distances.get(code,1000),rank)
            if distances or recognized:
                candidates = sorted(distances,key=lambda code:(distances[code],code))[:4]
                suggestions = list(dict.fromkeys(candidates+list(recognized)))[:5]
                flags.append("货号候选（未替换）："+"、".join(suggestions))
        extended = {}
        for text,score in item.get("extra_codes",[]):
            if score < .43 or not re.fullmatch(r"[A-Z0-9 ._/-]+",text.upper()):
                continue
            for entry in normalized.get(_code_key(text),[]):
                extended[entry]=max(extended.get(entry,0.),score)
        if extended:
            known=sorted(extended,key=lambda c:(-extended[c],c))
            flags.append("出框补读货号候选（未替换）："+"、".join(known[:4]))
            suggestions=list(dict.fromkeys(suggestions+known[:4]))[:9]
        if item.get("gutter",{}).get("right_ambiguous") and values[0] in catalog:
            if values[0] not in suggestions:
                suggestions.insert(0,values[0])
            values[0] = ""
        flags.insert(0,f"票据第 {item['row']} 行")
        result = ReceiptRow(values,tuple(item["box"]),scores,flags)
        result.receipt_code_candidates = tuple(suggestions)
        rows.append(result)
    return rows


def read_e8_export(path):
    """Index the first sheet of an Easy8 goods export; never modify the file."""
    if path.suffix.lower() not in (".xls", ".xlsx"):
        raise ValueError("请选择易吧商品资料导出的 .xls 或 .xlsx。")
    before = signature([{"path": str(path)}])
    hits = {}
    with catalog_sheet(path, None) as (_, read_rows):
        records = iter(read_rows())
        headers = [canonical_header(c.value) for c in next(records, ())]
        if headers.count("商品编码") != 1 or headers.count("商品名称") != 1:
            raise ValueError("找不到唯一的「商品编码」和「商品名称」列，请在易吧商品资料重新导出。")
        c, n = headers.index("商品编码"), headers.index("商品名称")
        for row in records:
            code, name = excel_code(row[c]), cell_text(row[n].value)
            if not code or not name:
                continue
            hits.setdefault(code.casefold(), []).append((code, name))
    if not hits:
        raise ValueError("这份易吧商品档案没有可核对的商品。")
    if before != signature([{"path": str(path)}]):
        raise ValueError("读取期间易吧导出文件发生变化，请重新选择。")
    return hits, before


class ReceiptApp(CatalogApp):
    def __init__(self,root):
        self._standard_original = None
        self._standard_corners = None
        self._corner_dots = None
        self._standard_overlay = False
        self._queue_standard = queue.Queue()
        self._suggestion_row = None
        super().__init__(root)

    def build(self):
        super().build()
        from tkinter import ttk
        action_bar = self.action_buttons[0].master
        panel = ttk.Frame(self.root,padding=(14,3,14,7))
        panel.pack(after=action_bar,fill="x")
        ttk.Label(panel,text="标准票据｜货号・两个数值列・金额").pack(side="left",padx=(0,8))
        self._order = ttk.Combobox(panel,state="readonly",values=ORDER_OPTIONS,width=22)
        self._order.current(0)
        self._order.pack(side="left",padx=(0,7))
        self._standard_buttons = []
        for title,run,style in (("识别标准票",self.run_standard,"Accent.TButton"),
                                ("点选四角",self.choose_corners,"TButton"),
                                ("查看原图",self.show_original,"TButton")):
            b = ttk.Button(panel,text=title,command=run,style=style)
            b.pack(side="left",padx=3)
            self._standard_buttons.append(b)
        suggestions = ttk.Frame(self.product.master)
        suggestions.grid(row=5,column=0,columnspan=4,sticky="ew",padx=3,pady=3)
        ttk.Label(suggestions,text="票据货号候选").pack(side="left")
        self._code_choice = ttk.Combobox(suggestions,state="disabled",width=16)
        self._code_choice.pack(side="left",padx=6)
        self._code_accept = ttk.Button(suggestions,text="使用选中的货号",command=self.use_code_candidate)
        self._code_accept.state(["disabled"])
        self._code_accept.pack(side="left")
        self._code_choice.bind("<<ComboboxSelected>>",self._suggestion_selected)

    def details(self):
        super().details()
        row = self.rows[self.i] if self.i is not None and self.i<len(self.rows) else None
        if row is self._suggestion_row:
            return
        self._suggestion_row = row
        codes = getattr(row,"receipt_code_candidates",())
        self._code_choice.configure(values=codes,state="readonly" if codes else "disabled")
        self._code_choice.set("请选择候选" if codes else "无候选")
        self._code_accept.state(["disabled"])

    def _suggestion_selected(self,_event=None):
        available = self._code_choice.get() in getattr(self._suggestion_row,"receipt_code_candidates",())
        self._code_accept.state(["!disabled"] if available else ["disabled"])

    def use_code_candidate(self):
        selected = self._code_choice.get()
        if self.i is not None and selected in getattr(self._suggestion_row,"receipt_code_candidates",()):
            self.values[0].set(selected)
            self.status.set(f"已填入候选 {selected}，仍须对照原票并确认此行。")


    def assistant(self):
        if self.assist_button.instate(["disabled"]) or not self._ready():
            return
        self._close_assistants()
        codes = [r.values[0].strip() for r in self.rows]
        chosen = filedialog.askopenfilename(
            parent=self.root, title="选择当前账套刚导出的易吧商品资料",
            filetypes=[("易吧商品资料 Excel", "*.xls *.xlsx")])
        if not chosen:
            self.status.set("未选择易吧商品档案，未打开录入助手。")
            return
        path = Path(chosen)
        try:
            hits, stamp = read_e8_export(path)
        except Exception as exc:
            messagebox.showerror("无法核对易吧档案", str(exc), parent=self.root)
            return
        blocked, found = [], []
        for i, code in enumerate(codes, 1):
            entries = hits.get(code.casefold(), [])
            if len(entries) == 1 and entries[0][0] == code:
                found.append((code, entries[0][1]))
            else:
                reason = ("没有与票面完全同写法的编码" if len(entries) <= 1
                          else "该编码在档案中不唯一")
                blocked.append(f"第 {i} 行 {code}：{reason}")
        if blocked:
            self.status.set(f"易吧快照中 {len(blocked)} 行未通过核对，整张票未送入助手。")
            preview = blocked[:10] + ([f"另有 {len(blocked)-10} 行"] if len(blocked)>10 else [])
            messagebox.showwarning("易吧货号未核对通过",
                "\n".join(preview) + "\n\n请核对原票和实际编码，勿凭名称选货。", parent=self.root)
            return
        names = "\n".join(f"{code}：{name}" for code, name in found[:8])
        if len(found) > 8:
            names += f"\n另有 {len(found)-8} 行"
        if not messagebox.askyesno("继续手动录入？",
                f"快照 {path.name} 中有 {len(found)} 行同号（并非实时库存）。\n{names}"
                "\n\n请确认是当前账套刚导出的档案，并在易吧逐行核对商品、规格和金额。"
                "\n继续只打开复制助手，不录入或保存易吧单据。", parent=self.root):
            return
        if stamp != signature([{"path": str(path)}]):
            self.status.set("易吧导出文件已变化，请重新核对。")
            return
        if codes != [r.values[0].strip() for r in self.rows]:
            self.status.set("票面货号已变化，请重新核对。")
            return
        super().assistant()

    def load_photo(self,path,demo):
        old = self.photo
        super().load_photo(path,demo)
        if self.photo is not old:
            self._standard_original = self.photo
            self._standard_corners = self._corner_dots = None
            self._standard_overlay = False
            if not demo:
                self.status.set("标准收据可选数值列顺序后点「识别标准票」；其他版式仍可使用「本地 OCR」。")

    def rotate(self):
        if not self._standard_overlay:
            original = self.photo
            super().rotate()
            if self.photo is not original:
                self._standard_original = self.photo
                self._standard_corners = self._corner_dots = None
        else:
            self.status.set("当前预览已拉正。要旋转原图，请重新打开照片后旋转。")

    def choose_corners(self):
        if self.job or self._standard_original is None:
            return
        if self._standard_overlay:
            self.status.set("请先重新打开原照片，再点选四角；当前拉正预览和明细未动。")
            return
        self.calibration = None
        self._corner_dots = []
        self.status.set("点选标准表格的左上角（含表头和合计整张表）。按 Esc 可取消。")
        self.canvas.focus_set()
        self.redraw()

    def begin_calibration(self):
        self._corner_dots = None
        super().begin_calibration()

    def cancel_calibration(self,_event=None):
        if self._corner_dots is not None:
            self._corner_dots = None
            self.redraw()
            self.status.set("已取消手动点选。")
        else:
            super().cancel_calibration(_event)

    def redraw(self):
        super().redraw()
        if self._corner_dots and self.draw_state:
            x,y,w,h = self.draw_state
            points = [(x+u*w,y+v*h) for u,v in self._corner_dots]
            for j,(px,py) in enumerate(points,1):
                self.canvas.create_oval(px-5,py-5,px+5,py+5,fill="#d56c15",outline="white")
                self.canvas.create_text(px+12,py,text=str(j),fill="#132a37")
            if len(points)>1:
                self.canvas.create_line(*[a for xy in points for a in xy],fill="#d56c15",width=2)

    def photo_click(self,event):
        if self._corner_dots is None:
            return super().photo_click(event)
        if not self.draw_state:
            return
        x,y,w,h = self.draw_state
        u,v = (event.x-x)/w,(event.y-y)/h
        if not (0<=u<=1 and 0<=v<=1):
            return
        self._corner_dots.append((u,v))
        labels = ("右上角","右下角","左下角")
        if len(self._corner_dots)<4:
            self.status.set(f"继续点选表格的{labels[len(self._corner_dots)-1]}。")
        else:
            w,h = self._standard_original.size
            self._standard_corners = [(u*w,v*h) for u,v in self._corner_dots]
            self._corner_dots = None
            self.status.set("四角已选好。点「识别标准票」使用手动拉正。")
        self.redraw()

    def show_original(self):
        if self._standard_original is None:
            return
        from PIL import Image, ImageTk
        im = self._standard_original.copy()
        im.thumbnail((920,780),Image.Resampling.LANCZOS)
        d = tk.Toplevel(self.root)
        d.title("票据原图（只读）")
        l = tk.Label(d)
        l.image = ImageTk.PhotoImage(im)
        l.configure(image=l.image)
        l.pack()

    def run_standard(self):
        if self.job or self._standard_original is None or not self.replace_ok():
            return
        source = self._standard_original
        self.calibration = self._corner_dots = None
        swapped = self._order.current()==0
        self.status.set("正在本机拉正表格并逐格识别。原始照片和易吧不会改动。")
        for b in (*self.action_buttons,*self._standard_buttons):
            b.state(["disabled"])
        self._order.configure(state="disabled")
        self.job = threading.Thread(target=self._standard_job,args=(source,swapped,self._standard_corners),daemon=True)
        self.job.start()
        self.root.after(125,self._poll_standard)

    def _standard_job(self,source,swapped,corners):
        from PIL import Image
        try:
            with tempfile.TemporaryDirectory(dir=scaner.ROOT,prefix=".receipt-") as directory:
                folder = Path(directory)
                raw,view,out = (folder/n for n in ("photo.png","straight.png","rows.json"))
                source.save(raw)
                argv = [sys.executable,str(Path(__file__).resolve()),"--receipt-worker",
                        str(raw),str(view),str(out),"1" if swapped else "0",json.dumps(corners)]
                p = subprocess.run(argv,capture_output=True,encoding="utf-8",errors="replace",timeout=240,
                                   creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0))
                if p.returncode:
                    if out.is_file():
                        raise ValueError(json.loads(out.read_text(encoding="utf-8"))["error"])
                    raise RuntimeError("识别进程失败："+"\n".join(p.stderr.splitlines()[-3:]))
                payload = json.loads(out.read_text(encoding="utf-8"))
                with Image.open(view) as f:
                    image = f.convert("RGB")
                self._queue_standard.put((payload,image,None))
        except Exception as exc:
            self._queue_standard.put((None,None,str(exc)))

    def _poll_standard(self):
        try:
            payload,view,error = self._queue_standard.get_nowait()
        except queue.Empty:
            self.root.after(125,self._poll_standard)
            return
        self.job = None
        self._order.configure(state="readonly")
        for b in (*self.action_buttons,*self._standard_buttons):
            b.state(["!disabled"])
        if error or not payload["rows"]:
            self.status.set("标准票未完成识别，现有明细保留。")
            messagebox.showwarning("未完成标准票识别",error or "没有可靠的明细。请重拍或手动点选四角。",parent=self.root)
            return
        self.photo = view
        self.bounds = tuple(payload["bounds"])
        self._standard_overlay = True
        self.rows = receipt_rows(payload,self.catalog)
        self.total_input.set(payload["total"] or "")
        self.refresh(0)
        extra = f"票面合计读到 {payload['total']}，也请核对。" if payload["total"] else "合计未读清，请按原票填入。"
        self.status.set(f"已拉正并识别出 {len(self.rows)} 行。{extra} 请检查漏行、货号和涂改项。")


def launch(check_only=False):
    root = tk.Tk()
    if check_only:
        root.withdraw()
    app = ReceiptApp(root)
    root.title("Scaner · 标准票据（核对易吧档案）")
    if len(app.counts) == 4 and not app.catalog_waiting:
        app.status.set("已载入四份商品表。打开照片后点「识别标准票」；默认按样票顺序读取单价和数量。")
    if check_only:
        ready = len(app.counts) == len(app.srcs) == 4 and not app.catalog_waiting
        print(app.catalog_text.get()+" · 标准票据模式已开启 · 货号自适应出框 · 易吧档案预检")
        root.destroy()
        return 0 if ready else 1
    root.mainloop()
    return 0


if __name__ == "__main__":
    if "--receipt-worker" in sys.argv[1:]:
        _,_,src,preview,out,swapped,corners = sys.argv
        try:
            obj = recognize_receipt(Path(src),Path(preview),swapped=="1",json.loads(corners))
        except Exception as exc:
            Path(out).write_text(json.dumps({"error":str(exc)},ensure_ascii=False),encoding="utf-8")
            raise SystemExit(2)
        Path(out).write_text(json.dumps(obj,ensure_ascii=False),encoding="utf-8")
        raise SystemExit(0)
    raise SystemExit(launch("--check" in sys.argv[1:]))
