"""Synthetic sample files only; never reads business data."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent
DEMO = ROOT / "demo"
DEMO.mkdir(exist_ok=True)
goods = [
    ("00123", "样品胶带", "48mm", "卷", 12.5),
    ("SC-2177", "样品灯泡", "12V", "盒", 8.25),
    ("028609", "样品清洁剂", "300ml", "瓶", 15),
    ("A-008", "样品手套", "L", "双", 9),
]
for i, product in enumerate(goods, 1):
    wb = Workbook()
    ws = wb.active
    ws.title = "商品资料"
    ws.append(("商品编码", "商品名称", "规格", "单位", "参考价"))
    ws.append(product)
    ws["A2"].number_format = "@"
    wb.save(DEMO / f"catalog-{i}.xlsx")
    wb.close()

im = Image.new("RGB", (1240, 850), "#fffdf8")
d = ImageDraw.Draw(im)
fonts = Path("C:/Windows/Fonts")
cn = fonts / "msyh.ttc"
latin = fonts / "arial.ttf"
title, label, hand = (ImageFont.truetype(str(p), s) for p, s in ((cn, 53), (cn, 34), (latin, 58)))
d.text((90, 54), "手写单据 · 功能演示", font=title, fill="#1b2f3b")
d.text((93, 129), "合成样例，不包含真实商品或客户资料", font=label, fill="#586775")
for x, s in zip((120, 470, 714, 951), ("商品编码", "数量", "单价", "金额")):
    d.text((x, 190), s, font=label, fill="#263947")
for y, row in zip((280, 412, 544), (("00123", "2", "12.50", "25.00"), ("SC-2177", "3", "8.25", "24.75"), ("028609", "1", "15.00", "15.00"))):
    for x, val in zip((120, 470, 714, 951), row):
        d.text((x, y), val, font=hand, fill="#1a244d")
for y in (248, 384, 515, 647):
    d.line((90, y, 1153, y), fill="#cbd4d9", width=2)
d.text((88, 700), "整单合计", font=label, fill="#243845")
d.text((951, 688), "64.75", font=hand, fill="#1a244d")
im.save(DEMO / "receipt.png")
