"""Read-only Excel lookup, OCR box alignment, and receipt review."""
from __future__ import annotations

from collections import defaultdict
from contextlib import contextmanager
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from statistics import median
from types import SimpleNamespace
import os
import re
import tempfile
import unicodedata

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.cell import column_index_from_string

FIELDS = ("code", "name", "spec", "unit", "price")
HEADERS = {
    "code": ("商品编码", "商品编号", "货号", "编号", "编码", "物品编号", "SKU", "条码", "Barcode"),
    "name": ("商品名称", "商品名", "品名", "名称", "物品名称"),
    "spec": ("商品规格", "规格", "规格型号", "型号"),
    "unit": ("计量单位", "单位"),
    "price": ("参考价", "参考售价", "零售价", "售价", "单价", "参考进价"),
}
LABELS = ("商品编码", "数量", "单价", "票面金额")
CENT = Decimal("0.01")


def canonical_header(value):
    return "".join(unicodedata.normalize("NFKC", str(value or "")).split()).casefold()


def cell_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def excel_code(cell):
    value = getattr(cell, "value", None)
    fmt = getattr(cell, "number_format", "") or ""
    if isinstance(value, bool) or (value is not None and not isinstance(value, (int, float, str))):
        raise ValueError("编码列存在日期或布尔值；请在商品表副本中将商品编码存成文本。")
    if isinstance(value, (int, float)) and abs(value) >= 10**15:
        raise ValueError("编码有 16 位以上数字且以数值保存，Excel 可能已丢失精度；请提供文本编码。")
    if isinstance(value, (int, float)) and value >= 0 and int(value) == value and re.fullmatch(r"0{2,}", fmt):
        return str(int(value)).zfill(len(fmt))
    return cell_text(value)


@dataclass(frozen=True)
class Product:
    code: str
    name: str
    spec: str
    unit: str
    price: str
    source: str

    @property
    def details(self):
        return (self.name, self.spec, self.unit, self.price)


def find_column(header, config, key):
    selector = (config.get(key) or "").strip()
    if selector.startswith("#"):
        return column_index_from_string(selector[1:].upper()) - 1
    candidates = (selector,) if selector else HEADERS[key]
    for candidate in candidates:
        hits = [i for i, value in enumerate(header) if canonical_header(candidate) == canonical_header(value)]
        if len(hits) > 1:
            raise ValueError(f"表头「{candidate}」重复，请用 #列字母 指定具体列。")
        if hits:
            return hits[0]
    if selector or key == "code":
        raise ValueError(f"找不到列「{selector or '商品编码'}」；请填写表头或 #列字母（如 #A）。")
    return None


@contextmanager
def catalog_sheet(path, title):
    """Yield a uniform read-only worksheet interface for XLS or OOXML."""
    if path.suffix.lower() == ".xls":
        import xlrd

        wb = xlrd.open_workbook(str(path), on_demand=True, formatting_info=True, use_mmap=True)
        try:
            sheet = wb.sheet_by_name(title) if title else wb.sheet_by_index(0)

            def legacy_rows(min_row=1, max_row=None, max_col=None, values_only=False):
                width = sheet.ncols if max_col is None else max_col
                stop = sheet.nrows if max_row is None else min(max_row, sheet.nrows)
                for r in range(min_row - 1, stop):
                    row = []
                    for c in range(width):
                        cell = sheet.cell(r, c) if c < sheet.ncols else None
                        value, fmt = None, ""
                        if cell is not None:
                            if cell.ctype == xlrd.XL_CELL_DATE:
                                value = xlrd.xldate_as_datetime(cell.value, wb.datemode)
                            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                                value = bool(cell.value)
                            elif cell.ctype not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                                value = cell.value
                            if cell.xf_index is not None:
                                number_format = wb.format_map.get(wb.xf_list[cell.xf_index].format_key)
                                if number_format is not None:
                                    fmt = number_format.format_str
                        row.append(value if values_only else SimpleNamespace(value=value, number_format=fmt))
                    yield row

            yield sheet.name, legacy_rows
        finally:
            wb.release_resources()
    else:
        wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            sheet = wb[title] if title else wb.worksheets[0]
            yield sheet.title, sheet.iter_rows
        finally:
            wb.close()


def load_catalog(sources):
    index, counts = defaultdict(list), []
    if len(sources) > 4:
        raise ValueError("最多配置四份商品工作簿。")
    paths = set()
    for config in sources:
        path = Path(config["path"]).resolve()
        if path in paths:
            raise ValueError(f"重复选择工作簿：{path.name}")
        paths.add(path)
        if path.suffix.lower() not in {".xls", ".xlsx", ".xlsm"}:
            raise ValueError(f"{path.name}：请选择 .xls、.xlsx 或 .xlsm。")
        try:
            with catalog_sheet(path, config.get("sheet")) as (sheet_name, iter_rows):
                header_row = int(config.get("header") or 1)
                if header_row < 0:
                    raise ValueError("表头行号不能小于 0。")
                # Header 0 denotes a headerless sheet; columns must be mapped explicitly.
                header = next(iter_rows(min_row=header_row, max_row=header_row, values_only=True), ()) if header_row else ()
                columns = [find_column(header, config, key) for key in FIELDS]
                max_col = max(i for i in columns if i is not None) + 1
                count = 0
                for i, cells in enumerate(iter_rows(min_row=header_row + 1, max_col=max_col), header_row + 1):
                    code_cell = cells[columns[0]]
                    code = excel_code(code_cell)
                    if not code:
                        continue
                    extra = [cell_text(getattr(cells[c], "value", None)) if c is not None else "" for c in columns[1:]]
                    item = Product(code, *extra, f"{path.name} / {sheet_name}:{i}")
                    if all(previous.details != item.details for previous in index[code]):
                        index[code].append(item)
                    count += 1
                if not count:
                    raise ValueError("编码列没有可用记录，或公式没有已保存的计算结果。")
                counts.append((path.name, count))
        except Exception as ex:
            raise ValueError(f"{path.name}：{ex}") from ex
    return dict(index), counts


def decimal_input(text):
    # Do not infer decimal separators or substitute OCR glyphs such as O/0.
    s = unicodedata.normalize("NFKC", str(text)).strip().replace("−", "-")
    if not re.fullmatch(r"-?(?:[0-9]{1,12}|[0-9]{1,3}(?:,[0-9]{3}){1,3})(?:\.[0-9]{1,6})?", s):
        return None
    cleaned = s.replace(",", "")
    if len(cleaned.replace("-", "").replace(".", "")) > 15:
        return None
    return Decimal(cleaned)


def plain(number):
    return format(number, "f") if number is not None else ""


@dataclass
class ReceiptRow:
    values: list[str] = field(default_factory=lambda: ["", "", "", ""])
    box: tuple[float, float] | None = None
    scores: list[float | None] = field(default_factory=lambda: [None] * 4)
    ocr_flags: list[str] = field(default_factory=list)
    product_index: int | None = None
    confirmed: bool = False


def validate(row, catalog):
    code, qty_text, price_text, amount_text = row.values
    products = catalog.get(code.strip(), [])
    matched = None
    hard = []
    if not code.strip():
        hard.append("缺少编码")
    elif not products:
        hard.append("编码未匹配")
    elif len(products) == 1:
        matched = products[0]
    elif row.product_index is not None and 0 <= row.product_index < len(products):
        matched = products[row.product_index]
    else:
        hard.append("同编码资料冲突，请选来源")
    qty, price, amount = map(decimal_input, (qty_text, price_text, amount_text))
    if qty is None or qty == 0:
        hard.append("数量无效")
    if price is None or price < 0:
        hard.append("单价无效")
    if amount is None or amount.as_tuple().exponent < -2:
        hard.append("票面金额无效（最多两位小数）")
    expected = (qty * price).quantize(CENT, rounding=ROUND_HALF_UP) if qty is not None and price is not None else None
    if amount is not None and expected is not None and amount != expected:
        hard.append(f"金额不符，应为 {plain(expected)}")
    warnings = list(row.ocr_flags)
    warnings.extend(f"{LABELS[i]}低置信 {score:.0%}" for i, score in enumerate(row.scores) if score is not None and score < 0.86)
    return matched, (qty, price, amount, expected), hard, warnings


def align_boxes(boxes, bounds, size):
    """Assign recognized boxes to one of four calibrated columns and visual rows."""
    w, h = size
    xs = [x * w for x in bounds[:5]]
    top, bottom = (n * h for n in bounds[5:])
    cells = []
    for b in boxes:
        l, t, r, d = b["rect"]
        x, y = (l + r) / 2, (t + d) / 2
        if not (xs[0] <= x <= xs[4] and top <= y <= bottom):
            continue
        col = next((i for i in range(4) if x < xs[i + 1]), 3)
        overlap = max(0, min(r, xs[col + 1]) - max(l, xs[col]))
        cells.append((y, col, l, t, r, d, str(b["text"]).strip(), float(b["score"]), overlap < (r-l) * .78))
    if not cells:
        return []
    limit = max(3.0, median(c[5] - c[3] for c in cells) * .65)
    groups = []
    for cell in sorted(cells):
        y = cell[0]
        if groups and y - median(x[0] for x in groups[-1]) <= limit:
            groups[-1].append(cell)
        else:
            groups.append([cell])
    rows = []
    for g in groups:
        row = ReceiptRow(box=(max(0., min(c[3] for c in g) / h), min(1., max(c[5] for c in g) / h)))
        for col in range(4):
            pieces = sorted((c for c in g if c[1] == col), key=lambda c: c[2])
            if pieces:
                row.values[col] = " ".join(c[6] for c in pieces).strip()
                row.scores[col] = min(c[7] for c in pieces)
            if len(pieces) > 1:
                row.ocr_flags.append(f"{LABELS[col]}有多个片段")
            if any(c[8] for c in pieces):
                row.ocr_flags.append(f"{LABELS[col]}可能跨列")
        rows.append(row)
    return rows


def export_review(path, rows, catalog, protected=()):
    path = Path(path).resolve()
    if any(path == Path(p).resolve() for p in protected):
        raise ValueError("不能覆盖已选商品资料。")
    headers = ("商品编码", "商品名称", "规格", "单位", "数量", "单价", "票面金额", "校验金额", "商品来源")
    wb = Workbook()
    try:
        ws = wb.active
        ws.title = "已确认明细"
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="17313D")
            cell.font = Font(bold=True, color="FFFFFF")
        for row in rows:
            product, (qty, price, amount, expected), errors, _ = validate(row, catalog)
            if not row.confirmed or errors:
                raise ValueError("请先确认所有行，且解决金额与商品资料问题。")
            ws.append((product.code, product.name, product.spec, product.unit, qty, price, amount, expected, product.source))
            for column in (1, 2, 3, 4, 9):
                cell = ws.cell(ws.max_row, column)
                cell.data_type = "s"
                cell.number_format = "@"
            for column in (7, 8):
                ws.cell(ws.max_row, column).number_format = '0.00'
        for column, width in zip("ABCDEFGHI", (18, 30, 20, 10, 12, 15, 16, 16, 40)):
            ws.column_dimensions[column].width = width
        ws.freeze_panes = "E2"
        ws.auto_filter.ref = ws.dimensions
        fd, temp = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
        os.close(fd)
        try:
            wb.save(temp)
            os.replace(temp, path)
        finally:
            Path(temp).unlink(missing_ok=True)
    finally:
        wb.close()
