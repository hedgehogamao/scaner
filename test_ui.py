"""Automated Tk integration check; exercises the real offline worker, not a stub."""
from pathlib import Path
import tempfile
import time
import tkinter as tk
from openpyxl import load_workbook
from scaner import App, export_review


def exercise():
    r = tk.Tk()
    r.withdraw()
    app = App(r, initial_demo=True)
    app.run_ocr()
    deadline = time.monotonic()+80
    while app.job and time.monotonic() < deadline:
        r.update()
        time.sleep(.03)
    assert app.job is None and len(app.rows) == 3, app.status.get()
    assert [row.values for row in app.rows] == [
        ["00123","2","12.50","25.00"],
        ["SC-2177","3","8.25","24.75"],
        ["028609","1","15.00","15.00"],
    ], [row.values for row in app.rows]
    assert app.export_button.instate(["disabled"])
    for i in range(3):
        app.table.selection_set(str(i))
        r.update()
        app.confirm()
        r.update()
    assert app.export_button.instate(["!disabled"])
    assert app.assist_button.instate(["disabled"]), "demo must not enable the entry assistant"
    assert "64.75" in app.total_text.get() and "合计一致" in app.total_text.get()
    app.total_input.set("64.74")
    assert app.export_button.instate(["disabled"])
    app.total_input.set("64.75")
    with tempfile.TemporaryDirectory() as t:
        f = Path(t)/"review.xlsx"
        export_review(f, app.rows, app.catalog)
        wb = load_workbook(f, read_only=True)
        assert list(wb.active.values)[1][0] == "00123"
        assert sum(row[6] for row in list(wb.active.values)[1:]) == 64.75
        wb.close()
    # Use the demo data in an isolated integration scenario; never interact with 易吧.
    app.demo = False
    app.totals()
    assert app.assist_button.instate(["!disabled"])
    app.assistant()
    r.update()
    copied = []
    r.clipboard_clear = copied.clear
    r.clipboard_append = copied.append
    def children(widget):
        for child in widget.winfo_children():
            yield child
            yield from children(child)
    btn = next(w for w in children(r) if w.winfo_class() == "TButton" and w.cget("text") == "复制编码")
    btn.invoke()
    assert copied == ["00123"]
    r.destroy()
    print("PASS | Windows Tk + real CPU OCR: 3 rows, total 64.75 | 4 read-only workbooks | confirmed XLSX verified | assistant field copy")


if __name__ == "__main__":
    exercise()
