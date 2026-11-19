from pathlib import Path
from hashlib import sha256
import tempfile
import unittest

from openpyxl import Workbook, load_workbook
from core import ReceiptRow, align_boxes, decimal_input, export_review, load_catalog, validate

ROOT = Path(__file__).resolve().parent
DEMO = ROOT / "demo"


class CoreTests(unittest.TestCase):
    def setUp(self):
        self.sources = [dict(path=p) for p in sorted(DEMO.glob("catalog-*.xlsx"))]
        self.catalog, self.counts = load_catalog(self.sources)

    def test_four_catalogs_read_only_and_exact_text(self):
        hashes = [sha256(x["path"].read_bytes()).hexdigest() for x in self.sources]
        self.assertEqual(len(self.counts), 4)
        self.assertEqual(set(self.catalog), {"00123","SC-2177","028609","A-008"})
        self.assertNotIn("123", self.catalog)
        self.assertEqual(hashes,[sha256(x["path"].read_bytes()).hexdigest() for x in self.sources])

    def test_column_mapping_different_headers_and_duplicate_conflict(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d)/"different.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "第二套"
            ws.append(["说明"])
            ws.append(["自定义品名", "内部货号", "价格"])
            ws.append(["同编号不同商品", 123, 8.25])
            ws["B3"].number_format = "00000"
            wb.save(path)
            wb.close()
            cat,_ = load_catalog([self.sources[0],{"path":path,"sheet":"第二套","header":2,"code":"#B","name":"自定义品名","price":"价格"}])
            row = ReceiptRow(["00123","2","12.50","25.00"])
            self.assertIn("冲突", validate(row, cat)[2][0])
            row.product_index = 1
            p,_,errs,_ = validate(row,cat)
            self.assertEqual(p.name,"同编号不同商品")
            self.assertEqual(p.price,"8.25")
            self.assertEqual(errs,[])

    def test_numeric_and_confidence_checks(self):
        self.assertIsNone(decimal_input("1,25"))
        self.assertIsNone(decimal_input("12.OO"))
        row = ReceiptRow(["00123","2","12.505","25.01"], scores=[.80, .96, .99, .99])
        p,items,errs,warnings = validate(row, self.catalog)
        self.assertEqual(errs, [])
        self.assertEqual(str(items[-1]), "25.01")
        self.assertIn("低置信", warnings[0])
        row.values[3] = "25.10"
        self.assertIn("金额不符",validate(row,self.catalog)[2][0])

    def test_alignment_catches_missing_cells(self):
        samples = [{"rect":[x,y,x+12,y+10],"text":v,"score":.99}
                   for x,y,v in [(10,10,"00123"),(35,11,"2"),(57,10,"12.5"),(81,12,"25"),(10,40,"X-8"),(57,42,"8.25"),(81,41,"16.50")]]
        rows = align_boxes(samples,(0,.25,.50,.75,1,0,1),(100,100))
        self.assertEqual([r.values for r in rows],[["00123","2","12.5","25"],["X-8","","8.25","16.50"]])

    def test_export_requires_confirmation_and_writes_text_codes(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d)/"review.xlsx"
            row = ReceiptRow(["00123","2","12.50","25.00"])
            with self.assertRaises(ValueError):
                export_review(path,[row],self.catalog)
            self.assertFalse(path.exists())
            row.confirmed = True
            export_review(path,[row],self.catalog)
            wb = load_workbook(path, read_only=True)
            self.assertEqual(wb.active["A2"].value,"00123")
            self.assertEqual(wb.active["A2"].data_type,"s")
            self.assertEqual(wb.active["G2"].value,25)
            wb.close()
            with self.assertRaises(ValueError):
                export_review(self.sources[0]["path"],[row],self.catalog,[s["path"] for s in self.sources])


if __name__ == "__main__":
    unittest.main(verbosity=2)
