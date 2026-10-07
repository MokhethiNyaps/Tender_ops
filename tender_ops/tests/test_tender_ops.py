import hashlib
import sys
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import tender_ops

PROJECT_ROOT = Path(__file__).resolve().parents[2]
OUTPUT_ROOT = PROJECT_ROOT / "tender_ops" / "output"


class TenderOpsTests(unittest.TestCase):
    def test_sha256_file(self):
        path = PROJECT_ROOT / "Affidavit-EME-CSC.pdf"
        self.assertEqual(tender_ops.sha256_file(path), hashlib.sha256(path.read_bytes()).hexdigest())

    def test_pdf_page_count(self):
        path = PROJECT_ROOT / "Affidavit-EME-CSC.pdf"
        self.assertEqual(tender_ops.pdf_page_count(path), 2)

    def test_xlsx_sheet_names(self):
        path = PROJECT_ROOT / tender_ops.WORKBOOK
        self.assertEqual(tender_ops.xlsx_sheet_names(path), ["P's& G's", "C1.5 (2)"])

    def test_generate_outputs_and_keeps_release_closed(self):
        summary = tender_ops.generate(PROJECT_ROOT, OUTPUT_ROOT, date(2026, 10, 6))
        self.assertFalse(summary["submission_ready"])
        self.assertGreaterEqual(summary["source_files"], 5)
        self.assertTrue((OUTPUT_ROOT / "source_manifest.md").exists())
        self.assertTrue((OUTPUT_ROOT / "requirements_register.csv").exists())
        self.assertIn("EXPIRED", (OUTPUT_ROOT / "source_manifest.md").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
