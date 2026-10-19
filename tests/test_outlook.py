import os
import sys
import unittest
from datetime import datetime, date
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from outlook_adapter import build_outlook_date_filter, get_file_type_category, SUPPORTED_EXTENSIONS

class TestOutlookAdapter(unittest.TestCase):
    def test_build_date_filter_single_day(self):
        d = date(2026, 10, 19)
        # Should return restriction string format for Outlook MAPI
        filter_str = build_outlook_date_filter(start_date=d, end_date=d)
        self.assertIn("[ReceivedTime] >=", filter_str)
        self.assertIn("[ReceivedTime] <=", filter_str)
        self.assertIn("10/19/2026", filter_str)

    def test_build_date_filter_range(self):
        d_start = date(2026, 10, 1)
        d_end = date(2026, 10, 19)
        filter_str = build_outlook_date_filter(start_date=d_start, end_date=d_end)
        self.assertIn("10/01/2026", filter_str)
        self.assertIn("10/19/2026", filter_str)

    def test_supported_extensions_and_categories(self):
        self.assertIn(".pdf", SUPPORTED_EXTENSIONS)
        self.assertIn(".png", SUPPORTED_EXTENSIONS)
        self.assertIn(".jpg", SUPPORTED_EXTENSIONS)
        self.assertIn(".jpeg", SUPPORTED_EXTENSIONS)
        self.assertIn(".xlsx", SUPPORTED_EXTENSIONS)
        self.assertIn(".docx", SUPPORTED_EXTENSIONS)

        self.assertEqual(get_file_type_category("report.pdf"), "PDF")
        self.assertEqual(get_file_type_category("photo.png"), "Image")
        self.assertEqual(get_file_type_category("sheet.xlsx"), "Excel")
        self.assertEqual(get_file_type_category("doc.docx"), "Word")
        self.assertIsNone(get_file_type_category("archive.zip"))

if __name__ == "__main__":
    unittest.main()
