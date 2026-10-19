import unittest
import os
import sys
import shutil
import tempfile
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from storage_manager import StorageManager, AttachmentRecord

class TestStorageManager(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.incoming_dir = os.path.join(self.test_dir, "Pending")
        self.archive_dir = os.path.join(self.test_dir, "Printed_Archive")
        self.db_path = os.path.join(self.test_dir, "printed_history.json")
        self.storage = StorageManager(
            pending_dir=self.incoming_dir,
            archive_dir=self.archive_dir,
            history_file=self.db_path
        )

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_sanitize_filename(self):
        unsafe = "Invoice: 2026/10/19 *draft*?.pdf"
        safe = self.storage.sanitize_filename(unsafe)
        self.assertEqual(safe, "Invoice_ 2026_10_19 _draft__.pdf")

    def test_record_attachment_and_history(self):
        # Create a mock pending file
        os.makedirs(self.incoming_dir, exist_ok=True)
        dummy_file = os.path.join(self.incoming_dir, "test.pdf")
        with open(dummy_file, "wb") as f:
            f.write(b"%PDF-1.4 mock")

        record = AttachmentRecord(
            file_id="msg1_att1",
            message_id="msg1",
            sender="billing@company.com",
            subject="Invoice Oct 2026",
            received_time="2026-10-19 08:00:00",
            filename="test.pdf",
            local_path=dummy_file,
            page_count=2,
            is_printed=False
        )

        self.assertFalse(self.storage.is_already_printed(record.file_id))
        
        # Move to archive
        archived_path = self.storage.archive_printed_file(record)
        self.assertTrue(os.path.exists(archived_path))
        self.assertFalse(os.path.exists(dummy_file))
        self.assertTrue(self.storage.is_already_printed(record.file_id))

if __name__ == "__main__":
    unittest.main()
