import unittest
import os
import sys
from unittest.mock import MagicMock
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from print_spooler import SequentialPrintEngine, PrintJob

class TestPrintSpooler(unittest.TestCase):
    def test_queue_enqueuing_and_callbacks(self):
        engine = SequentialPrintEngine(delay_seconds=0.01)
        mock_callback = MagicMock()
        engine.set_status_callback(mock_callback)

        job1 = PrintJob(
            file_id="id1",
            file_path="dummy1.pdf",
            sender="alice@test.com",
            filename="doc1.pdf",
            page_count=2
        )
        engine.enqueue_jobs([job1])
        
        self.assertEqual(engine.get_queue_length(), 1)
        self.assertFalse(engine.is_running())

if __name__ == "__main__":
    unittest.main()
