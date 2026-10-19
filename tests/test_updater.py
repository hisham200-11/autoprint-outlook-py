import os
import sys
import unittest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from updater import check_for_updates, is_version_newer

class TestUpdater(unittest.TestCase):
    def test_version_comparison(self):
        # Newer versions
        self.assertTrue(is_version_newer("1.0.0", "1.0.1"))
        self.assertTrue(is_version_newer("1.0.0", "1.1.0"))
        self.assertTrue(is_version_newer("1.0.0", "2.0.0"))
        self.assertTrue(is_version_newer("v1.0.0", "v1.0.1"))
        
        # Same or older versions
        self.assertFalse(is_version_newer("1.0.1", "1.0.0"))
        self.assertFalse(is_version_newer("1.0.0", "1.0.0"))
        self.assertFalse(is_version_newer("v2.0.0", "v1.9.9"))

    @patch("urllib.request.urlopen")
    def test_check_for_updates_available(self, mock_urlopen):
        mock_response = MagicMock()
        mock_response.read.return_value = b'''{
            "tag_name": "v1.0.1",
            "name": "AutoPrint Outlook v1.0.1",
            "body": "Bugfixes and performance improvements",
            "assets": [
                {
                    "name": "AutoPrint-Outlook-Standalone.zip",
                    "browser_download_url": "https://github.com/hisham200-11/autoprint-outlook-py/releases/download/v1.0.1/AutoPrint-Outlook-Standalone.zip"
                }
            ]
        }'''
        mock_urlopen.return_value.__enter__.return_value = mock_response

        update_info = check_for_updates("hisham200-11/autoprint-outlook-py", current_version="1.0.0")
        self.assertTrue(update_info["update_available"])
        self.assertEqual(update_info["latest_version"], "v1.0.1")
        self.assertEqual(update_info["download_url"], "https://github.com/hisham200-11/autoprint-outlook-py/releases/download/v1.0.1/AutoPrint-Outlook-Standalone.zip")

if __name__ == "__main__":
    unittest.main()
