import os
import re
import json
import shutil
import time
from dataclasses import dataclass, asdict
from typing import Optional, Dict

@dataclass
class AttachmentRecord:
    file_id: str
    message_id: str
    sender: str
    subject: str
    received_time: str
    filename: str
    local_path: str
    page_count: int = 1
    is_printed: bool = False
    printed_at: Optional[str] = None
    archived_path: Optional[str] = None

class StorageManager:
    def __init__(self, pending_dir: str = "Downloads/Pending", archive_dir: str = "Printed_Archive", history_file: str = "printed_history.json"):
        self.pending_dir = os.path.abspath(pending_dir)
        self.archive_dir = os.path.abspath(archive_dir)
        self.history_file = os.path.abspath(history_file)
        
        os.makedirs(self.pending_dir, exist_ok=True)
        os.makedirs(self.archive_dir, exist_ok=True)
        self.history: Dict[str, dict] = self._load_history()

    def _load_history(self) -> Dict[str, dict]:
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _save_history(self):
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.history, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving history: {e}")

    def sanitize_filename(self, name: str) -> str:
        # Replace invalid Windows filename characters
        return re.sub(r'[\\/*?:"<>|]', "_", name)

    def is_already_printed(self, file_id: str) -> bool:
        return file_id in self.history

    def get_destination_archive_path(self, record: AttachmentRecord) -> str:
        safe_sender = self.sanitize_filename(record.sender or "Unknown_Sender")
        sender_archive_dir = os.path.join(self.archive_dir, safe_sender)
        os.makedirs(sender_archive_dir, exist_ok=True)
        
        base_name = self.sanitize_filename(record.filename)
        dest_path = os.path.join(sender_archive_dir, base_name)
        
        # If file exists, avoid overwrite by attaching timestamp or count
        if os.path.exists(dest_path):
            name, ext = os.path.splitext(base_name)
            timestamp = int(time.time())
            dest_path = os.path.join(sender_archive_dir, f"{name}_{timestamp}{ext}")
            
        return dest_path

    def archive_printed_file(self, record: AttachmentRecord) -> str:
        dest_path = self.get_destination_archive_path(record)
        
        # Retry loop to account for Windows file locks
        for attempt in range(5):
            try:
                if os.path.exists(record.local_path):
                    shutil.move(record.local_path, dest_path)
                    break
                else:
                    break
            except PermissionError:
                time.sleep(0.3)
                
        record.is_printed = True
        record.printed_at = time.strftime("%Y-%m-%d %H:%M:%S")
        record.archived_path = dest_path
        
        self.history[record.file_id] = asdict(record)
        self._save_history()
        return dest_path

    def clear_pending_files(self) -> int:
        """
        Deletes all unprinted/staging files currently stored in Downloads/Pending to free disk space.
        Returns the number of files deleted.
        """
        deleted_count = 0
        if os.path.exists(self.pending_dir):
            for filename in os.listdir(self.pending_dir):
                file_path = os.path.join(self.pending_dir, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                        deleted_count += 1
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                        deleted_count += 1
                except Exception as e:
                    print(f"Error deleting pending file {file_path}: {e}")
        return deleted_count

    def clear_archive_files(self) -> int:
        """
        Deletes all archived printed files in Printed_Archive to free disk space.
        Returns the number of files deleted.
        """
        deleted_count = 0
        if os.path.exists(self.archive_dir):
            for root, dirs, files in os.walk(self.archive_dir, topdown=False):
                for f in files:
                    fp = os.path.join(root, f)
                    try:
                        os.unlink(fp)
                        deleted_count += 1
                    except Exception:
                        pass
                for d in dirs:
                    dp = os.path.join(root, d)
                    try:
                        os.rmdir(dp)
                    except Exception:
                        pass
        return deleted_count
