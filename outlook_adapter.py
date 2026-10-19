import os
import sys
from datetime import datetime, date, time as dtime
from typing import List, Dict, Optional, Tuple, Callable
import pypdf

SUPPORTED_EXTENSIONS = {
    ".pdf": "PDF",
    ".png": "Image",
    ".jpg": "Image",
    ".jpeg": "Image",
    ".xlsx": "Excel",
    ".xls": "Excel",
    ".docx": "Word",
    ".doc": "Word",
}

def get_file_type_category(filename: str) -> Optional[str]:
    ext = os.path.splitext(filename)[1].lower()
    return SUPPORTED_EXTENSIONS.get(ext, None)

def get_document_page_count(local_path: str, category: str) -> int:
    try:
        if category == "PDF":
            reader = pypdf.PdfReader(local_path)
            return len(reader.pages)
        elif category == "Image":
            return 1
        elif category in ("Word", "Excel"):
            return 1
    except Exception:
        pass
    return 1

# Format for Outlook DASL / Jet query: "MM/dd/yyyy hh:mm AMPM"
def build_outlook_date_filter(start_date: date, end_date: date) -> str:
    start_str = start_date.strftime("%m/%d/%Y 00:00 AM")
    end_str = end_date.strftime("%m/%d/%Y 11:59 PM")
    return f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] <= '{end_str}'"

class OutlookAdapter:
    def __init__(self):
        self._outlook = None
        self._namespace = None

    def connect(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
            import win32com.client
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            return True, "Connected to Outlook successfully"
        except Exception as e:
            return False, f"Could not connect to Outlook: {str(e)}"

    def get_available_folders(self) -> List[Tuple[str, str]]:
        """
        Returns list of (display_name, folder_id or path)
        """
        if not self._namespace:
            success, msg = self.connect()
            if not success:
                return [("Default Inbox", "inbox")]
        
        folders = [("Default Inbox", "inbox")]
        try:
            # Add top-level Inbox subfolders and main folders
            inbox = self._namespace.GetDefaultFolder(6) # 6 = olFolderInbox
            for sub in inbox.Folders:
                folders.append((f"Inbox > {sub.Name}", sub.EntryID))
                
            # Check default store root folders
            root = inbox.Parent
            for folder in root.Folders:
                name = folder.Name
                if name.lower() in ["to print", "print", "printed", "archive", "action"]:
                    folders.append((f"Folder: {name}", folder.EntryID))
        except Exception:
            pass
        return folders

    def scan_attachments(
        self,
        folder_id: str,
        start_date: Optional[date],
        end_date: Optional[date],
        storage_manager,
        progress_callback: Optional[Callable[[str, float], None]] = None
    ) -> List[dict]:
        """
        Scans emails in selected folder, extracts PDF attachments to storage_manager.pending_dir,
        and returns list of attachment metadata dictionaries.
        """
        # COM apartment threading requirement: each thread must call CoInitialize
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        ns = outlook_app.GetNamespace("MAPI")

        try:
            if folder_id == "inbox":
                target_folder = ns.GetDefaultFolder(6)
            else:
                target_folder = ns.GetFolderFromID(folder_id)
        except Exception as e:
            raise RuntimeError(f"Error accessing folder: {e}")

        items = target_folder.Items
        items.Sort("[ReceivedTime]", True) # Descending (newest first)

        if start_date and end_date:
            date_filter = build_outlook_date_filter(start_date, end_date)
            try:
                items = items.Restrict(date_filter)
            except Exception as e:
                print(f"Warning: Outlook Restrict query failed ({e}), falling back to iterative filter.")

        results = []
        total_items = len(items) if hasattr(items, '__len__') else 0
        count = 0

        for item in items:
            count += 1
            if progress_callback and total_items > 0:
                progress_callback(f"Scanning email {count} of {total_items}...", count / max(total_items, 1))

            try:
                # Check for attachments
                if not hasattr(item, "Attachments") or item.Attachments.Count == 0:
                    continue

                # Metadata
                subject = getattr(item, "Subject", "No Subject")
                received_time_val = getattr(item, "ReceivedTime", None)
                received_time_str = str(received_time_val) if received_time_val else ""
                
                # Sender
                sender = "Unknown Sender"
                try:
                    sender = item.SenderName
                    if hasattr(item, "SenderEmailAddress") and item.SenderEmailAddress:
                        sender = f"{item.SenderName} <{item.SenderEmailAddress}>"
                except Exception:
                    pass

                entry_id = getattr(item, "EntryID", f"email_{count}")

                # Process attachments
                for i in range(1, item.Attachments.Count + 1):
                    att = item.Attachments.Item(i)
                    filename = att.FileName
                    category = get_file_type_category(filename)
                    if not category:
                        continue

                    # File identifier
                    file_id = f"{entry_id}_{i}_{storage_manager.sanitize_filename(filename)}"
                    
                    # Safe destination in pending
                    safe_filename = f"{count}_{storage_manager.sanitize_filename(filename)}"
                    local_path = os.path.join(storage_manager.pending_dir, safe_filename)

                    # Extract attachment if not already saved
                    if not os.path.exists(local_path):
                        att.SaveAsFile(local_path)

                    page_count = get_document_page_count(local_path, category)
                    is_printed = storage_manager.is_already_printed(file_id)

                    att_dict = {
                        "file_id": file_id,
                        "message_id": entry_id,
                        "sender": sender,
                        "subject": subject,
                        "received_time": received_time_str,
                        "filename": filename,
                        "category": category,
                        "local_path": local_path,
                        "page_count": page_count,
                        "is_printed": is_printed
                    }
                    results.append(att_dict)

            except Exception as e:
                print(f"Error processing item: {e}")
                continue

        if progress_callback:
            progress_callback(f"Scan complete. Found {len(results)} PDF attachments.", 1.0)

        return results
