# Outlook Bulk Document & Attachment Printer

A clean, user-friendly desktop application designed to extract attachments from classic Microsoft Outlook, organize them by sender/email, and bulk-print them sequentially without jamming the Windows print spooler.

---

## Supported File Formats

- **PDF Documents** (`.pdf`)
- **Images** (`.png`, `.jpg`, `.jpeg`)
- **Word Documents** (`.docx`, `.doc`)
- **Excel Spreadsheets** (`.xlsx`, `.xls`)

---

## Key Features

1. **Multi-Format Extraction & Filtering**:
   - Filter attachments by type with quick chips: **All**, **PDF**, **Image**, **Word**, **Excel**.
   - Filter by date timeline: **Today**, **Yesterday**, **Last 7 Days**, or **All**.
   - Select between default **Inbox** or any dedicated Outlook folder (e.g. `To Print`).
2. **Organized & Grouped by Sender**:
   - Displays type badge, sender name and email, attachment filename, page count, date received, and email subject.
3. **Smart Checklist & State Persistence**:
   - Attachments already printed are recognized automatically and badged with `[✓ Printed]`.
   - Unprinted attachments are pre-checked for immediate 1-click printing.
4. **Anti-Jam Sequential Print Engine**:
   - Dispatches documents strictly **1-by-1**.
   - Word & Excel documents are printed silently in the background via Office COM automation without popup windows.
   - Monitors the Windows print spooler (`win32print`) to prevent document interleaving or printer driver crashes.
   - 2.0s post-spool stabilization delay before dispatching the next document.
5. **Automatic Archiving**:
   - Once successfully sent to the spooler, each PDF is moved into `./Printed_Archive/{Sender}/{Filename}`.
   - Quick **"📂 Open Printed Archive"** button to view printed files anytime.
6. **Clean User Interface**:
   - Plain high-contrast **White, Deep Black, and Vibrant Orange** theme.
   - Real-time status bar showing active document name, print progress, and an emergency **"Cancel Queue"** button.

---

## How to Run

Double-click `run_outlook_printer.bat` in the root folder, or execute via terminal:

```bash
python -m outlook_pdf_tool.app
```

---

## Architecture & Tests

Built strictly following **Capability Architecture** and **Test-Driven Development (TDD)**:
- `outlook_adapter.py`: MAPI COM isolation and indexed query restriction.
- `storage_manager.py`: Path sanitization, atomic moves, and JSON history ledger.
- `print_spooler.py`: FIFO worker thread with Windows print spooler monitoring.
- `app.py`: Tkinter desktop interface.

To run the automated test suite:
```bash
python -m unittest discover outlook_pdf_tool/tests
```
