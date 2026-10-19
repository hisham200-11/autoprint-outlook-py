import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date, timedelta
from typing import Dict, List, Optional

# Core modules (local package imports)
from storage_manager import StorageManager, AttachmentRecord
from outlook_adapter import OutlookAdapter
from print_spooler import SequentialPrintEngine, PrintJob
import updater

# UI-UX Design Tokens: Impeccable Clean White, Neutral Blacks, Vibrant High-Contrast Orange
COLOR_BG_WHITE = "#FFFFFF"
COLOR_BG_LIGHT = "#F4F5F7"
COLOR_CARD_BORDER = "#E2E8F0"
COLOR_TEXT_MAIN = "#0F172A"       # Slate 900 (Deep, high readability)
COLOR_TEXT_SECONDARY = "#475569"  # Slate 600
COLOR_TEXT_MUTED = "#64748B"      # Slate 500
COLOR_ORANGE = "#F97316"          # Energetic modern amber/orange
COLOR_ORANGE_HOVER = "#EA580C"    # Darker hover orange
COLOR_ORANGE_LIGHT = "#FFF7ED"    # Soft orange selection tint
COLOR_BORDER = "#CBD5E1"          # Slate 300
COLOR_GREEN = "#16A34A"           # Success emerald green
COLOR_GREEN_BG = "#DCFCE7"        # Soft emerald tint

FONT_FAMILY = "Segoe UI"
FONT_TITLE = (FONT_FAMILY, 15, "bold")
FONT_HEADING = (FONT_FAMILY, 10, "bold")
FONT_BODY = (FONT_FAMILY, 9)
FONT_BODY_BOLD = (FONT_FAMILY, 9, "bold")
FONT_SMALL = (FONT_FAMILY, 8)

class OutlookPDFPrinterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AutoPrint Outlook — Bulk Attachment Extractor & Spooler")
        self.geometry("1180x760")
        self.minsize(1000, 620)
        self.configure(bg=COLOR_BG_LIGHT)

        # Capabilities
        self.storage_manager = StorageManager()
        self.outlook_adapter = OutlookAdapter()
        self.print_engine = SequentialPrintEngine(delay_seconds=2.0)

        # App state
        self.attachments_data: List[dict] = []
        self.item_checkboxes: Dict[str, tk.BooleanVar] = {}
        self.current_selected_timeline = "all"
        self.current_type_filter = "All"

        # Wire print engine callbacks
        self.print_engine.set_status_callback(self._on_print_status_update)
        self.print_engine.set_job_done_callback(self._on_print_job_done)

        # Style configuration
        self._setup_styles()
        
        # Build UI layout with sticky top and bottom anchors so Full Screen never obscures buttons!
        self._build_header()
        self._build_filter_bar()
        self._build_bottom_bar()   # Docked to bottom FIRST (pack side="bottom")
        self._build_table_area()   # Expands to fill remaining middle space (pack fill="both", expand=True)

        # Initial load in background (populate folders and scan)
        self.after(300, self._load_initial_data)

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Treeview styling with professional optical padding and row height
        style.configure(
            "Custom.Treeview",
            background=COLOR_BG_WHITE,
            foreground=COLOR_TEXT_MAIN,
            fieldbackground=COLOR_BG_WHITE,
            rowheight=32,
            font=FONT_BODY,
            borderwidth=0
        )
        style.configure(
            "Custom.Treeview.Heading",
            background=COLOR_BG_LIGHT,
            foreground=COLOR_TEXT_MAIN,
            font=FONT_HEADING,
            relief="flat",
            padding=(8, 8)
        )
        style.map(
            "Custom.Treeview",
            background=[("selected", COLOR_ORANGE_LIGHT)],
            foreground=[("selected", COLOR_TEXT_MAIN)]
        )

    def _build_header(self):
        header_frame = tk.Frame(self, bg=COLOR_BG_WHITE, bd=1, relief="flat", highlightthickness=1, highlightbackground=COLOR_CARD_BORDER)
        header_frame.pack(side="top", fill="x", padx=16, pady=(12, 6))

        # Title + Subtitle
        title_box = tk.Frame(header_frame, bg=COLOR_BG_WHITE)
        title_box.pack(side="left", padx=20, pady=10)

        lbl_title = tk.Label(
            title_box,
            text="Outlook Attachment Extractor & Bulk Printer",
            font=FONT_TITLE,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_WHITE
        )
        lbl_title.pack(anchor="w")

        lbl_sub = tk.Label(
            title_box,
            text="Extract PDFs, Images (PNG/JPG), Word, and Excel docs. Select and print sequentially without spooler jams.",
            font=FONT_SMALL,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_WHITE
        )
        lbl_sub.pack(anchor="w")

        # Action Buttons on Header
        actions_box = tk.Frame(header_frame, bg=COLOR_BG_WHITE)
        actions_box.pack(side="right", padx=20, pady=12)

        btn_update = tk.Button(
            actions_box,
            text="🔄 Check for Updates",
            font=FONT_SMALL,
            bg=COLOR_BG_WHITE,
            fg=COLOR_TEXT_MAIN,
            activebackground=COLOR_BORDER,
            bd=1,
            relief="solid",
            padx=10,
            pady=6,
            cursor="hand2",
            command=self._start_update_check
        )
        btn_update.pack(side="left", padx=(0, 10))

        btn_cleanup = tk.Button(
            actions_box,
            text="🗑 Clean Up Saved Files",
            font=FONT_SMALL,
            bg=COLOR_BG_WHITE,
            fg="#DC2626",
            activebackground="#FEE2E2",
            bd=1,
            relief="solid",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._prompt_clean_saved_files
        )
        btn_cleanup.pack(side="left", padx=(0, 10))

        btn_archive = tk.Button(
            actions_box,
            text="📂 Open Archive",
            font=FONT_BODY_BOLD,
            bg=COLOR_BG_LIGHT,
            fg=COLOR_TEXT_MAIN,
            activebackground=COLOR_BORDER,
            bd=1,
            relief="solid",
            padx=14,
            pady=6,
            cursor="hand2",
            command=self._open_archive_folder
        )
        btn_archive.pack(side="left")

    def _build_filter_bar(self):
        filter_card = tk.Frame(self, bg=COLOR_BG_WHITE, bd=1, relief="flat", highlightthickness=1, highlightbackground=COLOR_CARD_BORDER)
        filter_card.pack(fill="x", padx=15, pady=6)

        inner = tk.Frame(filter_card, bg=COLOR_BG_WHITE)
        inner.pack(fill="x", padx=20, pady=12)

        # 1. Folder Selector Dropdown
        lbl_folder = tk.Label(inner, text="Source Folder:", font=FONT_HEADING, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_WHITE)
        lbl_folder.grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)

        self.combo_folders = ttk.Combobox(inner, font=FONT_BODY, state="readonly", width=25)
        self.combo_folders.grid(row=0, column=1, sticky="w", padx=(0, 20), pady=4)
        self.combo_folders.set("Default Inbox")

        # 2. Timeline Buttons
        lbl_time = tk.Label(inner, text="Timeline:", font=FONT_HEADING, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_WHITE)
        lbl_time.grid(row=0, column=2, sticky="w", padx=(0, 8), pady=4)

        self.timeline_btn_frame = tk.Frame(inner, bg=COLOR_BG_WHITE)
        self.timeline_btn_frame.grid(row=0, column=3, sticky="w", padx=(0, 20), pady=4)

        self.timeline_buttons = {}
        for mode, label in [("today", "Today"), ("yesterday", "Yesterday"), ("last7", "Last 7 Days"), ("all", "All")]:
            btn = tk.Button(
                self.timeline_btn_frame,
                text=label,
                font=FONT_SMALL,
                bd=1,
                relief="solid",
                padx=10,
                pady=4,
                cursor="hand2",
                command=lambda m=mode: self._set_timeline_mode(m)
            )
            btn.pack(side="left", padx=2)
            self.timeline_buttons[mode] = btn

        self._update_timeline_button_styles()

        # 3. File Type Filters
        lbl_type = tk.Label(inner, text="Type:", font=FONT_HEADING, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_WHITE)
        lbl_type.grid(row=0, column=4, sticky="w", padx=(0, 8), pady=4)

        self.type_btn_frame = tk.Frame(inner, bg=COLOR_BG_WHITE)
        self.type_btn_frame.grid(row=0, column=5, sticky="w", padx=(0, 20), pady=4)

        self.type_buttons = {}
        for t_filter in ["All", "PDF", "Image", "Word", "Excel"]:
            btn = tk.Button(
                self.type_btn_frame,
                text=t_filter,
                font=FONT_SMALL,
                bd=1,
                relief="solid",
                padx=8,
                pady=4,
                cursor="hand2",
                command=lambda tf=t_filter: self._set_type_filter(tf)
            )
            btn.pack(side="left", padx=2)
            self.type_buttons[t_filter] = btn

        self._update_type_button_styles()

        # 4. Big Scan Button (Orange Accent)
        self.btn_scan = tk.Button(
            inner,
            text="⚡ Scan Attachments",
            font=FONT_HEADING,
            bg=COLOR_ORANGE,
            fg=COLOR_BG_WHITE,
            activebackground=COLOR_ORANGE_HOVER,
            activeforeground=COLOR_BG_WHITE,
            bd=0,
            padx=18,
            pady=7,
            cursor="hand2",
            command=self._start_scan_thread
        )
        self.btn_scan.grid(row=0, column=6, sticky="e", padx=(10, 0), pady=4)

        inner.columnconfigure(6, weight=1)

    def _build_table_area(self):
        container = tk.Frame(self, bg=COLOR_BG_WHITE, bd=1, relief="flat", highlightthickness=1, highlightbackground=COLOR_CARD_BORDER)
        # Pack to fill all remaining vertical and horizontal space between header and bottom bar
        container.pack(side="top", fill="both", expand=True, padx=16, pady=6)

        # Action bar above table (Select All, Deselect All, Stats)
        top_table_bar = tk.Frame(container, bg=COLOR_BG_WHITE)
        top_table_bar.pack(fill="x", padx=16, pady=10)

        btn_select_all = tk.Button(
            top_table_bar,
            text="✓ Select All",
            font=FONT_SMALL,
            bg=COLOR_BG_WHITE,
            fg=COLOR_TEXT_MAIN,
            bd=1,
            relief="solid",
            padx=8,
            pady=3,
            cursor="hand2",
            command=self._select_all_unprinted
        )
        btn_select_all.pack(side="left", padx=(0, 6))

        btn_deselect = tk.Button(
            top_table_bar,
            text="✕ Deselect All",
            font=FONT_SMALL,
            bg=COLOR_BG_WHITE,
            fg=COLOR_TEXT_MAIN,
            bd=1,
            relief="solid",
            padx=8,
            pady=3,
            cursor="hand2",
            command=self._deselect_all
        )
        btn_deselect.pack(side="left", padx=(0, 15))

        self.lbl_table_summary = tk.Label(
            top_table_bar,
            text="0 attachments found • 0 selected",
            font=FONT_BODY_BOLD,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_WHITE
        )
        self.lbl_table_summary.pack(side="left")

        # Treeview Table
        columns = ("select", "status", "type", "sender", "filename", "pages", "received", "subject")
        self.tree = ttk.Treeview(
            container,
            columns=columns,
            show="headings",
            style="Custom.Treeview",
            selectmode="extended"
        )

        self.tree.heading("select", text="Print?")
        self.tree.heading("status", text="Status")
        self.tree.heading("type", text="Type")
        self.tree.heading("sender", text="Sender (Source)")
        self.tree.heading("filename", text="Attachment Name")
        self.tree.heading("pages", text="Pages")
        self.tree.heading("received", text="Date Received")
        self.tree.heading("subject", text="Email Subject")

        self.tree.column("select", width=60, anchor="center")
        self.tree.column("status", width=110, anchor="center")
        self.tree.column("type", width=80, anchor="center")
        self.tree.column("sender", width=210, anchor="w")
        self.tree.column("filename", width=220, anchor="w")
        self.tree.column("pages", width=55, anchor="center")
        self.tree.column("received", width=140, anchor="center")
        self.tree.column("subject", width=240, anchor="w")

        # Scrollbar
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=(15, 0), pady=(0, 15))
        scroll.pack(side="right", fill="y", padx=(0, 15), pady=(0, 15))

        # Click event to toggle checkbox
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

    def _build_bottom_bar(self):
        bottom_card = tk.Frame(self, bg=COLOR_BG_WHITE, bd=1, relief="flat", highlightthickness=1, highlightbackground=COLOR_CARD_BORDER)
        # Dock strictly to the bottom of the window so it is ALWAYS visible in full screen
        bottom_card.pack(side="bottom", fill="x", padx=16, pady=(6, 14))

        inner = tk.Frame(bottom_card, bg=COLOR_BG_WHITE)
        inner.pack(fill="x", padx=20, pady=12)

        # Progress and printer status
        info_frame = tk.Frame(inner, bg=COLOR_BG_WHITE)
        info_frame.pack(side="left", fill="x", expand=True)

        self.lbl_printer_name = tk.Label(
            info_frame,
            text=f"Default Printer: {self.print_engine.get_default_printer_name() or 'Detecting...'}",
            font=FONT_SMALL,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_WHITE
        )
        self.lbl_printer_name.pack(anchor="w")

        self.lbl_status = tk.Label(
            info_frame,
            text="Ready. Select documents above and click Print.",
            font=FONT_HEADING,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_WHITE
        )
        self.lbl_status.pack(anchor="w", pady=(2, 0))

        # Action Buttons (Right aligned)
        btn_frame = tk.Frame(inner, bg=COLOR_BG_WHITE)
        btn_frame.pack(side="right")

        self.btn_cancel = tk.Button(
            btn_frame,
            text="⏹ Cancel Queue",
            font=FONT_BODY_BOLD,
            bg=COLOR_BG_WHITE,
            fg="#DC2626",
            activebackground="#FEE2E2",
            bd=1,
            relief="solid",
            padx=16,
            pady=8,
            cursor="hand2",
            state="disabled",
            command=self._cancel_print_queue
        )
        self.btn_cancel.pack(side="left", padx=(0, 12))

        self.btn_print = tk.Button(
            btn_frame,
            text="🖨 Print Selected Documents",
            font=FONT_HEADING,
            bg=COLOR_ORANGE,
            fg=COLOR_BG_WHITE,
            activebackground=COLOR_ORANGE_HOVER,
            activeforeground=COLOR_BG_WHITE,
            bd=0,
            padx=22,
            pady=9,
            cursor="hand2",
            command=self._start_print_queue
        )
        self.btn_print.pack(side="left")

    def _set_type_filter(self, t_filter: str):
        self.current_type_filter = t_filter
        self._update_type_button_styles()
        self._populate_table()

    def _update_type_button_styles(self):
        for t_filter, btn in self.type_buttons.items():
            if t_filter == self.current_type_filter:
                btn.configure(bg=COLOR_ORANGE, fg=COLOR_BG_WHITE, relief="flat")
            else:
                btn.configure(bg=COLOR_BG_WHITE, fg=COLOR_TEXT_MAIN, relief="solid")

    def _set_timeline_mode(self, mode: str):
        self.current_selected_timeline = mode
        self._update_timeline_button_styles()
        self._start_scan_thread()

    def _update_timeline_button_styles(self):
        for mode, btn in self.timeline_buttons.items():
            if mode == self.current_selected_timeline:
                btn.configure(bg=COLOR_ORANGE, fg=COLOR_BG_WHITE, relief="flat")
            else:
                btn.configure(bg=COLOR_BG_WHITE, fg=COLOR_TEXT_MAIN, relief="solid")

    def _get_timeline_dates(self):
        today = date.today()
        if self.current_selected_timeline == "today":
            return today, today
        elif self.current_selected_timeline == "yesterday":
            y = today - timedelta(days=1)
            return y, y
        elif self.current_selected_timeline == "last7":
            return today - timedelta(days=7), today
        else:
            return None, None

    def _load_initial_data(self):
        # Populate folders
        folders = self.outlook_adapter.get_available_folders()
        self.folder_map = {name: fid for name, fid in folders}
        self.combo_folders["values"] = list(self.folder_map.keys())
        if folders:
            self.combo_folders.current(0)
        # Automatically run first scan
        self._start_scan_thread()

    def _start_scan_thread(self):
        self.btn_scan.configure(state="disabled", text="⏳ Scanning...")
        self.lbl_status.configure(text="Connecting to Outlook and scanning emails...", fg=COLOR_ORANGE)
        
        thread = threading.Thread(target=self._run_scan_worker, daemon=True)
        thread.start()

    def _run_scan_worker(self):
        selected_folder_name = self.combo_folders.get()
        folder_id = self.folder_map.get(selected_folder_name, "inbox")
        start_d, end_d = self._get_timeline_dates()

        def on_progress(text, pct):
            self.after(0, lambda: self.lbl_status.configure(text=text))

        try:
            results = self.outlook_adapter.scan_attachments(
                folder_id=folder_id,
                start_date=start_d,
                end_date=end_d,
                storage_manager=self.storage_manager,
                progress_callback=on_progress
            )
            self.after(0, lambda: self._on_scan_completed(results))
        except Exception as e:
            self.after(0, lambda: self._on_scan_failed(str(e)))

    def _on_scan_completed(self, results: List[dict]):
        self.attachments_data = results
        self.btn_scan.configure(state="normal", text="⚡ Scan Outlook Attachments")
        self._populate_table()
        self.lbl_status.configure(text=f"Scan complete. {len(results)} PDF attachments ready.", fg=COLOR_GREEN)

    def _on_scan_failed(self, error_msg: str):
        self.btn_scan.configure(state="normal", text="⚡ Scan Outlook Attachments")
        self.lbl_status.configure(text=f"Scan error: {error_msg}", fg="#DC3545")
        messagebox.showerror("Outlook Error", f"Unable to scan Outlook:\n\n{error_msg}")

    def _populate_table(self):
        # Clear existing
        for item in self.tree.get_children():
            self.tree.delete(item)

        unprinted_count = 0
        for att in self.attachments_data:
            category = att.get("category", "PDF")
            # Apply type filter if not "All"
            if self.current_type_filter != "All" and category != self.current_type_filter:
                continue

            file_id = att["file_id"]
            is_printed = att["is_printed"]
            
            # Preserve existing checkbox if already set, else default to not is_printed
            if file_id not in self.item_checkboxes:
                self.item_checkboxes[file_id] = tk.BooleanVar(value=not is_printed)
            chk_var = self.item_checkboxes[file_id]

            if not is_printed:
                unprinted_count += 1
                status_display = "⏳ Pending"
            else:
                status_display = "✓ Printed"

            chk_symbol = "☑" if chk_var.get() else "☐"

            self.tree.insert(
                "",
                "end",
                iid=file_id,
                values=(
                    chk_symbol,
                    status_display,
                    category,
                    att["sender"],
                    att["filename"],
                    att["page_count"],
                    att["received_time"],
                    att["subject"]
                )
            )

        self._update_summary_label()

    def _on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            item_id = self.tree.identify_row(event.y)
            if not item_id:
                return

            # If user clicked the first column or anywhere on the row, toggle checkbox
            if column in ("#1", "#2", "#3"):
                self._toggle_item(item_id)

    def _toggle_item(self, item_id: str):
        if item_id in self.item_checkboxes:
            current = self.item_checkboxes[item_id].get()
            new_val = not current
            self.item_checkboxes[item_id].set(new_val)
            
            # Update row symbol
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = "☑" if new_val else "☐"
            self.tree.item(item_id, values=vals)
            self._update_summary_label()

    def _select_all_unprinted(self):
        for att in self.attachments_data:
            fid = att["file_id"]
            category = att.get("category", "PDF")
            if self.current_type_filter != "All" and category != self.current_type_filter:
                continue

            if not att["is_printed"] and fid in self.item_checkboxes:
                self.item_checkboxes[fid].set(True)
                if self.tree.exists(fid):
                    vals = list(self.tree.item(fid, "values"))
                    vals[0] = "☑"
                    self.tree.item(fid, values=vals)
        self._update_summary_label()

    def _deselect_all(self):
        for fid, var in self.item_checkboxes.items():
            var.set(False)
            if self.tree.exists(fid):
                vals = list(self.tree.item(fid, "values"))
                vals[0] = "☐"
                self.tree.item(fid, values=vals)
        self._update_summary_label()

    def _update_summary_label(self):
        total = len(self.attachments_data)
        selected = sum(1 for v in self.item_checkboxes.values() if v.get())
        total_pages = sum(att["page_count"] for att in self.attachments_data if self.item_checkboxes.get(att["file_id"], tk.BooleanVar(value=False)).get())
        self.lbl_table_summary.configure(
            text=f"{total} total attachments ({selected} selected to print • {total_pages} pages)"
        )

    def _start_print_queue(self):
        selected_jobs: List[PrintJob] = []
        for att in self.attachments_data:
            fid = att["file_id"]
            if self.item_checkboxes.get(fid, tk.BooleanVar(value=False)).get():
                job = PrintJob(
                    file_id=fid,
                    file_path=att["local_path"],
                    sender=att["sender"],
                    filename=att["filename"],
                    page_count=att["page_count"],
                    category=att.get("category", "PDF")
                )
                selected_jobs.append(job)

        if not selected_jobs:
            messagebox.showinfo("No Selection", "Please check at least one attachment to print.")
            return

        confirm = messagebox.askyesno(
            "Confirm Bulk Print",
            f"You are about to print {len(selected_jobs)} documents sequentially to:\n\n"
            f"Printer: {self.print_engine.get_default_printer_name()}\n\n"
            f"The spooler will process them 1-by-1 with safety pacing. Proceed?"
        )
        if not confirm:
            return

        # Disable UI buttons during print
        self.btn_print.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.btn_scan.configure(state="disabled")

        self.print_engine.enqueue_jobs(selected_jobs)
        self.print_engine.start_worker()

    def _cancel_print_queue(self):
        self.print_engine.cancel_queue()
        self.lbl_status.configure(text="Cancelling queue... waiting for active job to finish.", fg="#DC3545")

    def _on_print_status_update(self, status: str, job: Optional[PrintJob], current: int, total: int):
        self.after(0, lambda: self._handle_print_status_ui(status, job, current, total))

    def _handle_print_status_ui(self, status: str, job: Optional[PrintJob], current: int, total: int):
        if status == "spooling" and job:
            self.lbl_status.configure(
                text=f"🖨 Printing ({current}/{total}): '{job.filename}' from {job.sender}...",
                fg=COLOR_ORANGE
            )
        elif status == "complete":
            self.lbl_status.configure(text="✓ All selected documents printed and archived successfully!", fg=COLOR_GREEN)
            self.btn_print.configure(state="normal")
            self.btn_cancel.configure(state="disabled")
            self.btn_scan.configure(state="normal")
            messagebox.showinfo("Finished", "All selected documents have been sent to the printer and moved to the archive.")
        elif status == "cancelled":
            self.lbl_status.configure(text="Queue cancelled by user.", fg="#DC3545")
            self.btn_print.configure(state="normal")
            self.btn_cancel.configure(state="disabled")
            self.btn_scan.configure(state="normal")

    def _on_print_job_done(self, job: PrintJob):
        # Called when 1 PDF completes spooling: move to archive & mark UI
        self.after(0, lambda: self._handle_job_done_ui(job))

    def _handle_job_done_ui(self, job: PrintJob):
        # 1. Update storage record
        rec = AttachmentRecord(
            file_id=job.file_id,
            message_id="msg",
            sender=job.sender,
            subject="",
            received_time="",
            filename=job.filename,
            local_path=job.file_path,
            page_count=job.page_count
        )
        self.storage_manager.archive_printed_file(rec)

        # 2. Update memory state
        for att in self.attachments_data:
            if att["file_id"] == job.file_id:
                att["is_printed"] = True
                break

        # 3. Update table item
        if self.tree.exists(job.file_id):
            vals = list(self.tree.item(job.file_id, "values"))
            vals[0] = "☐" # Uncheck
            vals[1] = "✓ Printed" # Change status badge
            self.tree.item(job.file_id, values=vals)
            if job.file_id in self.item_checkboxes:
                self.item_checkboxes[job.file_id].set(False)

        self._update_summary_label()

    def _open_archive_folder(self):
        archive_path = self.storage_manager.archive_dir
        os.makedirs(archive_path, exist_ok=True)
        os.startfile(archive_path)

    def _prompt_clean_saved_files(self):
        pending_count = len(os.listdir(self.storage_manager.pending_dir)) if os.path.exists(self.storage_manager.pending_dir) else 0

        # Ask confirmation with choice
        dialog = messagebox.askyesnocancel(
            "Clean Up Saved Files",
            f"You currently have files extracted in:\n"
            f"📁 {self.storage_manager.pending_dir}\n\n"
            f"• Click YES: Delete temporary/staging files in 'Downloads/Pending' ({pending_count} items) to free space.\n"
            f"• Click NO: Delete BOTH staging files AND previous 'Printed_Archive' files.\n"
            f"• Click CANCEL: Keep everything.\n"
        )

        if dialog is True:
            # Delete only pending staging files
            deleted = self.storage_manager.clear_pending_files()
            self.lbl_status.configure(text=f"Cleaned up {deleted} temporary staging files.", fg=COLOR_GREEN)
            messagebox.showinfo("Cleanup Complete", f"Successfully deleted {deleted} temporary staging files from 'Downloads/Pending'.")
            self._start_scan_thread()
        elif dialog is False:
            # Delete both pending and archive
            del_pending = self.storage_manager.clear_pending_files()
            del_archive = self.storage_manager.clear_archive_files()
            total_del = del_pending + del_archive
            self.lbl_status.configure(text=f"Cleaned up {total_del} saved files (staging & archive).", fg=COLOR_GREEN)
            messagebox.showinfo("Cleanup Complete", f"Successfully deleted {del_pending} staging files and {del_archive} archived files.")
            self._start_scan_thread()

    def _start_update_check(self):
        self.lbl_status.configure(text="Checking GitHub for updates...", fg=COLOR_ORANGE)
        threading.Thread(target=self._run_update_check_worker, daemon=True).start()

    def _run_update_check_worker(self):
        res = updater.check_for_updates()
        self.after(0, lambda: self._handle_update_check_result(res))

    def _handle_update_check_result(self, res: dict):
        if not res.get("success"):
            self.lbl_status.configure(text=f"Update check failed: {res.get('error', 'Network error')}", fg="#DC2626")
            messagebox.showwarning(
                "Update Check",
                f"Could not check for updates:\n{res.get('error', 'Network error')}\n\nPlease check your internet connection."
            )
            return

        if not res.get("update_available"):
            self.lbl_status.configure(text=f"AutoPrint is up to date (Version {res.get('current_version')}).", fg=COLOR_GREEN)
            messagebox.showinfo(
                "Up to Date",
                f"You are running the latest version: v{res.get('current_version')}\n\nNo updates are needed!"
            )
            return

        # An update is available!
        latest = res.get("latest_version")
        notes = res.get("release_notes", "")
        download_url = res.get("download_url")

        confirm = messagebox.askyesno(
            "New Update Available! 🚀",
            f"A new version of AutoPrint Outlook is available!\n\n"
            f"• Current version: v{res.get('current_version')}\n"
            f"• New version: {latest}\n\n"
            f"What's New:\n{notes[:400]}\n\n"
            f"Would you like to download and update now?\n"
            f"(Note: All your saved history, settings, and printed archives will be 100% preserved)."
        )

        if not confirm:
            self.lbl_status.configure(text="Update postponed by user.", fg=COLOR_TEXT_MUTED)
            return

        if not download_url:
            messagebox.showerror(
                "Download Error",
                f"Release {latest} was found on GitHub, but no standalone zip package was attached.\n"
                f"Please download it manually from:\nhttps://github.com/{updater.GITHUB_REPO}/releases"
            )
            return

        self._start_download_and_apply_update(download_url)

    def _start_download_and_apply_update(self, download_url: str):
        self.lbl_status.configure(text="Downloading update package from GitHub...", fg=COLOR_ORANGE)
        
        def download_worker():
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                zip_path = updater.download_and_prepare_update(download_url, base_dir)
                self.after(0, lambda: self._apply_update_and_restart(base_dir, zip_path))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Update Error", f"Failed to download update: {e}"))
                self.after(0, lambda: self.lbl_status.configure(text=f"Update failed: {e}", fg="#DC2626"))

        threading.Thread(target=download_worker, daemon=True).start()

    def _apply_update_and_restart(self, base_dir: str, zip_path: str):
        messagebox.showinfo(
            "Applying Update",
            "Update downloaded successfully!\n\n"
            "The application will now close, update program files safely, and restart automatically."
        )
        updater.launch_safe_updater_script(base_dir, zip_path)

if __name__ == "__main__":
    app = OutlookPDFPrinterApp()
    app.mainloop()
