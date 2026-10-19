import os
import sys
import time
import queue
import threading
from dataclasses import dataclass
from typing import List, Optional, Callable

@dataclass
class PrintJob:
    file_id: str
    file_path: str
    sender: str
    filename: str
    page_count: int
    category: str = "PDF"

class SequentialPrintEngine:
    def __init__(self, delay_seconds: float = 2.0, max_spool_wait: float = 15.0):
        self.delay_seconds = delay_seconds
        self.max_spool_wait = max_spool_wait
        self._job_queue = queue.Queue()
        self._is_running = False
        self._should_cancel = False
        self._worker_thread: Optional[threading.Thread] = None
        self._status_callback: Optional[Callable[[str, PrintJob, int, int], None]] = None
        self._job_done_callback: Optional[Callable[[PrintJob], None]] = None

    def set_status_callback(self, cb: Callable[[str, PrintJob, int, int], None]):
        self._status_callback = cb

    def set_job_done_callback(self, cb: Callable[[PrintJob], None]):
        self._job_done_callback = cb

    def enqueue_jobs(self, jobs: List[PrintJob]):
        for job in jobs:
            self._job_queue.put(job)

    def get_queue_length(self) -> int:
        return self._job_queue.qsize()

    def is_running(self) -> bool:
        return self._is_running

    def cancel_queue(self):
        self._should_cancel = True
        # Drain the queue
        while not self._job_queue.empty():
            try:
                self._job_queue.get_nowait()
                self._job_queue.task_done()
            except queue.Empty:
                break

    def get_default_printer_name(self) -> Optional[str]:
        try:
            import win32print
            return win32print.GetDefaultPrinter()
        except Exception:
            return None

    def _get_spooler_job_count(self, printer_name: str) -> int:
        try:
            import win32print
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                # EnumJobs level 1 returns active print jobs
                jobs = win32print.EnumJobs(hprinter, 0, 99, 1)
                return len(jobs)
            finally:
                win32print.ClosePrinter(hprinter)
        except Exception:
            return 0

    def start_worker(self):
        if self._is_running:
            return
        self._should_cancel = False
        self._is_running = True
        self._worker_thread = threading.Thread(target=self._run_loop, daemon=True)
        self._worker_thread.start()

    def _run_loop(self):
        # COM apartment threading requirement for worker thread
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        total_jobs = self._job_queue.qsize()
        processed = 0

        default_printer = self.get_default_printer_name()

        while not self._job_queue.empty() and not self._should_cancel:
            job: PrintJob = self._job_queue.get()
            processed += 1

            if self._status_callback:
                self._status_callback("spooling", job, processed, total_jobs)

            # Check spooler condition
            if default_printer:
                start_wait = time.time()
                while self._get_spooler_job_count(default_printer) > 2:
                    if self._should_cancel or (time.time() - start_wait) > self.max_spool_wait:
                        break
                    time.sleep(0.5)

            # Dispatch print
            try:
                self._dispatch_print(job.file_path, job.category, default_printer)
            except Exception as e:
                if self._status_callback:
                    self._status_callback(f"error: {e}", job, processed, total_jobs)

            # Wait anti-jam delay
            time.sleep(self.delay_seconds)

            # Notify job completion
            if self._job_done_callback and not self._should_cancel:
                self._job_done_callback(job)

            if self._status_callback:
                self._status_callback("finished_item", job, processed, total_jobs)

            self._job_queue.task_done()

        self._is_running = False
        if self._status_callback:
            status_text = "cancelled" if self._should_cancel else "complete"
            self._status_callback(status_text, None, processed, total_jobs)

    def _dispatch_print(self, file_path: str, category: str, printer_name: Optional[str]):
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        # Specialized handling for Word and Excel to guarantee silent background printing
        if category == "Word":
            self._print_word_document(abs_path)
            return
        elif category == "Excel":
            self._print_excel_document(abs_path)
            return

        # PDF and Image printing via Windows ShellExecute (Silent)
        try:
            import win32api
            win32api.ShellExecute(0, "print", abs_path, None, ".", 0)
        except Exception as e:
            try:
                os.startfile(abs_path, "print")
            except Exception as fallback_err:
                raise RuntimeError(f"Print dispatch failed: {fallback_err}")

    def _print_word_document(self, file_path: str):
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(file_path, ReadOnly=True, AddToRecentFiles=False)
            doc.PrintOut(Background=False)
            doc.Close(False)
        finally:
            word.Quit()

    def _print_excel_document(self, file_path: str):
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(file_path, ReadOnly=True, UpdateLinks=False)
            wb.PrintOut()
            wb.Close(False)
        finally:
            excel.Quit()
