import logging
import os
import queue
import subprocess
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from PIL import Image, ImageTk

from src.io.excel_writer import write_excel
from src.io.file_utils import get_image_files, get_targets
from src.qrcode.scanner import QRScanner


APP_TITLE = "Pro QR Extractor"
LOGO_PATH = Path(__file__).resolve().parent / "logo.png"


class QueueLogHandler(logging.Handler):
    """Send formatted log lines to the Tkinter event queue."""

    def __init__(self, event_queue: queue.Queue):
        super().__init__()
        self.event_queue = event_queue

    def emit(self, record: logging.LogRecord) -> None:
        try:
            message = self.format(record)
        except Exception:
            message = record.getMessage()
        self.event_queue.put(("log", message))


class QRExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1240x760")
        self.root.minsize(1024, 680)

        self.event_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.extracted_data: list[dict] = []
        self.processed_count = 0
        self._logo_image: ImageTk.PhotoImage | None = None

        self.root_dir_var = tk.StringVar()

        self.start_button: ttk.Button | None = None
        self.export_button: ttk.Button | None = None
        self.progress_bar: ttk.Progressbar | None = None
        self.tree: ttk.Treeview | None = None
        self.log_text: scrolledtext.ScrolledText | None = None

        self._build_ui()
        self.root.after(100, self._process_event_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(0, weight=1)

        self._build_left_panel(container)
        self._build_right_panel(container)

    def _build_left_panel(self, parent: ttk.Frame) -> None:
        left_frame = ttk.Frame(parent, padding=(12, 12, 16, 12), width=220)
        left_frame.grid(row=0, column=0, sticky="nsw")
        left_frame.grid_propagate(False)
        left_frame.columnconfigure(0, weight=1)

        logo_label = ttk.Label(left_frame)
        logo_label.grid(row=0, column=0, pady=(0, 16))
        self._load_logo(logo_label)

        ttk.Label(
            left_frame,
            text="Thông tin hệ thống",
            font=("TkDefaultFont", 12, "bold"),
        ).grid(row=1, column=0, sticky="w", pady=(0, 12))

        info_items = [
            ("Tên ứng dụng:", "Pro QR Extractor"),
            ("Mô tả:", "Phần mềm đọc mã QR từ ảnh thiết bị"),
            ("Phiên bản:", "1.0.0"),
            ("Ngày phát hành:", "16/04/2026"),
        ]
        current_row = 2
        for title, value in info_items:
            ttk.Label(left_frame, text=title, font=("TkDefaultFont", 10, "bold")).grid(
                row=current_row, column=0, sticky="w", pady=(6, 0)
            )
            current_row += 1
            ttk.Label(left_frame, text=value, wraplength=180, justify="left").grid(
                row=current_row, column=0, sticky="w", pady=(0, 6)
            )
            current_row += 1

    def _build_right_panel(self, parent: ttk.Frame) -> None:
        right_frame = ttk.Frame(parent, padding=(8, 8, 8, 8))
        right_frame.grid(row=0, column=1, sticky="nsew")
        right_frame.columnconfigure(1, weight=1)
        right_frame.rowconfigure(4, weight=1)
        right_frame.rowconfigure(6, weight=0)

        ttk.Label(right_frame, text="Thư mục chứa ảnh:").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        ttk.Entry(right_frame, textvariable=self.root_dir_var).grid(
            row=0, column=1, sticky="ew", padx=(12, 8), pady=(0, 8)
        )
        ttk.Button(right_frame, text="Browse", command=self._browse_root_folder).grid(
            row=0, column=2, sticky="ew", pady=(0, 8)
        )

        action_frame = ttk.Frame(right_frame)
        action_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 10))

        self.start_button = ttk.Button(
            action_frame, text="Bắt đầu quét", command=self._start_scan
        )
        self.start_button.grid(row=0, column=0, padx=(0, 8))

        self.export_button = ttk.Button(
            action_frame,
            text="Xuất Excel",
            command=self._export_excel,
            state="disabled",
        )
        self.export_button.grid(row=0, column=1)

        self.progress_bar = ttk.Progressbar(right_frame, mode="indeterminate")
        self.progress_bar.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        self.progress_bar.grid_remove()

        table_frame = ttk.Frame(right_frame)
        table_frame.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(0, 12))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("stt", "folder", "file_name", "qr_content", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        headings = {
            "stt": "STT",
            "folder": "Folder",
            "file_name": "File name",
            "qr_content": "QR content",
            "status": "Status",
        }
        widths = {
            "stt": 40,
            "folder": 150,
            "file_name": 200,
            "qr_content": 300,
            "status": 80,
        }
        for column_id in columns:
            self.tree.heading(column_id, text=headings[column_id])
            self.tree.column(
                column_id,
                width=widths[column_id],
                anchor="w",
                stretch=column_id != "stt",
            )

        tree_vsb = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree.yview
        )
        tree_hsb = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_vsb.grid(row=0, column=1, sticky="ns")
        tree_hsb.grid(row=1, column=0, sticky="ew")

        ttk.Label(
            right_frame,
            text="System Logs / Trạng thái xử lý",
            font=("TkDefaultFont", 11, "bold"),
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 6))

        self.log_text = scrolledtext.ScrolledText(
            right_frame,
            height=10,
            wrap="word",
            background="#1e1e1e",
            foreground="#d4d4d4",
            insertbackground="#d4d4d4",
        )
        self.log_text.grid(row=6, column=0, columnspan=3, sticky="nsew")
        self.log_text.configure(state="disabled")

    def _load_logo(self, label: ttk.Label) -> None:
        if not LOGO_PATH.exists():
            label.configure(text="logo.png")
            return

        try:
            image = Image.open(LOGO_PATH)
            image.thumbnail((160, 160))
            self._logo_image = ImageTk.PhotoImage(image)
            label.configure(image=self._logo_image)
        except Exception:
            label.configure(text="logo.png")

    def _browse_root_folder(self) -> None:
        directory = filedialog.askdirectory()
        if directory:
            self.root_dir_var.set(directory)

    def _start_scan(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return

        root_input = self.root_dir_var.get().strip()
        if not root_input:
            messagebox.showerror("Lỗi", "Vui lòng chọn thư mục chứa ảnh.")
            return

        root_dir = Path(root_input).expanduser()
        if not root_dir.exists() or not root_dir.is_dir():
            messagebox.showerror("Lỗi", f"Thư mục không hợp lệ:\n{root_dir}")
            return

        targets, layout = get_targets(root_dir)
        if layout == "empty":
            messagebox.showerror(
                "Lỗi",
                "Không tìm thấy thư mục con hoặc file ảnh hợp lệ trong thư mục đã chọn.",
            )
            return

        total_images = sum(len(get_image_files(target)) for target in targets)
        if total_images == 0:
            messagebox.showerror("Lỗi", "Không tìm thấy file ảnh hợp lệ để xử lý.")
            return

        self.processed_count = 0
        self.extracted_data = []
        self._clear_tree()
        self._clear_logs()
        self._set_running_state(True)

        self.worker_thread = threading.Thread(
            target=self._run_extraction,
            args=(root_dir, layout, total_images),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_extraction(
        self,
        root_dir: Path,
        layout: str,
        total_images: int,
    ) -> None:
        log_handlers: list[logging.Handler] = []

        try:
            log_dir = Path(tempfile.gettempdir()) / "qr_extractor_gui"
            self._ensure_writable_directory(log_dir)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_path = log_dir / f"extraction_log_{timestamp}.txt"
            
            log_handlers = self._setup_run_logging(log_path)

            logger = logging.getLogger("qr_gui")
            logger.info("Bắt đầu quét thư mục: %s", root_dir)
            logger.info("Layout detected: %s", layout)

            if layout == "nested" and get_image_files(root_dir):
                logger.warning(
                    "Phát hiện cả ảnh trực tiếp trong thư mục gốc và thư mục con. "
                    "Theo logic hiện tại, chỉ các thư mục con cấp 1 sẽ được xử lý."
                )

            targets, _ = get_targets(root_dir)
            logger.info("Tìm thấy %d file ảnh.", total_images)

            scanner = QRScanner()
            folders_data: list[dict] = []
            success_total = 0
            fail_total = 0

            for target in targets:
                folder_name = target.name
                reading_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                image_files = get_image_files(target)
                details = []
                success_count = 0
                fail_count = 0

                if not image_files:
                    logger.warning("Không tìm thấy ảnh trong nhóm: %s", folder_name)

                for image_path in image_files:
                    qr_content, status = scanner.scan_image(image_path)
                    if status == "success":
                        success_count += 1
                        logger.info(
                            "Đang xử lý: %s/%s -> thành công",
                            folder_name,
                            image_path.name,
                        )
                    else:
                        fail_count += 1
                        logger.info(
                            "Đang xử lý: %s/%s -> fail (không tìm thấy QR)",
                            folder_name,
                            image_path.name,
                        )

                    details.append(
                        {
                            "folder": folder_name,
                            "file_name": image_path.name,
                            "qr_content": qr_content,
                            "status": status,
                        }
                    )
                    self.event_queue.put(
                        (
                            "row",
                            {
                                "folder": folder_name,
                                "file_name": image_path.name,
                                "qr_content": qr_content,
                                "status": status,
                            },
                        )
                    )

                folders_data.append(
                    {
                        "folder_name": folder_name,
                        "reading_date": reading_date,
                        "total_files": len(image_files),
                        "success_count": success_count,
                        "fail_count": fail_count,
                        "details": details,
                    }
                )
                success_total += success_count
                fail_total += fail_count

                logger.info(
                    "Tổng kết nhóm %s: %d file, %d thành công, %d thất bại.",
                    folder_name,
                    len(image_files),
                    success_count,
                    fail_count,
                )

            logger.info("Hoàn tất quét. Đã sẵn sàng xuất dữ liệu ra Excel.")

            self.event_queue.put(
                (
                    "done",
                    {
                        "ok": True,
                        "folders_data": folders_data,
                        "processed": total_images,
                        "success": success_total,
                        "fail": fail_total,
                    },
                )
            )
        except Exception as exc:
            if log_handlers:
                logging.getLogger("qr_gui").exception("Quá trình quét gặp lỗi: %s", exc)
            self.event_queue.put(
                (
                    "done",
                    {
                        "ok": False,
                        "error": str(exc),
                    },
                )
            )
        finally:
            self._teardown_run_logging(log_handlers)

    def _ensure_writable_directory(self, directory: Path) -> bool:
        try:
            directory.mkdir(parents=True, exist_ok=True)
            probe = tempfile.NamedTemporaryFile(dir=directory, delete=True)
            probe.close()
            return True
        except Exception:
            return False

    def _setup_run_logging(self, log_path: Path) -> list[logging.Handler]:
        formatter = logging.Formatter("[%(asctime)s] %(message)s", "%H:%M:%S")
        handlers: list[logging.Handler] = []

        queue_handler = QueueLogHandler(self.event_queue)
        queue_handler.setFormatter(formatter)
        handlers.append(queue_handler)

        file_handler = logging.FileHandler(log_path, mode="w", encoding="utf-8")
        file_handler.setFormatter(formatter)
        handlers.append(file_handler)

        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        for handler in handlers:
            root_logger.addHandler(handler)

        return handlers

    def _teardown_run_logging(self, handlers: list[logging.Handler]) -> None:
        root_logger = logging.getLogger()
        for handler in handlers:
            root_logger.removeHandler(handler)
            handler.close()

    def _process_event_queue(self) -> None:
        while True:
            try:
                event_type, payload = self.event_queue.get_nowait()
            except queue.Empty:
                break

            if event_type == "log":
                self._append_log(str(payload))
            elif event_type == "row":
                self._append_result_row(payload)
            elif event_type == "done":
                self._handle_run_finished(payload)

        self.root.after(100, self._process_event_queue)

    def _append_log(self, message: str) -> None:
        assert self.log_text is not None
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _append_result_row(self, data: dict) -> None:
        assert self.tree is not None
        self.processed_count += 1
        self.tree.insert(
            "",
            "end",
            values=(
                self.processed_count,
                data["folder"],
                data["file_name"],
                data["qr_content"],
                data["status"],
            ),
        )

    def _handle_run_finished(self, payload: dict) -> None:
        self._set_running_state(False)

        if payload.get("ok"):
            self.extracted_data = payload.get("folders_data", [])
            if self.export_button is not None and self.extracted_data:
                self.export_button.configure(state="normal")

            messagebox.showinfo(
                "Hoàn thành",
                (
                    f"Hoàn thành! Đã xử lý {payload['processed']} file, "
                    f"thành công {payload['success']}, thất bại {payload['fail']}."
                ),
            )
            return

        if self.extracted_data and self.export_button is not None:
            self.export_button.configure(state="normal")

        error_text = payload.get("error", "Không rõ nguyên nhân.")
        messagebox.showerror(
            "Lỗi",
            f"Quá trình quét thất bại.\nChi tiết:\n{error_text}",
        )

    def _set_running_state(self, running: bool) -> None:
        if self.start_button is not None:
            self.start_button.configure(state="disabled" if running else "normal")

        if self.export_button is not None:
            self.export_button.configure(state="disabled" if running else "disabled")

        assert self.progress_bar is not None
        if running:
            self.progress_bar.grid()
            self.progress_bar.start(12)
        else:
            self.progress_bar.stop()
            self.progress_bar.grid_remove()

    def _clear_tree(self) -> None:
        assert self.tree is not None
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _clear_logs(self) -> None:
        assert self.log_text is not None
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _export_excel(self) -> None:
        if not self.extracted_data:
            messagebox.showwarning(
                "Cảnh báo", "Chưa có dữ liệu để xuất. Hãy chạy quét trước."
            )
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"QR_Extraction_Result_{timestamp}.xlsx"
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            title="Lưu file Excel kết quả",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )

        if not filepath:
            return  # User cancelled

        try:
            write_excel(Path(filepath), self.extracted_data)
            messagebox.showinfo("Thành công", f"Đã lưu kết quả tại:\n{filepath}")
        except Exception as exc:
            messagebox.showerror(
                "Lỗi", f"Không thể lưu file Excel.\nChi tiết:\n{exc}"
            )

    def _on_close(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            should_close = messagebox.askyesno(
                "Đang xử lý",
                "Tiến trình quét vẫn đang chạy. Bạn có chắc muốn thoát không?",
            )
            if not should_close:
                return

        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    app = QRExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
