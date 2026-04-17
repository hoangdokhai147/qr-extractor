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
from src.io.file_utils import get_image_files
from src.qrcode.scanner import QRScanner


def get_resource_path(relative_path: str) -> Path:
    """Get absolute path for resources"""
    try:
        base_path = Path(sys._MEIPASS)  # type: ignore
    except AttributeError:
        base_path = Path(__file__).resolve().parent

    return base_path / relative_path


APP_TITLE = "Pro QR Extractor"
LOGO_PATH = get_resource_path("static/logo.png")


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
        self.event_queue.put(("log", {"level": record.levelname, "message": message}))


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
        self.last_excel_path: Path | None = None
        self._logo_image: ImageTk.PhotoImage | None = None

        self.root_dir_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()

        self.start_button: ttk.Button | None = None
        self.open_button: ttk.Button | None = None
        self.progress_bar: ttk.Progressbar | None = None
        self.tree: ttk.Treeview | None = None
        self.log_text: scrolledtext.ScrolledText | None = None

        self._build_ui()

        now_str = datetime.now().strftime("%H:%M:%S")
        self.event_queue.put(
            ("log", {"level": "INFO", "message": f"[{now_str}] Sẵn sàng nhận lệnh..."})
        )

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
            ("Mô tả:", "Phần mềm đọc mã QR từ ảnh trong folder"),
            ("Phiên bản:", "1.0.0"),
            ("Ngày phát hành:", "16/04/2026"),
            (
                "Hướng dẫn:",
                "1. Chọn thư mục chứa ảnh\n(Lưu ý: Trong thư mục chỉ có ảnh, không được có thư mục con)\n\n2. Chọn nơi lưu Excel\n(Nếu không chọn, mặc định sẽ là folder chứa ảnh cần quét)\n\n3. Bấm 'Bắt đầu quét'",
            ),
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

        ttk.Label(right_frame, text="Thư mục lưu Excel kết quả:").grid(
            row=1, column=0, sticky="w", pady=(0, 8)
        )
        ttk.Entry(right_frame, textvariable=self.output_dir_var).grid(
            row=1, column=1, sticky="ew", padx=(12, 8), pady=(0, 8)
        )
        ttk.Button(right_frame, text="Browse", command=self._browse_output_folder).grid(
            row=1, column=2, sticky="ew", pady=(0, 8)
        )

        action_frame = ttk.Frame(right_frame)
        action_frame.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 10))

        self.start_button = ttk.Button(
            action_frame, text="Bắt đầu quét", command=self._start_scan
        )
        self.start_button.grid(row=0, column=0, padx=(0, 8))

        self.open_button = ttk.Button(
            action_frame,
            text="Mở file Excel kết quả",
            command=self._open_last_excel,
            state="disabled",
        )
        self.open_button.grid(row=0, column=1)

        self.progress_bar = ttk.Progressbar(right_frame, mode="determinate")
        self.progress_bar.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        self.progress_bar.grid_remove()

        table_frame = ttk.Frame(right_frame)
        table_frame.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(0, 12))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = (
            "stt", "folder", "file_name", "qr_content", 
            "col1", "col2", "col3", "col4", "col5", "col6", 
            "status"
        )
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        headings = {
            "stt": "STT",
            "folder": "Folder",
            "file_name": "File name",
            "qr_content": "QR content",
            "col1": "Mã (Dòng 1)",
            "col2": "TÊN DỰ ÁN",
            "col3": "TÊN CỘT",
            "col4": "KIỆN SỐ",
            "col5": "SỐ CHI TIẾT",
            "col6": "KL TĨNH",
            "status": "Status",
        }
        widths = {
            "stt": 40,
            "folder": 100,
            "file_name": 150,
            "qr_content": 220,
            "col1": 100,
            "col2": 150,
            "col3": 100,
            "col4": 80,
            "col5": 80,
            "col6": 80,
            "status": 70,
        }
        for column_id in columns:
            self.tree.heading(column_id, text=headings[column_id])
            self.tree.column(
                column_id,
                width=widths[column_id],
                minwidth=widths[column_id],
                anchor="w",
                stretch=False,
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
        self.log_text.tag_configure("neon_green", foreground="#39ff14")
        self.log_text.tag_configure(
            "error_red", foreground="#ff4444", font=("TkDefaultFont", 10, "bold")
        )
        self.log_text.configure(state="disabled")

    def _load_logo(self, label: ttk.Label) -> None:
        if not LOGO_PATH.exists():
            label.configure(text="static/logo.png")
            return

        try:
            image = Image.open(LOGO_PATH)
            image.thumbnail((160, 160))
            self._logo_image = ImageTk.PhotoImage(image)
            label.configure(image=self._logo_image)
        except Exception:
            label.configure(text="static/logo.png")

    def _browse_root_folder(self) -> None:
        directory = filedialog.askdirectory()
        if directory:
            self.root_dir_var.set(directory)

    def _browse_output_folder(self) -> None:
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir_var.set(directory)

    def _log_and_show_error(self, message: str, is_warning: bool = False) -> None:
        now_str = datetime.now().strftime("%H:%M:%S")
        level_str = "WARNING" if is_warning else "ERROR"
        prefix = "CẢNH BÁO" if is_warning else "LỖI"

        self.event_queue.put(
            ("log", {"level": level_str, "message": f"[{now_str}] {prefix}: {message}"})
        )

        if is_warning:
            messagebox.showwarning("Cảnh báo", message)
        else:
            messagebox.showerror("Lỗi", message)

    def _start_scan(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return

        root_input = self.root_dir_var.get().strip()
        if not root_input:
            self._log_and_show_error("Vui lòng chọn thư mục chứa ảnh.")
            return

        root_dir = Path(root_input).expanduser()
        if not root_dir.exists() or not root_dir.is_dir():
            self._log_and_show_error(f"Thư mục không hợp lệ:\n{root_dir}")
            return

        image_files = get_image_files(root_dir)
        total_images = len(image_files)
        if total_images == 0:
            self._log_and_show_error(
                "Không tìm thấy file ảnh hợp lệ để xử lý trong thư mục này."
            )
            return

        output_input = self.output_dir_var.get().strip()
        requested_output_dir = (
            Path(output_input).expanduser() if output_input else root_dir
        )

        self.processed_count = 0
        self.extracted_data = []
        self._clear_tree()
        # self._clear_logs() -> commented out so we don't erase the greeting logs

        assert self.progress_bar is not None
        self.progress_bar["maximum"] = total_images
        self.progress_bar["value"] = 0
        self._set_running_state(True)

        self.worker_thread = threading.Thread(
            target=self._run_extraction,
            args=(root_dir, requested_output_dir, total_images),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_extraction(
        self,
        root_dir: Path,
        requested_output_dir: Path,
        total_images: int,
    ) -> None:
        log_handlers: list[logging.Handler] = []
        active_output_dir: Path | None = None
        new_excel_path: Path | None = None

        try:
            active_output_dir, new_excel_path, log_path, output_warning = (
                self._resolve_output_targets(root_dir, requested_output_dir)
            )
            log_handlers = self._setup_run_logging(log_path)

            logger = logging.getLogger("qr_gui")
            logger.info("Bắt đầu quét thư mục: %s", root_dir)
            if output_warning:
                logger.warning("%s", output_warning)

            logger.info("Tìm thấy %d file ảnh.", total_images)

            scanner = QRScanner()
            folders_data: list[dict] = []

            folder_name = root_dir.name
            reading_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            image_files = get_image_files(root_dir)

            details = []
            success_total = 0
            fail_total = 0

            for image_path in image_files:
                qr_content, status = scanner.scan_image(image_path)
                if status == "success":
                    success_total += 1
                else:
                    fail_total += 1

                col1 = col2 = col3 = col4 = col5 = col6 = ""
                if qr_content:
                    lines = [
                        line.strip() for line in qr_content.split("\n") if line.strip()
                    ]
                    if len(lines) > 0:
                        col1 = lines[0]
                    if len(lines) > 1:
                        col2 = lines[1].replace("TEN DU AN:", "").strip()
                    if len(lines) > 2:
                        col3 = lines[2].replace("TEN COT:", "").strip()
                    if len(lines) > 3:
                        col4 = lines[3].replace("KIEN SO:", "").strip()
                    if len(lines) > 4:
                        col5 = lines[4].replace("SO CHI TIET:", "").strip()
                    if len(lines) > 5:
                        col6 = lines[5].replace("KL TINH:", "").strip()

                row_data = {
                    "folder": folder_name,
                    "file_name": image_path.name,
                    "qr_content": qr_content,
                    "col1": col1,
                    "TEN DU AN": col2,
                    "TEN COT": col3,
                    "KIEN SO": col4,
                    "SO CHI TIET": col5,
                    "KL TINH": col6,
                    "status": status,
                }
                details.append(row_data)
                self.event_queue.put(("row", row_data))

            folders_data.append(
                {
                    "folder_name": folder_name,
                    "reading_date": reading_date,
                    "total_files": total_images,
                    "success_count": success_total,
                    "fail_count": fail_total,
                    "details": details,
                }
            )

            logger.info(
                "Tổng kết xong: %d file, %d thành công, %d thất bại.",
                total_images,
                success_total,
                fail_total,
            )

            assert new_excel_path is not None
            write_excel(new_excel_path, folders_data)
            logger.info("Excel đã lưu tại: %s", new_excel_path.resolve())

            self.event_queue.put(
                (
                    "done",
                    {
                        "ok": True,
                        "excel_path": new_excel_path,
                        "output_dir": active_output_dir,
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

    def _resolve_output_targets(
        self, root_dir: Path, requested_output_dir: Path
    ) -> tuple[Path, Path, Path, str | None]:
        requested_output_dir = requested_output_dir.expanduser()

        if self._ensure_writable_directory(requested_output_dir):
            active_output_dir = requested_output_dir
            warning = None
        else:
            active_output_dir = Path(tempfile.gettempdir()) / "qr_extractor_gui"
            self._ensure_writable_directory(active_output_dir)
            warning = (
                f"Không thể ghi vào thư mục output đã chọn. "
                f"Sử dụng thư mục tạm: {active_output_dir}"
            )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = active_output_dir / f"QR_Extraction_Result_{timestamp}.xlsx"
        log_path = active_output_dir / "extraction_log.txt"

        if requested_output_dir == root_dir and warning is None:
            active_output_dir = root_dir

        return active_output_dir, excel_path, log_path, warning

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
                self._append_log(payload)
            elif event_type == "row":
                self._append_result_row(payload)
            elif event_type == "done":
                self._handle_run_finished(payload)

        self.root.after(100, self._process_event_queue)

    def _append_log(self, data: dict) -> None:
        assert self.log_text is not None
        self.log_text.configure(state="normal")

        message = data.get("message", "") if isinstance(data, dict) else str(data)
        level = data.get("level", "INFO") if isinstance(data, dict) else "INFO"

        color_tag = (
            "error_red" if level in ("ERROR", "WARNING", "CRITICAL") else "neon_green"
        )

        if message.startswith("[") and "] " in message:
            idx = message.find("] ") + 2
            timestamp = message[:idx]
            content = message[idx:]
            self.log_text.insert("end", timestamp)  # Default color
            self.log_text.insert("end", content + "\n", color_tag)
        else:
            self.log_text.insert("end", f"{message}\n", color_tag)

        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _append_result_row(self, data: dict) -> None:
        assert self.tree is not None

        self.processed_count += 1

        if self.progress_bar is not None:
            self.progress_bar["value"] = self.processed_count

        self.tree.insert(
            "",
            "end",
            values=(
                self.processed_count,
                data["folder"],
                data["file_name"],
                data["qr_content"],
                data["col1"],
                data["TEN DU AN"],
                data["TEN COT"],
                data["KIEN SO"],
                data["SO CHI TIET"],
                data["KL TINH"],
                data["status"],
            ),
        )

    def _handle_run_finished(self, payload: dict) -> None:
        self._set_running_state(False)

        if payload.get("ok"):
            excel_path = payload.get("excel_path")
            if isinstance(excel_path, Path) and excel_path.exists():
                self.last_excel_path = excel_path
                if self.open_button is not None:
                    self.open_button.configure(state="normal")

            messagebox.showinfo(
                "Hoàn thành",
                (
                    f"Hoàn thành! Đã xử lý {payload['processed']} file, "
                    f"thành công {payload['success']}, thất bại {payload['fail']}.\n"
                    f"Excel đã tự động lưu tại:\n{excel_path}"
                ),
            )
            return

        if self.last_excel_path and self.last_excel_path.exists():
            if self.open_button is not None:
                self.open_button.configure(state="normal")

        error_text = payload.get("error", "Không rõ nguyên nhân.")
        self._log_and_show_error(
            f"Quá trình quét và xuất Excel thất bại.\nChi tiết:\n{error_text}"
        )

    def _set_running_state(self, running: bool) -> None:
        if self.start_button is not None:
            self.start_button.configure(state="disabled" if running else "normal")

        if self.open_button is not None:
            self.open_button.configure(state="disabled" if running else "disabled")

        assert self.progress_bar is not None
        if running:
            self.progress_bar.grid()

    def _clear_tree(self) -> None:
        assert self.tree is not None
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _clear_logs(self) -> None:
        assert self.log_text is not None
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _open_last_excel(self) -> None:
        if not self.last_excel_path or not self.last_excel_path.exists():
            self._log_and_show_error(
                "Chưa có file Excel nào được tạo. Hãy chạy quét trước.", is_warning=True
            )
            if self.open_button is not None:
                self.open_button.configure(state="disabled")
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(str(self.last_excel_path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self.last_excel_path)])
            else:
                subprocess.Popen(["xdg-open", str(self.last_excel_path)])
        except Exception as exc:
            self._log_and_show_error(
                f"Không thể mở file Excel đã tạo.\nChi tiết:\n{exc}"
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
