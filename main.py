import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import (
    CANCELLED_MESSAGE,
    CONFLICT_AUTO_RENAME,
    CONFLICT_OVERWRITE,
    CONFLICT_SKIP,
    batch_convert_folder,
    convert_pdf_to_docx,
    convert_pdf_to_images,
    convert_pdf_to_pptx,
    merge_pdfs,
    split_pdf,
)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    DND_FILES = None
    TkinterDnD = None

OP_PDF_TO_PPTX = "PDF -> PPTX"
OP_PDF_TO_DOCX = "PDF -> DOCX"
OP_PDF_TO_PNG = "PDF -> PNG"
OP_PDF_TO_JPG = "PDF -> JPG"
OP_MERGE = "Merge PDFs"
OP_SPLIT = "Split PDF"
OP_BATCH = "Batch Convert Folder"

OPERATIONS = (
    OP_PDF_TO_PPTX,
    OP_PDF_TO_DOCX,
    OP_PDF_TO_PNG,
    OP_PDF_TO_JPG,
    OP_MERGE,
    OP_SPLIT,
    OP_BATCH,
)

BATCH_TARGET_FORMATS = ("PPTX", "DOCX", "PNG", "JPG")
CONFLICT_POLICY_DISPLAY = ("Overwrite", "Skip Existing", "Auto Rename")
CONFLICT_POLICY_MAP = {
    "Overwrite": CONFLICT_OVERWRITE,
    "Skip Existing": CONFLICT_SKIP,
    "Auto Rename": CONFLICT_AUTO_RENAME,
}


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Converter")
        self.root.geometry("960x780")
        self.root.resizable(True, True)
        self.root.minsize(900, 720)

        self.operation = tk.StringVar(value=OP_PDF_TO_PPTX)
        self.page_range = tk.StringVar(value="")
        self.batch_target_format = tk.StringVar(value="PPTX")
        self.conflict_policy_display = tk.StringVar(value="Auto Rename")
        self.input_password = tk.StringVar(value="")
        self.output_password = tk.StringVar(value="")
        self.render_dpi = tk.StringVar(value="144")
        self.jpg_quality = tk.StringVar(value="90")
        self.write_failure_log = tk.BooleanVar(value=True)

        self.selected_input: str | tuple[str, ...] = ""
        self.file_queue: list[str] = []
        self.cancel_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.is_running = False

        self.style = ttk.Style()
        self.style.configure("TButton", padding=6)

        main_frame = ttk.Frame(root, padding="16")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="PDF Converter", font=("Helvetica", 16, "bold")).pack(pady=(0, 12))

        operation_frame = ttk.Frame(main_frame)
        operation_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(operation_frame, text="Operation:", width=14).pack(side=tk.LEFT)
        self.operation_combo = ttk.Combobox(
            operation_frame,
            textvariable=self.operation,
            values=OPERATIONS,
            state="readonly",
            width=30,
        )
        self.operation_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.operation_combo.bind("<<ComboboxSelected>>", self.on_operation_changed)

        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 8))
        self.input_label = ttk.Label(
            input_frame,
            text="No input selected",
            anchor="w",
            relief="sunken",
            padding=(6, 6),
        )
        self.input_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.select_btn = ttk.Button(input_frame, text="Select Input", command=self.select_input)
        self.select_btn.pack(side=tk.RIGHT)

        queue_frame = ttk.LabelFrame(main_frame, text="File Queue")
        queue_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        queue_body = ttk.Frame(queue_frame, padding="8")
        queue_body.pack(fill=tk.BOTH, expand=True)
        queue_body.columnconfigure(0, weight=1)
        queue_body.rowconfigure(0, weight=1)

        self.queue_listbox = tk.Listbox(queue_body, selectmode=tk.EXTENDED, height=8)
        self.queue_listbox.grid(row=0, column=0, sticky="nsew")

        queue_scroll = ttk.Scrollbar(queue_body, orient=tk.VERTICAL, command=self.queue_listbox.yview)
        queue_scroll.grid(row=0, column=1, sticky="ns")
        self.queue_listbox.config(yscrollcommand=queue_scroll.set)

        queue_button_frame = ttk.Frame(queue_body)
        queue_button_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.queue_add_btn = ttk.Button(queue_button_frame, text="Add PDFs", command=self.add_queue_files)
        self.queue_add_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.queue_remove_btn = ttk.Button(
            queue_button_frame,
            text="Remove Selected",
            command=self.remove_queue_selection,
        )
        self.queue_remove_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.queue_clear_btn = ttk.Button(queue_button_frame, text="Clear Queue", command=self.clear_queue)
        self.queue_clear_btn.pack(side=tk.LEFT)

        self.queue_hint_label = ttk.Label(queue_body, foreground="gray")
        self.queue_hint_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        options_frame = ttk.LabelFrame(main_frame, text="Options")
        options_frame.pack(fill=tk.X, pady=(0, 8))
        options_frame.columnconfigure(1, weight=1)
        options_frame.columnconfigure(3, weight=1)

        ttk.Label(options_frame, text="Page range:").grid(row=0, column=0, sticky="w", padx=8, pady=5)
        self.page_range_entry = ttk.Entry(options_frame, textvariable=self.page_range)
        self.page_range_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=5)

        ttk.Label(options_frame, text="Conflict policy:").grid(row=0, column=2, sticky="w", padx=8, pady=5)
        self.conflict_combo = ttk.Combobox(
            options_frame,
            textvariable=self.conflict_policy_display,
            values=CONFLICT_POLICY_DISPLAY,
            state="readonly",
            width=16,
        )
        self.conflict_combo.grid(row=0, column=3, sticky="ew", padx=8, pady=5)

        ttk.Label(options_frame, text="Batch target:").grid(row=1, column=0, sticky="w", padx=8, pady=5)
        self.batch_combo = ttk.Combobox(
            options_frame,
            textvariable=self.batch_target_format,
            values=BATCH_TARGET_FORMATS,
            state="readonly",
            width=12,
        )
        self.batch_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=5)

        ttk.Label(options_frame, text="Input PDF password:").grid(row=1, column=2, sticky="w", padx=8, pady=5)
        self.input_password_entry = ttk.Entry(options_frame, textvariable=self.input_password, show="*")
        self.input_password_entry.grid(row=1, column=3, sticky="ew", padx=8, pady=5)

        ttk.Label(options_frame, text="Output PDF password:").grid(row=2, column=0, sticky="w", padx=8, pady=5)
        self.output_password_entry = ttk.Entry(options_frame, textvariable=self.output_password, show="*")
        self.output_password_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=5)

        ttk.Label(options_frame, text="Render DPI:").grid(row=2, column=2, sticky="w", padx=8, pady=5)
        self.render_dpi_spin = ttk.Spinbox(options_frame, from_=72, to=600, increment=12, textvariable=self.render_dpi)
        self.render_dpi_spin.grid(row=2, column=3, sticky="ew", padx=8, pady=5)

        ttk.Label(options_frame, text="JPG quality:").grid(row=3, column=0, sticky="w", padx=8, pady=5)
        self.jpg_quality_spin = ttk.Spinbox(options_frame, from_=1, to=100, increment=1, textvariable=self.jpg_quality)
        self.jpg_quality_spin.grid(row=3, column=1, sticky="ew", padx=8, pady=5)

        self.failure_log_check = ttk.Checkbutton(
            options_frame,
            text="Save batch failure log (CSV)",
            variable=self.write_failure_log,
        )
        self.failure_log_check.grid(row=3, column=2, columnspan=2, sticky="w", padx=8, pady=5)

        self.page_help = ttk.Label(
            main_frame,
            text="Page range format: 1-3,5,8-10 (leave empty for all pages).",
            foreground="gray",
        )
        self.page_help.pack(anchor="w", pady=(0, 8))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(4, 10))

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X)
        self.convert_btn = ttk.Button(action_frame, text="Convert", command=self.start_conversion, state=tk.DISABLED)
        self.convert_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.cancel_btn = ttk.Button(
            action_frame,
            text="Cancel",
            command=self.cancel_conversion,
            state=tk.DISABLED,
        )
        self.cancel_btn.pack(side=tk.LEFT)

        self.status_label = ttk.Label(main_frame, text="Ready", foreground="gray")
        self.status_label.pack(pady=(10, 0), anchor="w")

        self._bind_drag_and_drop()
        self.on_operation_changed()

    def _bind_drag_and_drop(self):
        if not DND_AVAILABLE:
            self.queue_hint_label.config(text="Drag-and-drop unavailable. Install `tkinterdnd2` to enable it.")
            return

        try:
            self.queue_listbox.drop_target_register(DND_FILES)
            self.queue_listbox.dnd_bind("<<Drop>>", self.on_drop_files)
            self.queue_hint_label.config(text="Drag PDF files into this queue.")
        except Exception:
            self.queue_hint_label.config(text="Drag-and-drop initialization failed. Queue still works with Add PDFs.")

    def on_drop_files(self, event):
        dropped_paths = self.root.tk.splitlist(event.data)
        self._add_files_to_queue(dropped_paths)
        return "break"

    def _add_files_to_queue(self, paths):
        existing = set(self.file_queue)
        added_count = 0
        for path in paths:
            normalized_path = os.path.normpath(path)
            if not normalized_path.lower().endswith(".pdf"):
                continue
            if not os.path.isfile(normalized_path):
                continue
            if normalized_path in existing:
                continue
            self.file_queue.append(normalized_path)
            existing.add(normalized_path)
            added_count += 1

        if added_count > 0:
            self._refresh_queue_listbox()
            self._update_input_label()
            self._update_convert_state()

    def _refresh_queue_listbox(self):
        self.queue_listbox.delete(0, tk.END)
        for queued_file in self.file_queue:
            self.queue_listbox.insert(tk.END, queued_file)

    def add_queue_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if paths:
            self._add_files_to_queue(paths)

    def remove_queue_selection(self):
        selected_indices = list(self.queue_listbox.curselection())
        if not selected_indices:
            return
        for index in reversed(selected_indices):
            self.file_queue.pop(index)
        self._refresh_queue_listbox()
        self._update_input_label()
        self._update_convert_state()

    def clear_queue(self):
        self.file_queue.clear()
        self._refresh_queue_listbox()
        self._update_input_label()
        self._update_convert_state()

    def _resolved_input_for_operation(self, operation: str | None = None) -> str | tuple[str, ...]:
        target_operation = operation or self.operation.get()

        if target_operation == OP_BATCH:
            if isinstance(self.selected_input, str):
                return self.selected_input
            return ""

        if target_operation == OP_MERGE:
            if self.file_queue:
                return tuple(self.file_queue)
            if isinstance(self.selected_input, tuple):
                return self.selected_input
            return ()

        if isinstance(self.selected_input, str) and self.selected_input:
            return self.selected_input
        if self.file_queue:
            return self.file_queue[0]
        return ""

    def _has_input(self) -> bool:
        resolved_input = self._resolved_input_for_operation()
        if isinstance(resolved_input, tuple):
            return len(resolved_input) > 0
        return bool(resolved_input)

    def _update_input_label(self):
        operation = self.operation.get()
        resolved_input = self._resolved_input_for_operation(operation)

        if operation == OP_BATCH:
            if isinstance(resolved_input, str) and resolved_input:
                self.input_label.config(text=resolved_input)
            else:
                self.input_label.config(text="No input folder selected")
            return

        if operation == OP_MERGE:
            if isinstance(resolved_input, tuple) and resolved_input:
                self.input_label.config(text=f"{len(resolved_input)} PDF files selected")
            else:
                self.input_label.config(text="No PDF files selected")
            return

        if isinstance(resolved_input, str) and resolved_input:
            if resolved_input in self.file_queue:
                self.input_label.config(text=f"{os.path.basename(resolved_input)} (from queue)")
            else:
                self.input_label.config(text=os.path.basename(resolved_input))
        else:
            self.input_label.config(text="No PDF selected")

    def _refresh_dynamic_controls(self):
        operation = self.operation.get()
        if self.is_running:
            return

        if operation == OP_MERGE:
            self.select_btn.config(text="Select PDFs")
        elif operation == OP_BATCH:
            self.select_btn.config(text="Select Folder")
        else:
            self.select_btn.config(text="Select PDF")

        if operation == OP_MERGE:
            self.page_range_entry.config(state=tk.DISABLED)
        else:
            self.page_range_entry.config(state=tk.NORMAL)

        if operation == OP_BATCH:
            self.batch_combo.config(state="readonly")
        else:
            self.batch_combo.config(state=tk.DISABLED)

        if operation in (OP_MERGE, OP_SPLIT):
            self.output_password_entry.config(state=tk.NORMAL)
        else:
            self.output_password_entry.config(state=tk.DISABLED)

        if operation in (OP_PDF_TO_PPTX, OP_PDF_TO_PNG, OP_PDF_TO_JPG, OP_BATCH):
            self.render_dpi_spin.config(state=tk.NORMAL)
        else:
            self.render_dpi_spin.config(state=tk.DISABLED)

        if operation in (OP_PDF_TO_JPG, OP_BATCH):
            self.jpg_quality_spin.config(state=tk.NORMAL)
        else:
            self.jpg_quality_spin.config(state=tk.DISABLED)

        if operation == OP_BATCH:
            self.failure_log_check.config(state=tk.NORMAL)
        else:
            self.failure_log_check.config(state=tk.DISABLED)

        button_labels = {
            OP_PDF_TO_PPTX: "Convert to PPTX",
            OP_PDF_TO_DOCX: "Convert to DOCX",
            OP_PDF_TO_PNG: "Convert to PNG",
            OP_PDF_TO_JPG: "Convert to JPG",
            OP_MERGE: "Merge PDFs",
            OP_SPLIT: "Split PDF",
            OP_BATCH: "Batch Convert",
        }
        self.convert_btn.config(text=button_labels.get(operation, "Run"))

    def _update_convert_state(self):
        if self.is_running:
            self.convert_btn.config(state=tk.DISABLED)
            return
        self.convert_btn.config(state=tk.NORMAL if self._has_input() else tk.DISABLED)

    def on_operation_changed(self, _event=None):
        operation = self.operation.get()

        if operation == OP_BATCH and isinstance(self.selected_input, tuple):
            self.selected_input = ""
        elif operation != OP_BATCH and isinstance(self.selected_input, str) and os.path.isdir(self.selected_input):
            self.selected_input = ""

        if operation == OP_MERGE and isinstance(self.selected_input, str):
            self.selected_input = ()
        elif operation != OP_MERGE and isinstance(self.selected_input, tuple):
            self.selected_input = ""

        self.progress_var.set(0)
        self.status_label.config(text="Ready", foreground="gray")
        self._refresh_dynamic_controls()
        self._update_input_label()
        self._update_convert_state()

    def select_input(self):
        operation = self.operation.get()

        if operation == OP_MERGE:
            paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
            if paths:
                self.selected_input = tuple(paths)
                self._add_files_to_queue(paths)
                self._update_input_label()
                self._update_convert_state()
            return

        if operation == OP_BATCH:
            folder = filedialog.askdirectory(title="Select input folder")
            if folder:
                self.selected_input = folder
                self._update_input_label()
                self._update_convert_state()
            return

        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.selected_input = file_path
            self._add_files_to_queue([file_path])
            self._update_input_label()
            self._update_convert_state()

    def _parse_numeric_options(self) -> tuple[int, int] | None:
        try:
            render_dpi = int(self.render_dpi.get().strip())
            jpg_quality = int(self.jpg_quality.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Render DPI and JPG quality must be numeric values.")
            return None

        if render_dpi < 72 or render_dpi > 600:
            messagebox.showerror("Error", "Render DPI must be between 72 and 600.")
            return None
        if jpg_quality < 1 or jpg_quality > 100:
            messagebox.showerror("Error", "JPG quality must be between 1 and 100.")
            return None

        return render_dpi, jpg_quality

    def start_conversion(self):
        if self.is_running:
            return

        if not self._has_input():
            messagebox.showwarning("Warning", "Select an input first.")
            return

        numeric_options = self._parse_numeric_options()
        if not numeric_options:
            return
        render_dpi, jpg_quality = numeric_options

        operation = self.operation.get()
        input_data = self._resolved_input_for_operation(operation)

        if operation == OP_MERGE and isinstance(input_data, tuple) and len(input_data) < 2:
            messagebox.showwarning("Warning", "Please queue/select at least 2 PDF files to merge.")
            return

        output_target = self._ask_output_target(operation, input_data)
        if not output_target:
            return

        options = {
            "page_range_text": self.page_range.get().strip(),
            "input_password": self.input_password.get(),
            "output_password": self.output_password.get(),
            "output_conflict_policy": CONFLICT_POLICY_MAP[self.conflict_policy_display.get()],
            "render_dpi": render_dpi,
            "jpg_quality": jpg_quality,
            "batch_target_format": self.batch_target_format.get(),
            "write_failure_log": self.write_failure_log.get(),
        }

        self.cancel_event.clear()
        self._set_controls_running(True)
        self.status_label.config(text="Working...", foreground="blue")

        self.worker_thread = threading.Thread(
            target=self.run_conversion,
            args=(operation, input_data, output_target, options),
            daemon=True,
        )
        self.worker_thread.start()

    def _set_controls_running(self, is_running: bool):
        self.is_running = is_running
        if is_running:
            self.convert_btn.config(state=tk.DISABLED)
            self.cancel_btn.config(state=tk.NORMAL)
            self.select_btn.config(state=tk.DISABLED)
            self.operation_combo.config(state=tk.DISABLED)
            self.page_range_entry.config(state=tk.DISABLED)
            self.batch_combo.config(state=tk.DISABLED)
            self.conflict_combo.config(state=tk.DISABLED)
            self.input_password_entry.config(state=tk.DISABLED)
            self.output_password_entry.config(state=tk.DISABLED)
            self.render_dpi_spin.config(state=tk.DISABLED)
            self.jpg_quality_spin.config(state=tk.DISABLED)
            self.failure_log_check.config(state=tk.DISABLED)
            self.queue_add_btn.config(state=tk.DISABLED)
            self.queue_remove_btn.config(state=tk.DISABLED)
            self.queue_clear_btn.config(state=tk.DISABLED)
            self.queue_listbox.config(state=tk.DISABLED)
        else:
            self.cancel_btn.config(state=tk.DISABLED)
            self.select_btn.config(state=tk.NORMAL)
            self.operation_combo.config(state="readonly")
            self.conflict_combo.config(state="readonly")
            self.input_password_entry.config(state=tk.NORMAL)
            self.queue_add_btn.config(state=tk.NORMAL)
            self.queue_remove_btn.config(state=tk.NORMAL)
            self.queue_clear_btn.config(state=tk.NORMAL)
            self.queue_listbox.config(state=tk.NORMAL)
            self._refresh_dynamic_controls()
            self._update_convert_state()

    def cancel_conversion(self):
        if not self.is_running:
            return
        self.cancel_event.set()
        self.status_label.config(text="Cancelling...", foreground="orange")

    def _ask_output_target(self, operation: str, input_data: str | tuple[str, ...]) -> str:
        if operation == OP_MERGE:
            return filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                initialfile="merged.pdf",
            )

        if operation in (OP_PDF_TO_PPTX, OP_PDF_TO_DOCX):
            assert isinstance(input_data, str)
            base_name = os.path.splitext(os.path.basename(input_data))[0]
            extension = ".pptx" if operation == OP_PDF_TO_PPTX else ".docx"
            file_type = (
                ("PowerPoint Presentation", "*.pptx")
                if operation == OP_PDF_TO_PPTX
                else ("Word Document", "*.docx")
            )
            return filedialog.asksaveasfilename(
                defaultextension=extension,
                filetypes=[file_type],
                initialfile=f"{base_name}{extension}",
            )

        if operation in (OP_PDF_TO_PNG, OP_PDF_TO_JPG, OP_SPLIT, OP_BATCH):
            return filedialog.askdirectory(title="Select output folder")

        return ""

    def run_conversion(
        self,
        operation: str,
        input_data: str | tuple[str, ...],
        output_target: str,
        options: dict,
    ):
        def update_progress(percent: int):
            self.root.after(0, lambda: self.progress_var.set(percent))

        try:
            if operation == OP_PDF_TO_PPTX:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_pptx(
                    input_data,
                    output_target,
                    update_progress,
                    options["page_range_text"],
                    options["input_password"],
                    options["output_conflict_policy"],
                    options["render_dpi"],
                    self.cancel_event,
                )
            elif operation == OP_PDF_TO_DOCX:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_docx(
                    input_data,
                    output_target,
                    update_progress,
                    options["page_range_text"],
                    options["input_password"],
                    options["output_conflict_policy"],
                    self.cancel_event,
                )
            elif operation == OP_PDF_TO_PNG:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_images(
                    input_data,
                    output_target,
                    "png",
                    options["render_dpi"],
                    update_progress,
                    options["page_range_text"],
                    options["input_password"],
                    options["output_conflict_policy"],
                    options["jpg_quality"],
                    self.cancel_event,
                )
            elif operation == OP_PDF_TO_JPG:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_images(
                    input_data,
                    output_target,
                    "jpg",
                    options["render_dpi"],
                    update_progress,
                    options["page_range_text"],
                    options["input_password"],
                    options["output_conflict_policy"],
                    options["jpg_quality"],
                    self.cancel_event,
                )
            elif operation == OP_MERGE:
                assert isinstance(input_data, tuple)
                success, message = merge_pdfs(
                    input_data,
                    output_target,
                    update_progress,
                    options["input_password"],
                    options["output_password"],
                    options["output_conflict_policy"],
                    self.cancel_event,
                )
            elif operation == OP_SPLIT:
                assert isinstance(input_data, str)
                success, message = split_pdf(
                    input_data,
                    output_target,
                    update_progress,
                    options["page_range_text"],
                    options["input_password"],
                    options["output_password"],
                    options["output_conflict_policy"],
                    self.cancel_event,
                )
            elif operation == OP_BATCH:
                assert isinstance(input_data, str)
                success, message = batch_convert_folder(
                    input_data,
                    output_target,
                    options["batch_target_format"],
                    update_progress,
                    options["page_range_text"],
                    options["input_password"],
                    options["output_password"],
                    options["output_conflict_policy"],
                    options["render_dpi"],
                    options["jpg_quality"],
                    options["write_failure_log"],
                    self.cancel_event,
                )
            else:
                success, message = False, "Unknown operation."
        except Exception as exc:
            success, message = False, str(exc)

        self.root.after(0, lambda: self.conversion_finished(success, message))

    def conversion_finished(self, success: bool, message: str):
        self._set_controls_running(False)

        if success:
            self.status_label.config(text="Done", foreground="green")
            messagebox.showinfo("Success", message)
        else:
            if message.startswith(CANCELLED_MESSAGE):
                self.status_label.config(text="Cancelled", foreground="orange")
                messagebox.showinfo("Cancelled", message)
            else:
                self.status_label.config(text="Error", foreground="red")
                messagebox.showerror("Error", f"An error occurred:\n{message}")

        self.cancel_event.clear()
        self.progress_var.set(0)
        self.status_label.config(text="Ready", foreground="gray")
        self._update_input_label()
        self._update_convert_state()


def create_root() -> tk.Tk:
    if DND_AVAILABLE and TkinterDnD is not None:
        return TkinterDnD.Tk()
    return tk.Tk()


if __name__ == "__main__":
    root = create_root()
    app = App(root)
    root.mainloop()
