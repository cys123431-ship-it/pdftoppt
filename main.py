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
CONFLICT_POLICIES = (CONFLICT_OVERWRITE, CONFLICT_SKIP, CONFLICT_AUTO_RENAME)

LANG_OPTIONS = (("ko", "한국어"), ("en", "English"))
LANG_CODE_TO_LABEL = {code: label for code, label in LANG_OPTIONS}
LANG_LABEL_TO_CODE = {label: code for code, label in LANG_OPTIONS}

STATUS_COLORS = {
    "ready": "gray",
    "working": "blue",
    "done": "green",
    "cancelled": "orange",
    "error": "red",
    "cancelling": "orange",
}

I18N = {
    "ko": {
        "app_title": "PDF 변환기",
        "header_title": "PDF 변환기",
        "label_language": "언어:",
        "label_operation": "작업:",
        "no_input_selected": "입력이 선택되지 않았습니다",
        "queue_group": "파일 큐",
        "queue_add": "PDF 추가",
        "queue_remove": "선택 제거",
        "queue_clear": "큐 비우기",
        "queue_drop_unavailable": "드래그앤드롭을 사용하려면 `tkinterdnd2`를 설치하세요.",
        "queue_drop_ready": "여기로 PDF 파일을 드래그해서 큐에 추가할 수 있습니다.",
        "queue_drop_failed": "드래그앤드롭 초기화에 실패했습니다. 'PDF 추가' 버튼은 사용할 수 있습니다.",
        "options_group": "옵션",
        "label_page_range": "페이지 범위:",
        "label_conflict": "출력 충돌 정책:",
        "label_batch_target": "일괄 대상 형식:",
        "label_input_password": "입력 PDF 비밀번호:",
        "label_output_password": "출력 PDF 비밀번호:",
        "label_render_dpi": "렌더 DPI:",
        "label_jpg_quality": "JPG 품질:",
        "check_failure_log": "배치 실패 로그(CSV) 저장",
        "page_help": "페이지 범위 형식: 1-3,5,8-10 (비워두면 전체 페이지)",
        "btn_select_input": "입력 선택",
        "btn_select_pdf": "PDF 선택",
        "btn_select_pdfs": "PDF들 선택",
        "btn_select_folder": "폴더 선택",
        "btn_convert": "실행",
        "btn_convert_pptx": "PPTX로 변환",
        "btn_convert_docx": "DOCX로 변환",
        "btn_convert_png": "PNG로 변환",
        "btn_convert_jpg": "JPG로 변환",
        "btn_merge": "PDF 병합",
        "btn_split": "PDF 분할",
        "btn_batch": "일괄 변환",
        "btn_cancel": "취소",
        "status_ready": "준비됨",
        "status_working": "작업 중...",
        "status_done": "완료",
        "status_cancelled": "취소됨",
        "status_error": "오류",
        "status_cancelling": "취소하는 중...",
        "input_none_folder": "입력 폴더가 선택되지 않았습니다",
        "input_none_pdfs": "PDF 파일이 선택되지 않았습니다",
        "input_none_pdf": "PDF 파일이 선택되지 않았습니다",
        "input_selected_count": "{count}개 PDF 선택됨",
        "input_from_queue": "{name} (큐에서 선택)",
        "op_pdf_to_pptx": "PDF -> PPTX",
        "op_pdf_to_docx": "PDF -> DOCX",
        "op_pdf_to_png": "PDF -> PNG",
        "op_pdf_to_jpg": "PDF -> JPG",
        "op_merge": "PDF 병합",
        "op_split": "PDF 분할",
        "op_batch": "폴더 일괄 변환",
        "conflict_overwrite": "덮어쓰기",
        "conflict_skip": "기존 파일 건너뛰기",
        "conflict_auto_rename": "자동 이름 변경",
        "msg_numeric": "렌더 DPI와 JPG 품질은 숫자여야 합니다.",
        "msg_dpi": "렌더 DPI는 72~600 범위여야 합니다.",
        "msg_jpg": "JPG 품질은 1~100 범위여야 합니다.",
        "msg_select_input": "먼저 입력을 선택하세요.",
        "msg_need_two": "PDF 병합은 2개 이상의 PDF가 필요합니다.",
        "msg_unknown_operation": "알 수 없는 작업입니다.",
        "msg_error_prefix": "오류가 발생했습니다:\n{message}",
        "title_error": "오류",
        "title_success": "성공",
        "title_warning": "경고",
        "title_cancelled": "취소됨",
        "dialog_input_folder": "입력 폴더 선택",
        "dialog_output_folder": "출력 폴더 선택",
        "filetype_pdf": "PDF 파일",
        "filetype_pptx": "PowerPoint 프레젠테이션",
        "filetype_docx": "Word 문서",
    },
    "en": {
        "app_title": "PDF Converter",
        "header_title": "PDF Converter",
        "label_language": "Language:",
        "label_operation": "Operation:",
        "no_input_selected": "No input selected",
        "queue_group": "File Queue",
        "queue_add": "Add PDFs",
        "queue_remove": "Remove Selected",
        "queue_clear": "Clear Queue",
        "queue_drop_unavailable": "Drag-and-drop unavailable. Install `tkinterdnd2` to enable it.",
        "queue_drop_ready": "Drag PDF files into this queue.",
        "queue_drop_failed": "Drag-and-drop initialization failed. Queue still works with Add PDFs.",
        "options_group": "Options",
        "label_page_range": "Page range:",
        "label_conflict": "Conflict policy:",
        "label_batch_target": "Batch target:",
        "label_input_password": "Input PDF password:",
        "label_output_password": "Output PDF password:",
        "label_render_dpi": "Render DPI:",
        "label_jpg_quality": "JPG quality:",
        "check_failure_log": "Save batch failure log (CSV)",
        "page_help": "Page range format: 1-3,5,8-10 (leave empty for all pages).",
        "btn_select_input": "Select Input",
        "btn_select_pdf": "Select PDF",
        "btn_select_pdfs": "Select PDFs",
        "btn_select_folder": "Select Folder",
        "btn_convert": "Run",
        "btn_convert_pptx": "Convert to PPTX",
        "btn_convert_docx": "Convert to DOCX",
        "btn_convert_png": "Convert to PNG",
        "btn_convert_jpg": "Convert to JPG",
        "btn_merge": "Merge PDFs",
        "btn_split": "Split PDF",
        "btn_batch": "Batch Convert",
        "btn_cancel": "Cancel",
        "status_ready": "Ready",
        "status_working": "Working...",
        "status_done": "Done",
        "status_cancelled": "Cancelled",
        "status_error": "Error",
        "status_cancelling": "Cancelling...",
        "input_none_folder": "No input folder selected",
        "input_none_pdfs": "No PDF files selected",
        "input_none_pdf": "No PDF selected",
        "input_selected_count": "{count} PDF files selected",
        "input_from_queue": "{name} (from queue)",
        "op_pdf_to_pptx": "PDF -> PPTX",
        "op_pdf_to_docx": "PDF -> DOCX",
        "op_pdf_to_png": "PDF -> PNG",
        "op_pdf_to_jpg": "PDF -> JPG",
        "op_merge": "Merge PDFs",
        "op_split": "Split PDF",
        "op_batch": "Batch Convert Folder",
        "conflict_overwrite": "Overwrite",
        "conflict_skip": "Skip Existing",
        "conflict_auto_rename": "Auto Rename",
        "msg_numeric": "Render DPI and JPG quality must be numeric values.",
        "msg_dpi": "Render DPI must be between 72 and 600.",
        "msg_jpg": "JPG quality must be between 1 and 100.",
        "msg_select_input": "Select an input first.",
        "msg_need_two": "Please queue/select at least 2 PDF files to merge.",
        "msg_unknown_operation": "Unknown operation.",
        "msg_error_prefix": "An error occurred:\n{message}",
        "title_error": "Error",
        "title_success": "Success",
        "title_warning": "Warning",
        "title_cancelled": "Cancelled",
        "dialog_input_folder": "Select input folder",
        "dialog_output_folder": "Select output folder",
        "filetype_pdf": "PDF Files",
        "filetype_pptx": "PowerPoint Presentation",
        "filetype_docx": "Word Document",
    },
}

OP_LABEL_KEYS = {
    OP_PDF_TO_PPTX: "op_pdf_to_pptx",
    OP_PDF_TO_DOCX: "op_pdf_to_docx",
    OP_PDF_TO_PNG: "op_pdf_to_png",
    OP_PDF_TO_JPG: "op_pdf_to_jpg",
    OP_MERGE: "op_merge",
    OP_SPLIT: "op_split",
    OP_BATCH: "op_batch",
}

CONFLICT_LABEL_KEYS = {
    CONFLICT_OVERWRITE: "conflict_overwrite",
    CONFLICT_SKIP: "conflict_skip",
    CONFLICT_AUTO_RENAME: "conflict_auto_rename",
}


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.geometry("960x780")
        self.root.resizable(True, True)
        self.root.minsize(900, 720)

        self.language_code = tk.StringVar(value="ko")
        self.language_display = tk.StringVar(value=LANG_CODE_TO_LABEL["ko"])

        self.operation = tk.StringVar(value=OP_PDF_TO_PPTX)
        self.operation_display = tk.StringVar(value="")
        self.page_range = tk.StringVar(value="")
        self.batch_target_format = tk.StringVar(value="PPTX")

        self.conflict_policy = tk.StringVar(value=CONFLICT_AUTO_RENAME)
        self.conflict_policy_display = tk.StringVar(value="")

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
        self.status_key = "ready"
        self.queue_hint_mode = "ready"

        self.operation_display_to_value: dict[str, str] = {}
        self.operation_value_to_display: dict[str, str] = {}
        self.conflict_display_to_value: dict[str, str] = {}
        self.conflict_value_to_display: dict[str, str] = {}

        self.style = ttk.Style()
        self.style.configure("TButton", padding=6)

        main_frame = ttk.Frame(root, padding="16")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.title_label = ttk.Label(main_frame, font=("Helvetica", 16, "bold"))
        self.title_label.pack(pady=(0, 12))

        operation_frame = ttk.Frame(main_frame)
        operation_frame.pack(fill=tk.X, pady=(0, 8))
        self.language_label = ttk.Label(operation_frame, width=6)
        self.language_label.pack(side=tk.LEFT)
        self.language_combo = ttk.Combobox(
            operation_frame,
            textvariable=self.language_display,
            values=[label for _, label in LANG_OPTIONS],
            state="readonly",
            width=10,
        )
        self.language_combo.pack(side=tk.LEFT, padx=(0, 12))
        self.language_combo.bind("<<ComboboxSelected>>", self.on_language_changed)

        self.operation_label = ttk.Label(operation_frame, width=10)
        self.operation_label.pack(side=tk.LEFT)
        self.operation_combo = ttk.Combobox(
            operation_frame,
            textvariable=self.operation_display,
            values=[],
            state="readonly",
            width=30,
        )
        self.operation_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.operation_combo.bind("<<ComboboxSelected>>", self.on_operation_selection_changed)

        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 8))
        self.input_label = ttk.Label(
            input_frame,
            text="",
            anchor="w",
            relief="sunken",
            padding=(6, 6),
        )
        self.input_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.select_btn = ttk.Button(input_frame, command=self.select_input)
        self.select_btn.pack(side=tk.RIGHT)

        self.queue_frame = ttk.LabelFrame(main_frame, text="")
        self.queue_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        queue_body = ttk.Frame(self.queue_frame, padding="8")
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
        self.queue_add_btn = ttk.Button(queue_button_frame, command=self.add_queue_files)
        self.queue_add_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.queue_remove_btn = ttk.Button(
            queue_button_frame,
            command=self.remove_queue_selection,
        )
        self.queue_remove_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.queue_clear_btn = ttk.Button(queue_button_frame, command=self.clear_queue)
        self.queue_clear_btn.pack(side=tk.LEFT)

        self.queue_hint_label = ttk.Label(queue_body, foreground="gray")
        self.queue_hint_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.options_frame = ttk.LabelFrame(main_frame, text="")
        self.options_frame.pack(fill=tk.X, pady=(0, 8))
        self.options_frame.columnconfigure(1, weight=1)
        self.options_frame.columnconfigure(3, weight=1)

        self.page_range_label = ttk.Label(self.options_frame, text="")
        self.page_range_label.grid(row=0, column=0, sticky="w", padx=8, pady=5)
        self.page_range_entry = ttk.Entry(self.options_frame, textvariable=self.page_range)
        self.page_range_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=5)

        self.conflict_label = ttk.Label(self.options_frame, text="")
        self.conflict_label.grid(row=0, column=2, sticky="w", padx=8, pady=5)
        self.conflict_combo = ttk.Combobox(
            self.options_frame,
            textvariable=self.conflict_policy_display,
            values=[],
            state="readonly",
            width=16,
        )
        self.conflict_combo.grid(row=0, column=3, sticky="ew", padx=8, pady=5)
        self.conflict_combo.bind("<<ComboboxSelected>>", self.on_conflict_policy_changed)

        self.batch_target_label = ttk.Label(self.options_frame, text="")
        self.batch_target_label.grid(row=1, column=0, sticky="w", padx=8, pady=5)
        self.batch_combo = ttk.Combobox(
            self.options_frame,
            textvariable=self.batch_target_format,
            values=BATCH_TARGET_FORMATS,
            state="readonly",
            width=12,
        )
        self.batch_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=5)

        self.input_password_label = ttk.Label(self.options_frame, text="")
        self.input_password_label.grid(row=1, column=2, sticky="w", padx=8, pady=5)
        self.input_password_entry = ttk.Entry(self.options_frame, textvariable=self.input_password, show="*")
        self.input_password_entry.grid(row=1, column=3, sticky="ew", padx=8, pady=5)

        self.output_password_label = ttk.Label(self.options_frame, text="")
        self.output_password_label.grid(row=2, column=0, sticky="w", padx=8, pady=5)
        self.output_password_entry = ttk.Entry(self.options_frame, textvariable=self.output_password, show="*")
        self.output_password_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=5)

        self.render_dpi_label = ttk.Label(self.options_frame, text="")
        self.render_dpi_label.grid(row=2, column=2, sticky="w", padx=8, pady=5)
        self.render_dpi_spin = ttk.Spinbox(self.options_frame, from_=72, to=600, increment=12, textvariable=self.render_dpi)
        self.render_dpi_spin.grid(row=2, column=3, sticky="ew", padx=8, pady=5)

        self.jpg_quality_label = ttk.Label(self.options_frame, text="")
        self.jpg_quality_label.grid(row=3, column=0, sticky="w", padx=8, pady=5)
        self.jpg_quality_spin = ttk.Spinbox(self.options_frame, from_=1, to=100, increment=1, textvariable=self.jpg_quality)
        self.jpg_quality_spin.grid(row=3, column=1, sticky="ew", padx=8, pady=5)

        self.failure_log_check = ttk.Checkbutton(
            self.options_frame,
            text="",
            variable=self.write_failure_log,
        )
        self.failure_log_check.grid(row=3, column=2, columnspan=2, sticky="w", padx=8, pady=5)

        self.page_help = ttk.Label(
            main_frame,
            text="",
            foreground="gray",
        )
        self.page_help.pack(anchor="w", pady=(0, 8))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(4, 10))

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X)
        self.convert_btn = ttk.Button(action_frame, text="", command=self.start_conversion, state=tk.DISABLED)
        self.convert_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.cancel_btn = ttk.Button(
            action_frame,
            text="",
            command=self.cancel_conversion,
            state=tk.DISABLED,
        )
        self.cancel_btn.pack(side=tk.LEFT)

        self.status_label = ttk.Label(main_frame, text="", foreground="gray")
        self.status_label.pack(pady=(10, 0), anchor="w")

        self._bind_drag_and_drop()
        self._apply_language()
        self.on_operation_changed()

    def _t(self, key: str, **kwargs) -> str:
        text = I18N.get(self.language_code.get(), I18N["en"]).get(key, key)
        if kwargs:
            return text.format(**kwargs)
        return text

    def _set_status(self, status_key: str):
        self.status_key = status_key
        self.status_label.config(text=self._t(f"status_{status_key}"), foreground=STATUS_COLORS[status_key])

    def _refresh_operation_display_values(self):
        self.operation_value_to_display = {op: self._t(OP_LABEL_KEYS[op]) for op in OPERATIONS}
        self.operation_display_to_value = {text: op for op, text in self.operation_value_to_display.items()}
        self.operation_combo.config(values=list(self.operation_value_to_display.values()))
        self.operation_display.set(self.operation_value_to_display[self.operation.get()])

    def _refresh_conflict_display_values(self):
        self.conflict_value_to_display = {policy: self._t(CONFLICT_LABEL_KEYS[policy]) for policy in CONFLICT_POLICIES}
        self.conflict_display_to_value = {
            text: policy for policy, text in self.conflict_value_to_display.items()
        }
        self.conflict_combo.config(values=list(self.conflict_value_to_display.values()))
        self.conflict_policy_display.set(self.conflict_value_to_display[self.conflict_policy.get()])

    def _apply_language(self):
        self.root.title(self._t("app_title"))
        self.title_label.config(text=self._t("header_title"))
        self.language_label.config(text=self._t("label_language"))
        self.language_display.set(LANG_CODE_TO_LABEL.get(self.language_code.get(), "English"))
        self.operation_label.config(text=self._t("label_operation"))
        self.select_btn.config(text=self._t("btn_select_input"))

        self.queue_frame.config(text=self._t("queue_group"))
        self.queue_add_btn.config(text=self._t("queue_add"))
        self.queue_remove_btn.config(text=self._t("queue_remove"))
        self.queue_clear_btn.config(text=self._t("queue_clear"))

        self.options_frame.config(text=self._t("options_group"))
        self.page_range_label.config(text=self._t("label_page_range"))
        self.conflict_label.config(text=self._t("label_conflict"))
        self.batch_target_label.config(text=self._t("label_batch_target"))
        self.input_password_label.config(text=self._t("label_input_password"))
        self.output_password_label.config(text=self._t("label_output_password"))
        self.render_dpi_label.config(text=self._t("label_render_dpi"))
        self.jpg_quality_label.config(text=self._t("label_jpg_quality"))
        self.failure_log_check.config(text=self._t("check_failure_log"))
        self.page_help.config(text=self._t("page_help"))

        self.cancel_btn.config(text=self._t("btn_cancel"))
        self._refresh_operation_display_values()
        self._refresh_conflict_display_values()
        self._set_status(self.status_key)
        self._refresh_dynamic_controls()
        self._update_input_label()
        self._update_queue_hint()

    def on_language_changed(self, _event=None):
        selected = self.language_display.get()
        code = LANG_LABEL_TO_CODE.get(selected)
        if not code or code == self.language_code.get():
            return
        self.language_code.set(code)
        self._apply_language()

    def on_operation_selection_changed(self, _event=None):
        selected = self.operation_display.get()
        operation = self.operation_display_to_value.get(selected)
        if not operation:
            return
        self.operation.set(operation)
        self.on_operation_changed()

    def on_conflict_policy_changed(self, _event=None):
        selected = self.conflict_policy_display.get()
        policy = self.conflict_display_to_value.get(selected)
        if policy:
            self.conflict_policy.set(policy)

    def _bind_drag_and_drop(self):
        if not DND_AVAILABLE:
            self.queue_hint_mode = "unavailable"
            self._update_queue_hint()
            return

        try:
            self.queue_listbox.drop_target_register(DND_FILES)
            self.queue_listbox.dnd_bind("<<Drop>>", self.on_drop_files)
            self.queue_hint_mode = "ready"
        except Exception:
            self.queue_hint_mode = "failed"
        self._update_queue_hint()

    def _update_queue_hint(self):
        if self.queue_hint_mode == "unavailable":
            self.queue_hint_label.config(text=self._t("queue_drop_unavailable"))
        elif self.queue_hint_mode == "failed":
            self.queue_hint_label.config(text=self._t("queue_drop_failed"))
        else:
            self.queue_hint_label.config(text=self._t("queue_drop_ready"))

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
        paths = filedialog.askopenfilenames(filetypes=[(self._t("filetype_pdf"), "*.pdf")])
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
                self.input_label.config(text=self._t("input_none_folder"))
            return

        if operation == OP_MERGE:
            if isinstance(resolved_input, tuple) and resolved_input:
                self.input_label.config(text=self._t("input_selected_count", count=len(resolved_input)))
            else:
                self.input_label.config(text=self._t("input_none_pdfs"))
            return

        if isinstance(resolved_input, str) and resolved_input:
            if resolved_input in self.file_queue:
                self.input_label.config(text=self._t("input_from_queue", name=os.path.basename(resolved_input)))
            else:
                self.input_label.config(text=os.path.basename(resolved_input))
        else:
            self.input_label.config(text=self._t("input_none_pdf"))

    def _refresh_dynamic_controls(self):
        operation = self.operation.get()
        if self.is_running:
            return

        if operation == OP_MERGE:
            self.select_btn.config(text=self._t("btn_select_pdfs"))
        elif operation == OP_BATCH:
            self.select_btn.config(text=self._t("btn_select_folder"))
        else:
            self.select_btn.config(text=self._t("btn_select_pdf"))

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
            OP_PDF_TO_PPTX: self._t("btn_convert_pptx"),
            OP_PDF_TO_DOCX: self._t("btn_convert_docx"),
            OP_PDF_TO_PNG: self._t("btn_convert_png"),
            OP_PDF_TO_JPG: self._t("btn_convert_jpg"),
            OP_MERGE: self._t("btn_merge"),
            OP_SPLIT: self._t("btn_split"),
            OP_BATCH: self._t("btn_batch"),
        }
        self.convert_btn.config(text=button_labels.get(operation, self._t("btn_convert")))

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

        self.operation_display.set(self.operation_value_to_display.get(operation, self.operation_display.get()))
        self.progress_var.set(0)
        self._set_status("ready")
        self._refresh_dynamic_controls()
        self._update_input_label()
        self._update_convert_state()

    def select_input(self):
        operation = self.operation.get()

        if operation == OP_MERGE:
            paths = filedialog.askopenfilenames(filetypes=[(self._t("filetype_pdf"), "*.pdf")])
            if paths:
                self.selected_input = tuple(paths)
                self._add_files_to_queue(paths)
                self._update_input_label()
                self._update_convert_state()
            return

        if operation == OP_BATCH:
            folder = filedialog.askdirectory(title=self._t("dialog_input_folder"))
            if folder:
                self.selected_input = folder
                self._update_input_label()
                self._update_convert_state()
            return

        file_path = filedialog.askopenfilename(filetypes=[(self._t("filetype_pdf"), "*.pdf")])
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
            messagebox.showerror(self._t("title_error"), self._t("msg_numeric"))
            return None

        if render_dpi < 72 or render_dpi > 600:
            messagebox.showerror(self._t("title_error"), self._t("msg_dpi"))
            return None
        if jpg_quality < 1 or jpg_quality > 100:
            messagebox.showerror(self._t("title_error"), self._t("msg_jpg"))
            return None

        return render_dpi, jpg_quality

    def start_conversion(self):
        if self.is_running:
            return

        if not self._has_input():
            messagebox.showwarning(self._t("title_warning"), self._t("msg_select_input"))
            return

        numeric_options = self._parse_numeric_options()
        if not numeric_options:
            return
        render_dpi, jpg_quality = numeric_options

        operation = self.operation.get()
        input_data = self._resolved_input_for_operation(operation)

        if operation == OP_MERGE and isinstance(input_data, tuple) and len(input_data) < 2:
            messagebox.showwarning(self._t("title_warning"), self._t("msg_need_two"))
            return

        output_target = self._ask_output_target(operation, input_data)
        if not output_target:
            return

        options = {
            "page_range_text": self.page_range.get().strip(),
            "input_password": self.input_password.get(),
            "output_password": self.output_password.get(),
            "output_conflict_policy": self.conflict_policy.get(),
            "render_dpi": render_dpi,
            "jpg_quality": jpg_quality,
            "batch_target_format": self.batch_target_format.get(),
            "write_failure_log": self.write_failure_log.get(),
        }

        self.cancel_event.clear()
        self._set_controls_running(True)
        self._set_status("working")

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
            self.language_combo.config(state=tk.DISABLED)
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
            self.language_combo.config(state="readonly")
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
        self._set_status("cancelling")

    def _ask_output_target(self, operation: str, input_data: str | tuple[str, ...]) -> str:
        if operation == OP_MERGE:
            return filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[(self._t("filetype_pdf"), "*.pdf")],
                initialfile="merged.pdf",
            )

        if operation in (OP_PDF_TO_PPTX, OP_PDF_TO_DOCX):
            assert isinstance(input_data, str)
            base_name = os.path.splitext(os.path.basename(input_data))[0]
            extension = ".pptx" if operation == OP_PDF_TO_PPTX else ".docx"
            file_type = (
                (self._t("filetype_pptx"), "*.pptx")
                if operation == OP_PDF_TO_PPTX
                else (self._t("filetype_docx"), "*.docx")
            )
            return filedialog.asksaveasfilename(
                defaultextension=extension,
                filetypes=[file_type],
                initialfile=f"{base_name}{extension}",
            )

        if operation in (OP_PDF_TO_PNG, OP_PDF_TO_JPG, OP_SPLIT, OP_BATCH):
            return filedialog.askdirectory(title=self._t("dialog_output_folder"))

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
                success, message = False, self._t("msg_unknown_operation")
        except Exception as exc:
            success, message = False, str(exc)

        self.root.after(0, lambda: self.conversion_finished(success, message))

    def conversion_finished(self, success: bool, message: str):
        self._set_controls_running(False)

        if success:
            self._set_status("done")
            messagebox.showinfo(self._t("title_success"), message)
        else:
            if message.startswith(CANCELLED_MESSAGE):
                self._set_status("cancelled")
                messagebox.showinfo(self._t("title_cancelled"), message)
            else:
                self._set_status("error")
                messagebox.showerror(self._t("title_error"), self._t("msg_error_prefix", message=message))

        self.cancel_event.clear()
        self.progress_var.set(0)
        self._set_status("ready")
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
