import multiprocessing
import os
import subprocess
import sys
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
    localize_message,
    merge_pdfs,
    split_pdf,
)
from ocr_tools import (
    OCR_MODE_AUTO,
    OCR_MODE_FORCE,
    ensure_ocr_ready,
    ocr_pdf_to_docx,
    ocr_pdf_to_searchable_pdf,
    ocr_pdf_to_text,
)
from page_editor import PageEditorDialog
from settings_store import DEFAULT_SETTINGS, load_settings, save_settings

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
OP_OCR_PDF = "OCR -> Searchable PDF"
OP_OCR_TXT = "OCR -> TXT"
OP_OCR_DOCX = "OCR -> DOCX"
OP_EDIT = "Preview / Edit PDF"
OP_MERGE = "Merge PDFs"
OP_SPLIT = "Split PDF"
OP_BATCH = "Batch Convert Folder"

OPERATIONS = (
    OP_PDF_TO_PPTX,
    OP_PDF_TO_DOCX,
    OP_PDF_TO_PNG,
    OP_PDF_TO_JPG,
    OP_OCR_PDF,
    OP_OCR_TXT,
    OP_OCR_DOCX,
    OP_EDIT,
    OP_MERGE,
    OP_SPLIT,
    OP_BATCH,
)

BATCH_TARGET_FORMATS = ("PPTX", "DOCX", "PNG", "JPG")
CONFLICT_POLICIES = (CONFLICT_OVERWRITE, CONFLICT_SKIP, CONFLICT_AUTO_RENAME)
OCR_LANGUAGE_VALUES = ("kor+eng", "kor", "eng")
OCR_MODE_VALUES = (OCR_MODE_AUTO, OCR_MODE_FORCE)

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
        "app_title": "PDF 변환기 v1.5",
        "header_title": "PDF 변환기",
        "label_language": "언어:",
        "label_operation": "작업:",
        "queue_group": "파일 큐",
        "queue_add": "PDF 추가",
        "queue_remove": "선택 제거",
        "queue_clear": "큐 비우기",
        "queue_drop_unavailable": "드래그앤드롭을 사용하려면 tkinterdnd2가 필요합니다.",
        "queue_drop_ready": "PDF 파일을 여기로 드래그하면 큐에 추가됩니다.",
        "queue_drop_failed": "드래그앤드롭 초기화에 실패했습니다. PDF 추가 버튼은 사용할 수 있습니다.",
        "options_group": "옵션",
        "label_page_range": "페이지 범위:",
        "label_conflict": "출력 충돌 정책:",
        "label_batch_target": "일괄 대상 형식:",
        "label_input_password": "입력 PDF 비밀번호:",
        "label_output_password": "출력 PDF 비밀번호:",
        "label_render_dpi": "렌더 DPI:",
        "label_jpg_quality": "JPG 품질:",
        "label_ocr_language": "OCR 언어:",
        "label_ocr_mode": "OCR 방식:",
        "label_ocr_dpi": "OCR DPI:",
        "check_failure_log": "배치 실패 로그(CSV) 저장",
        "check_open_output": "완료 후 출력 폴더 열기",
        "page_help": "페이지 범위: 1-3,5,8-10 (비워두면 전체) · OCR 자동은 텍스트가 없는 페이지만 인식합니다.",
        "btn_select_pdf": "PDF 선택",
        "btn_select_pdfs": "PDF들 선택",
        "btn_select_folder": "폴더 선택",
        "btn_convert": "실행",
        "btn_convert_pptx": "PPTX로 변환",
        "btn_convert_docx": "DOCX로 변환",
        "btn_convert_png": "PNG로 변환",
        "btn_convert_jpg": "JPG로 변환",
        "btn_ocr_pdf": "검색 가능한 PDF 만들기",
        "btn_ocr_txt": "OCR 텍스트 추출",
        "btn_ocr_docx": "OCR DOCX 만들기",
        "btn_edit": "미리보기 / 페이지 편집 열기",
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
        "op_ocr_pdf": "OCR -> 검색 가능한 PDF",
        "op_ocr_txt": "OCR -> TXT",
        "op_ocr_docx": "OCR -> DOCX",
        "op_edit": "PDF 미리보기 / 페이지 편집",
        "op_merge": "PDF 병합",
        "op_split": "PDF 분할",
        "op_batch": "폴더 일괄 변환",
        "conflict_overwrite": "덮어쓰기",
        "conflict_skip": "기존 파일 건너뛰기",
        "conflict_auto_rename": "자동 이름 변경",
        "ocr_mode_auto": "자동 (텍스트 없는 페이지만)",
        "ocr_mode_force": "강제 (모든 페이지)",
        "msg_numeric": "현재 작업에 필요한 숫자 옵션을 올바르게 입력하세요.",
        "msg_dpi": "렌더 DPI는 72~600 범위여야 합니다.",
        "msg_jpg": "JPG 품질은 1~100 범위여야 합니다.",
        "msg_ocr_dpi": "OCR DPI는 72~600 범위여야 합니다.",
        "msg_select_input": "먼저 입력을 선택하세요.",
        "msg_need_two": "PDF 병합은 2개 이상의 PDF가 필요합니다.",
        "msg_unknown_operation": "알 수 없는 작업입니다.",
        "msg_error_prefix": "오류가 발생했습니다:\n{message}",
        "msg_close_running": "작업이 진행 중입니다. 작업을 취소하고 프로그램을 종료할까요?",
        "msg_ocr_missing": "OCR 엔진을 사용할 수 없습니다:\n{message}",
        "title_error": "오류",
        "title_success": "성공",
        "title_warning": "경고",
        "title_cancelled": "취소됨",
        "title_confirm": "종료 확인",
        "dialog_input_folder": "입력 폴더 선택",
        "dialog_output_folder": "출력 폴더 선택",
        "filetype_pdf": "PDF 파일",
        "filetype_pptx": "PowerPoint 프레젠테이션",
        "filetype_docx": "Word 문서",
        "filetype_txt": "텍스트 파일",
    },
    "en": {
        "app_title": "PDF Converter v1.5",
        "header_title": "PDF Converter",
        "label_language": "Language:",
        "label_operation": "Operation:",
        "queue_group": "File Queue",
        "queue_add": "Add PDFs",
        "queue_remove": "Remove Selected",
        "queue_clear": "Clear Queue",
        "queue_drop_unavailable": "Drag-and-drop requires tkinterdnd2.",
        "queue_drop_ready": "Drag PDF files here to add them to the queue.",
        "queue_drop_failed": "Drag-and-drop initialization failed. Add PDFs still works.",
        "options_group": "Options",
        "label_page_range": "Page range:",
        "label_conflict": "Conflict policy:",
        "label_batch_target": "Batch target:",
        "label_input_password": "Input PDF password:",
        "label_output_password": "Output PDF password:",
        "label_render_dpi": "Render DPI:",
        "label_jpg_quality": "JPG quality:",
        "label_ocr_language": "OCR language:",
        "label_ocr_mode": "OCR mode:",
        "label_ocr_dpi": "OCR DPI:",
        "check_failure_log": "Save batch failure log (CSV)",
        "check_open_output": "Open output folder when finished",
        "page_help": "Page range: 1-3,5,8-10 (blank = all). Auto OCR processes pages without usable text.",
        "btn_select_pdf": "Select PDF",
        "btn_select_pdfs": "Select PDFs",
        "btn_select_folder": "Select Folder",
        "btn_convert": "Run",
        "btn_convert_pptx": "Convert to PPTX",
        "btn_convert_docx": "Convert to DOCX",
        "btn_convert_png": "Convert to PNG",
        "btn_convert_jpg": "Convert to JPG",
        "btn_ocr_pdf": "Create Searchable PDF",
        "btn_ocr_txt": "Extract OCR Text",
        "btn_ocr_docx": "Create OCR DOCX",
        "btn_edit": "Open Preview / Page Editor",
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
        "op_ocr_pdf": "OCR -> Searchable PDF",
        "op_ocr_txt": "OCR -> TXT",
        "op_ocr_docx": "OCR -> DOCX",
        "op_edit": "PDF Preview / Page Editor",
        "op_merge": "Merge PDFs",
        "op_split": "Split PDF",
        "op_batch": "Batch Convert Folder",
        "conflict_overwrite": "Overwrite",
        "conflict_skip": "Skip Existing",
        "conflict_auto_rename": "Auto Rename",
        "ocr_mode_auto": "Auto (pages without text)",
        "ocr_mode_force": "Force (all pages)",
        "msg_numeric": "Enter valid numeric options required for this operation.",
        "msg_dpi": "Render DPI must be between 72 and 600.",
        "msg_jpg": "JPG quality must be between 1 and 100.",
        "msg_ocr_dpi": "OCR DPI must be between 72 and 600.",
        "msg_select_input": "Select an input first.",
        "msg_need_two": "Please queue/select at least 2 PDF files to merge.",
        "msg_unknown_operation": "Unknown operation.",
        "msg_error_prefix": "An error occurred:\n{message}",
        "msg_close_running": "A job is still running. Cancel it and close the application?",
        "msg_ocr_missing": "OCR engine is unavailable:\n{message}",
        "title_error": "Error",
        "title_success": "Success",
        "title_warning": "Warning",
        "title_cancelled": "Cancelled",
        "title_confirm": "Confirm exit",
        "dialog_input_folder": "Select input folder",
        "dialog_output_folder": "Select output folder",
        "filetype_pdf": "PDF Files",
        "filetype_pptx": "PowerPoint Presentation",
        "filetype_docx": "Word Document",
        "filetype_txt": "Text File",
    },
}

OP_LABEL_KEYS = {
    OP_PDF_TO_PPTX: "op_pdf_to_pptx",
    OP_PDF_TO_DOCX: "op_pdf_to_docx",
    OP_PDF_TO_PNG: "op_pdf_to_png",
    OP_PDF_TO_JPG: "op_pdf_to_jpg",
    OP_OCR_PDF: "op_ocr_pdf",
    OP_OCR_TXT: "op_ocr_txt",
    OP_OCR_DOCX: "op_ocr_docx",
    OP_EDIT: "op_edit",
    OP_MERGE: "op_merge",
    OP_SPLIT: "op_split",
    OP_BATCH: "op_batch",
}

CONFLICT_LABEL_KEYS = {
    CONFLICT_OVERWRITE: "conflict_overwrite",
    CONFLICT_SKIP: "conflict_skip",
    CONFLICT_AUTO_RENAME: "conflict_auto_rename",
}

OCR_MODE_LABEL_KEYS = {
    OCR_MODE_AUTO: "ocr_mode_auto",
    OCR_MODE_FORCE: "ocr_mode_force",
}

OCR_OPERATIONS = {OP_OCR_PDF, OP_OCR_TXT, OP_OCR_DOCX}


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.geometry("1040x900")
        self.root.resizable(True, True)
        self.root.minsize(940, 780)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        saved = load_settings()
        language = saved.get("language", "ko") if saved.get("language") in {"ko", "en"} else "ko"
        conflict = saved.get("conflict_policy", CONFLICT_AUTO_RENAME)
        if conflict not in CONFLICT_POLICIES:
            conflict = CONFLICT_AUTO_RENAME
        ocr_language = saved.get("ocr_language", "kor+eng")
        if ocr_language not in OCR_LANGUAGE_VALUES:
            ocr_language = "kor+eng"
        ocr_mode = saved.get("ocr_mode", OCR_MODE_AUTO)
        if ocr_mode not in OCR_MODE_VALUES:
            ocr_mode = OCR_MODE_AUTO

        self.language_code = tk.StringVar(value=language)
        self.language_display = tk.StringVar(value=LANG_CODE_TO_LABEL[language])
        self.operation = tk.StringVar(value=OP_PDF_TO_PPTX)
        self.operation_display = tk.StringVar(value="")
        self.page_range = tk.StringVar(value="")
        self.batch_target_format = tk.StringVar(value=str(saved.get("batch_target", "PPTX")))
        if self.batch_target_format.get() not in BATCH_TARGET_FORMATS:
            self.batch_target_format.set("PPTX")
        self.conflict_policy = tk.StringVar(value=conflict)
        self.conflict_policy_display = tk.StringVar(value="")
        self.input_password = tk.StringVar(value="")
        self.output_password = tk.StringVar(value="")
        self.render_dpi = tk.StringVar(value=str(saved.get("render_dpi", 144)))
        self.jpg_quality = tk.StringVar(value=str(saved.get("jpg_quality", 90)))
        self.ocr_language = tk.StringVar(value=ocr_language)
        self.ocr_mode = tk.StringVar(value=ocr_mode)
        self.ocr_mode_display = tk.StringVar(value="")
        self.ocr_dpi = tk.StringVar(value=str(saved.get("ocr_dpi", 300)))
        self.write_failure_log = tk.BooleanVar(value=bool(saved.get("write_failure_log", True)))
        self.open_output_folder = tk.BooleanVar(value=bool(saved.get("open_output_folder", True)))
        self.last_output_dir = str(saved.get("last_output_dir", ""))

        self.selected_input = ""
        self.file_queue: list[str] = []
        self.cancel_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.is_running = False
        self.closing = False
        self.status_key = "ready"
        self.queue_hint_mode = "ready"
        self.current_output_target = ""
        self.current_operation = ""

        self.operation_display_to_value: dict[str, str] = {}
        self.operation_value_to_display: dict[str, str] = {}
        self.conflict_display_to_value: dict[str, str] = {}
        self.conflict_value_to_display: dict[str, str] = {}
        self.ocr_mode_display_to_value: dict[str, str] = {}
        self.ocr_mode_value_to_display: dict[str, str] = {}

        self.style = ttk.Style()
        self.style.configure("TButton", padding=6)

        main_frame = ttk.Frame(root, padding="14")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.title_label = ttk.Label(main_frame, font=("Segoe UI", 16, "bold"))
        self.title_label.pack(pady=(0, 10))

        operation_frame = ttk.Frame(main_frame)
        operation_frame.pack(fill=tk.X, pady=(0, 8))
        self.language_label = ttk.Label(operation_frame, width=7)
        self.language_label.pack(side=tk.LEFT)
        self.language_combo = ttk.Combobox(
            operation_frame,
            textvariable=self.language_display,
            values=[label for _, label in LANG_OPTIONS],
            state="readonly",
            width=11,
        )
        self.language_combo.pack(side=tk.LEFT, padx=(0, 12))
        self.language_combo.bind("<<ComboboxSelected>>", self.on_language_changed)

        self.operation_label = ttk.Label(operation_frame, width=8)
        self.operation_label.pack(side=tk.LEFT)
        self.operation_combo = ttk.Combobox(
            operation_frame,
            textvariable=self.operation_display,
            values=[],
            state="readonly",
            width=34,
        )
        self.operation_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.operation_combo.bind("<<ComboboxSelected>>", self.on_operation_selection_changed)

        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 8))
        self.input_label = ttk.Label(input_frame, anchor="w", relief="sunken", padding=(6, 6))
        self.input_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.select_btn = ttk.Button(input_frame, command=self.select_input)
        self.select_btn.pack(side=tk.RIGHT)

        self.queue_frame = ttk.LabelFrame(main_frame)
        self.queue_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        queue_body = ttk.Frame(self.queue_frame, padding="8")
        queue_body.pack(fill=tk.BOTH, expand=True)
        queue_body.columnconfigure(0, weight=1)
        queue_body.rowconfigure(0, weight=1)
        self.queue_listbox = tk.Listbox(queue_body, selectmode=tk.EXTENDED, height=7)
        self.queue_listbox.grid(row=0, column=0, sticky="nsew")
        queue_scroll = ttk.Scrollbar(queue_body, orient=tk.VERTICAL, command=self.queue_listbox.yview)
        queue_scroll.grid(row=0, column=1, sticky="ns")
        self.queue_listbox.config(yscrollcommand=queue_scroll.set)

        queue_buttons = ttk.Frame(queue_body)
        queue_buttons.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.queue_add_btn = ttk.Button(queue_buttons, command=self.add_queue_files)
        self.queue_add_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.queue_remove_btn = ttk.Button(queue_buttons, command=self.remove_queue_selection)
        self.queue_remove_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.queue_clear_btn = ttk.Button(queue_buttons, command=self.clear_queue)
        self.queue_clear_btn.pack(side=tk.LEFT)
        self.queue_hint_label = ttk.Label(queue_body, foreground="gray")
        self.queue_hint_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(7, 0))

        self.options_frame = ttk.LabelFrame(main_frame)
        self.options_frame.pack(fill=tk.X, pady=(0, 8))
        self.options_frame.columnconfigure(1, weight=1)
        self.options_frame.columnconfigure(3, weight=1)

        self.page_range_label = ttk.Label(self.options_frame)
        self.page_range_label.grid(row=0, column=0, sticky="w", padx=8, pady=4)
        self.page_range_entry = ttk.Entry(self.options_frame, textvariable=self.page_range)
        self.page_range_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=4)
        self.conflict_label = ttk.Label(self.options_frame)
        self.conflict_label.grid(row=0, column=2, sticky="w", padx=8, pady=4)
        self.conflict_combo = ttk.Combobox(
            self.options_frame,
            textvariable=self.conflict_policy_display,
            values=[],
            state="readonly",
            width=20,
        )
        self.conflict_combo.grid(row=0, column=3, sticky="ew", padx=8, pady=4)
        self.conflict_combo.bind("<<ComboboxSelected>>", self.on_conflict_policy_changed)

        self.batch_target_label = ttk.Label(self.options_frame)
        self.batch_target_label.grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self.batch_combo = ttk.Combobox(
            self.options_frame,
            textvariable=self.batch_target_format,
            values=BATCH_TARGET_FORMATS,
            state="readonly",
            width=12,
        )
        self.batch_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        self.batch_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_dynamic_controls())
        self.input_password_label = ttk.Label(self.options_frame)
        self.input_password_label.grid(row=1, column=2, sticky="w", padx=8, pady=4)
        self.input_password_entry = ttk.Entry(self.options_frame, textvariable=self.input_password, show="*")
        self.input_password_entry.grid(row=1, column=3, sticky="ew", padx=8, pady=4)

        self.output_password_label = ttk.Label(self.options_frame)
        self.output_password_label.grid(row=2, column=0, sticky="w", padx=8, pady=4)
        self.output_password_entry = ttk.Entry(self.options_frame, textvariable=self.output_password, show="*")
        self.output_password_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=4)
        self.render_dpi_label = ttk.Label(self.options_frame)
        self.render_dpi_label.grid(row=2, column=2, sticky="w", padx=8, pady=4)
        self.render_dpi_spin = ttk.Spinbox(
            self.options_frame, from_=72, to=600, increment=12, textvariable=self.render_dpi
        )
        self.render_dpi_spin.grid(row=2, column=3, sticky="ew", padx=8, pady=4)

        self.jpg_quality_label = ttk.Label(self.options_frame)
        self.jpg_quality_label.grid(row=3, column=0, sticky="w", padx=8, pady=4)
        self.jpg_quality_spin = ttk.Spinbox(
            self.options_frame, from_=1, to=100, increment=1, textvariable=self.jpg_quality
        )
        self.jpg_quality_spin.grid(row=3, column=1, sticky="ew", padx=8, pady=4)
        self.failure_log_check = ttk.Checkbutton(self.options_frame, variable=self.write_failure_log)
        self.failure_log_check.grid(row=3, column=2, columnspan=2, sticky="w", padx=8, pady=4)

        self.ocr_language_label = ttk.Label(self.options_frame)
        self.ocr_language_label.grid(row=4, column=0, sticky="w", padx=8, pady=4)
        self.ocr_language_combo = ttk.Combobox(
            self.options_frame,
            textvariable=self.ocr_language,
            values=OCR_LANGUAGE_VALUES,
            state="readonly",
            width=14,
        )
        self.ocr_language_combo.grid(row=4, column=1, sticky="ew", padx=8, pady=4)
        self.ocr_mode_label = ttk.Label(self.options_frame)
        self.ocr_mode_label.grid(row=4, column=2, sticky="w", padx=8, pady=4)
        self.ocr_mode_combo = ttk.Combobox(
            self.options_frame,
            textvariable=self.ocr_mode_display,
            values=[],
            state="readonly",
            width=24,
        )
        self.ocr_mode_combo.grid(row=4, column=3, sticky="ew", padx=8, pady=4)
        self.ocr_mode_combo.bind("<<ComboboxSelected>>", self.on_ocr_mode_changed)

        self.ocr_dpi_label = ttk.Label(self.options_frame)
        self.ocr_dpi_label.grid(row=5, column=0, sticky="w", padx=8, pady=4)
        self.ocr_dpi_spin = ttk.Spinbox(
            self.options_frame, from_=72, to=600, increment=25, textvariable=self.ocr_dpi
        )
        self.ocr_dpi_spin.grid(row=5, column=1, sticky="ew", padx=8, pady=4)
        self.open_output_check = ttk.Checkbutton(self.options_frame, variable=self.open_output_folder)
        self.open_output_check.grid(row=5, column=2, columnspan=2, sticky="w", padx=8, pady=4)

        self.page_help = ttk.Label(main_frame, foreground="gray")
        self.page_help.pack(anchor="w", pady=(0, 7))
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(3, 8))

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X)
        self.convert_btn = ttk.Button(action_frame, command=self.start_conversion, state=tk.DISABLED)
        self.convert_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.cancel_btn = ttk.Button(action_frame, command=self.cancel_conversion, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT)
        self.status_label = ttk.Label(main_frame, foreground="gray")
        self.status_label.pack(pady=(8, 0), anchor="w")

        self._bind_drag_and_drop()
        self._apply_language()
        self.on_operation_changed()

    def _t(self, key: str, **kwargs) -> str:
        text = I18N.get(self.language_code.get(), I18N["en"]).get(key, key)
        return text.format(**kwargs) if kwargs else text

    def _set_status(self, status_key: str) -> None:
        self.status_key = status_key
        self.status_label.config(
            text=self._t(f"status_{status_key}"),
            foreground=STATUS_COLORS[status_key],
        )

    def _refresh_operation_display_values(self) -> None:
        self.operation_value_to_display = {op: self._t(OP_LABEL_KEYS[op]) for op in OPERATIONS}
        self.operation_display_to_value = {text: op for op, text in self.operation_value_to_display.items()}
        self.operation_combo.config(values=list(self.operation_value_to_display.values()))
        self.operation_display.set(self.operation_value_to_display[self.operation.get()])

    def _refresh_conflict_display_values(self) -> None:
        self.conflict_value_to_display = {
            policy: self._t(CONFLICT_LABEL_KEYS[policy]) for policy in CONFLICT_POLICIES
        }
        self.conflict_display_to_value = {
            text: policy for policy, text in self.conflict_value_to_display.items()
        }
        self.conflict_combo.config(values=list(self.conflict_value_to_display.values()))
        self.conflict_policy_display.set(self.conflict_value_to_display[self.conflict_policy.get()])

    def _refresh_ocr_mode_display_values(self) -> None:
        self.ocr_mode_value_to_display = {
            mode: self._t(OCR_MODE_LABEL_KEYS[mode]) for mode in OCR_MODE_VALUES
        }
        self.ocr_mode_display_to_value = {
            text: mode for mode, text in self.ocr_mode_value_to_display.items()
        }
        self.ocr_mode_combo.config(values=list(self.ocr_mode_value_to_display.values()))
        self.ocr_mode_display.set(self.ocr_mode_value_to_display[self.ocr_mode.get()])

    def _apply_language(self) -> None:
        self.root.title(self._t("app_title"))
        self.title_label.config(text=self._t("header_title"))
        self.language_label.config(text=self._t("label_language"))
        self.language_display.set(LANG_CODE_TO_LABEL.get(self.language_code.get(), "English"))
        self.operation_label.config(text=self._t("label_operation"))
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
        self.ocr_language_label.config(text=self._t("label_ocr_language"))
        self.ocr_mode_label.config(text=self._t("label_ocr_mode"))
        self.ocr_dpi_label.config(text=self._t("label_ocr_dpi"))
        self.failure_log_check.config(text=self._t("check_failure_log"))
        self.open_output_check.config(text=self._t("check_open_output"))
        self.page_help.config(text=self._t("page_help"))
        self.cancel_btn.config(text=self._t("btn_cancel"))
        self._refresh_operation_display_values()
        self._refresh_conflict_display_values()
        self._refresh_ocr_mode_display_values()
        self._set_status(self.status_key)
        self._refresh_dynamic_controls()
        self._update_input_label()
        self._update_queue_hint()

    def on_language_changed(self, _event=None) -> None:
        code = LANG_LABEL_TO_CODE.get(self.language_display.get())
        if code and code != self.language_code.get():
            self.language_code.set(code)
            self._apply_language()
            self._save_settings_safely()

    def on_operation_selection_changed(self, _event=None) -> None:
        operation = self.operation_display_to_value.get(self.operation_display.get())
        if operation:
            self.operation.set(operation)
            self.on_operation_changed()

    def on_conflict_policy_changed(self, _event=None) -> None:
        policy = self.conflict_display_to_value.get(self.conflict_policy_display.get())
        if policy:
            self.conflict_policy.set(policy)
            self._save_settings_safely()

    def on_ocr_mode_changed(self, _event=None) -> None:
        mode = self.ocr_mode_display_to_value.get(self.ocr_mode_display.get())
        if mode:
            self.ocr_mode.set(mode)

    def _bind_drag_and_drop(self) -> None:
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

    def _update_queue_hint(self) -> None:
        key = {
            "unavailable": "queue_drop_unavailable",
            "failed": "queue_drop_failed",
        }.get(self.queue_hint_mode, "queue_drop_ready")
        self.queue_hint_label.config(text=self._t(key))

    def on_drop_files(self, event):
        self._add_files_to_queue(self.root.tk.splitlist(event.data))
        return "break"

    def _add_files_to_queue(self, paths) -> None:
        existing = {_normalized_queue_path(path) for path in self.file_queue}
        added = False
        for path in paths:
            normalized_path = os.path.normpath(path)
            canonical = _normalized_queue_path(normalized_path)
            if not normalized_path.lower().endswith(".pdf"):
                continue
            if not os.path.isfile(normalized_path) or canonical in existing:
                continue
            self.file_queue.append(normalized_path)
            existing.add(canonical)
            added = True
        if added:
            self._refresh_queue_listbox()
            self._update_input_label()
            self._update_convert_state()

    def _refresh_queue_listbox(self) -> None:
        self.queue_listbox.delete(0, tk.END)
        for queued_file in self.file_queue:
            self.queue_listbox.insert(tk.END, queued_file)

    def add_queue_files(self) -> None:
        paths = filedialog.askopenfilenames(filetypes=[(self._t("filetype_pdf"), "*.pdf")])
        if paths:
            self._add_files_to_queue(paths)

    def remove_queue_selection(self) -> None:
        selected_indices = list(self.queue_listbox.curselection())
        for index in reversed(selected_indices):
            self.file_queue.pop(index)
        if selected_indices:
            self._refresh_queue_listbox()
            self._update_input_label()
            self._update_convert_state()

    def clear_queue(self) -> None:
        self.file_queue.clear()
        self._refresh_queue_listbox()
        self._update_input_label()
        self._update_convert_state()

    def _resolved_input_for_operation(self, operation: str | None = None) -> str | tuple[str, ...]:
        target_operation = operation or self.operation.get()
        if target_operation == OP_BATCH:
            return self.selected_input if self.selected_input and os.path.isdir(self.selected_input) else ""
        if target_operation == OP_MERGE:
            return tuple(self.file_queue)
        if self.selected_input and os.path.isfile(self.selected_input):
            return self.selected_input
        if self.file_queue:
            return self.file_queue[0]
        return ""

    def _has_input(self) -> bool:
        return bool(self._resolved_input_for_operation())

    def _update_input_label(self) -> None:
        operation = self.operation.get()
        resolved_input = self._resolved_input_for_operation(operation)
        if operation == OP_BATCH:
            self.input_label.config(
                text=resolved_input if isinstance(resolved_input, str) and resolved_input else self._t("input_none_folder")
            )
            return
        if operation == OP_MERGE:
            self.input_label.config(
                text=self._t("input_selected_count", count=len(resolved_input))
                if isinstance(resolved_input, tuple) and resolved_input
                else self._t("input_none_pdfs")
            )
            return
        if isinstance(resolved_input, str) and resolved_input:
            if any(_normalized_queue_path(resolved_input) == _normalized_queue_path(item) for item in self.file_queue):
                text = self._t("input_from_queue", name=os.path.basename(resolved_input))
            else:
                text = os.path.basename(resolved_input)
            self.input_label.config(text=text)
        else:
            self.input_label.config(text=self._t("input_none_pdf"))

    def _refresh_dynamic_controls(self) -> None:
        if self.is_running:
            return
        operation = self.operation.get()
        if operation == OP_MERGE:
            self.select_btn.config(text=self._t("btn_select_pdfs"))
        elif operation == OP_BATCH:
            self.select_btn.config(text=self._t("btn_select_folder"))
        else:
            self.select_btn.config(text=self._t("btn_select_pdf"))

        self.page_range_entry.config(
            state=tk.DISABLED if operation in (OP_MERGE, OP_EDIT) else tk.NORMAL
        )
        self.batch_combo.config(state="readonly" if operation == OP_BATCH else tk.DISABLED)
        self.output_password_entry.config(
            state=tk.NORMAL if operation in (OP_MERGE, OP_SPLIT, OP_OCR_PDF, OP_EDIT) else tk.DISABLED
        )

        needs_render_dpi = operation in (OP_PDF_TO_PPTX, OP_PDF_TO_PNG, OP_PDF_TO_JPG, OP_BATCH)
        self.render_dpi_spin.config(state=tk.NORMAL if needs_render_dpi else tk.DISABLED)
        needs_jpg = operation == OP_PDF_TO_JPG or (
            operation == OP_BATCH and self.batch_target_format.get() == "JPG"
        )
        self.jpg_quality_spin.config(state=tk.NORMAL if needs_jpg else tk.DISABLED)
        self.failure_log_check.config(state=tk.NORMAL if operation == OP_BATCH else tk.DISABLED)

        ocr_enabled = operation in OCR_OPERATIONS
        self.ocr_language_combo.config(state="readonly" if ocr_enabled else tk.DISABLED)
        self.ocr_mode_combo.config(state="readonly" if ocr_enabled else tk.DISABLED)
        self.ocr_dpi_spin.config(state=tk.NORMAL if ocr_enabled else tk.DISABLED)

        button_labels = {
            OP_PDF_TO_PPTX: self._t("btn_convert_pptx"),
            OP_PDF_TO_DOCX: self._t("btn_convert_docx"),
            OP_PDF_TO_PNG: self._t("btn_convert_png"),
            OP_PDF_TO_JPG: self._t("btn_convert_jpg"),
            OP_OCR_PDF: self._t("btn_ocr_pdf"),
            OP_OCR_TXT: self._t("btn_ocr_txt"),
            OP_OCR_DOCX: self._t("btn_ocr_docx"),
            OP_EDIT: self._t("btn_edit"),
            OP_MERGE: self._t("btn_merge"),
            OP_SPLIT: self._t("btn_split"),
            OP_BATCH: self._t("btn_batch"),
        }
        self.convert_btn.config(text=button_labels.get(operation, self._t("btn_convert")))

    def _update_convert_state(self) -> None:
        self.convert_btn.config(
            state=tk.DISABLED if self.is_running or not self._has_input() else tk.NORMAL
        )

    def on_operation_changed(self, _event=None) -> None:
        operation = self.operation.get()
        if operation == OP_BATCH and self.selected_input and not os.path.isdir(self.selected_input):
            self.selected_input = ""
        elif operation != OP_BATCH and self.selected_input and os.path.isdir(self.selected_input):
            self.selected_input = ""
        self.operation_display.set(
            self.operation_value_to_display.get(operation, self.operation_display.get())
        )
        self.progress_var.set(0)
        self._set_status("ready")
        self._refresh_dynamic_controls()
        self._update_input_label()
        self._update_convert_state()

    def select_input(self) -> None:
        operation = self.operation.get()
        if operation == OP_MERGE:
            paths = filedialog.askopenfilenames(filetypes=[(self._t("filetype_pdf"), "*.pdf")])
            if paths:
                self._add_files_to_queue(paths)
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

    def _parse_numeric_options(self, operation: str) -> tuple[int, int, int] | None:
        render_dpi = int(DEFAULT_SETTINGS["render_dpi"])
        jpg_quality = int(DEFAULT_SETTINGS["jpg_quality"])
        ocr_dpi = int(DEFAULT_SETTINGS["ocr_dpi"])
        batch_target = self.batch_target_format.get()
        needs_dpi = operation in (OP_PDF_TO_PPTX, OP_PDF_TO_PNG, OP_PDF_TO_JPG) or (
            operation == OP_BATCH and batch_target in ("PPTX", "PNG", "JPG")
        )
        needs_jpg = operation == OP_PDF_TO_JPG or (operation == OP_BATCH and batch_target == "JPG")
        needs_ocr_dpi = operation in OCR_OPERATIONS

        try:
            if needs_dpi:
                render_dpi = int(self.render_dpi.get().strip())
            if needs_jpg:
                jpg_quality = int(self.jpg_quality.get().strip())
            if needs_ocr_dpi:
                ocr_dpi = int(self.ocr_dpi.get().strip())
        except ValueError:
            messagebox.showerror(self._t("title_error"), self._t("msg_numeric"))
            return None

        if needs_dpi and not 72 <= render_dpi <= 600:
            messagebox.showerror(self._t("title_error"), self._t("msg_dpi"))
            return None
        if needs_jpg and not 1 <= jpg_quality <= 100:
            messagebox.showerror(self._t("title_error"), self._t("msg_jpg"))
            return None
        if needs_ocr_dpi and not 72 <= ocr_dpi <= 600:
            messagebox.showerror(self._t("title_error"), self._t("msg_ocr_dpi"))
            return None
        return render_dpi, jpg_quality, ocr_dpi

    def start_conversion(self) -> None:
        if self.is_running:
            return
        if not self._has_input():
            messagebox.showwarning(self._t("title_warning"), self._t("msg_select_input"))
            return

        operation = self.operation.get()
        input_data = self._resolved_input_for_operation(operation)
        if operation == OP_MERGE and isinstance(input_data, tuple) and len(input_data) < 2:
            messagebox.showwarning(self._t("title_warning"), self._t("msg_need_two"))
            return

        if operation == OP_EDIT:
            assert isinstance(input_data, str)
            self._save_settings_safely()
            try:
                PageEditorDialog(
                    self.root,
                    input_data,
                    input_password=self.input_password.get(),
                    output_password=self.output_password.get(),
                    output_conflict_policy=self.conflict_policy.get(),
                    language=self.language_code.get(),
                    initial_dir=self._initial_output_dir(input_data),
                    on_saved=self._editor_saved,
                )
            except Exception as exc:
                messagebox.showerror(
                    self._t("title_error"),
                    self._t("msg_error_prefix", message=str(exc)),
                )
            return

        numeric_options = self._parse_numeric_options(operation)
        if numeric_options is None:
            return
        render_dpi, jpg_quality, ocr_dpi = numeric_options

        if operation in OCR_OPERATIONS:
            try:
                ensure_ocr_ready(self.ocr_language.get())
            except Exception as exc:
                messagebox.showerror(
                    self._t("title_error"),
                    self._t("msg_ocr_missing", message=str(exc)),
                )
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
            "ocr_language": self.ocr_language.get(),
            "ocr_mode": self.ocr_mode.get(),
            "ocr_dpi": ocr_dpi,
            "batch_target_format": self.batch_target_format.get(),
            "write_failure_log": self.write_failure_log.get(),
        }
        self.current_output_target = output_target
        self.current_operation = operation
        self._remember_output_dir(operation, output_target)
        self._save_settings_safely()

        self.cancel_event.clear()
        self._set_controls_running(True)
        self._set_status("working")
        self.worker_thread = threading.Thread(
            target=self.run_conversion,
            args=(operation, input_data, output_target, options),
            daemon=True,
        )
        self.worker_thread.start()

    def _set_controls_running(self, is_running: bool) -> None:
        self.is_running = is_running
        if is_running:
            for widget in (
                self.convert_btn,
                self.select_btn,
                self.operation_combo,
                self.language_combo,
                self.page_range_entry,
                self.batch_combo,
                self.conflict_combo,
                self.input_password_entry,
                self.output_password_entry,
                self.render_dpi_spin,
                self.jpg_quality_spin,
                self.failure_log_check,
                self.ocr_language_combo,
                self.ocr_mode_combo,
                self.ocr_dpi_spin,
                self.open_output_check,
                self.queue_add_btn,
                self.queue_remove_btn,
                self.queue_clear_btn,
                self.queue_listbox,
            ):
                widget.config(state=tk.DISABLED)
            self.cancel_btn.config(state=tk.NORMAL)
        else:
            self.cancel_btn.config(state=tk.DISABLED)
            self.select_btn.config(state=tk.NORMAL)
            self.operation_combo.config(state="readonly")
            self.language_combo.config(state="readonly")
            self.conflict_combo.config(state="readonly")
            self.input_password_entry.config(state=tk.NORMAL)
            self.open_output_check.config(state=tk.NORMAL)
            self.queue_add_btn.config(state=tk.NORMAL)
            self.queue_remove_btn.config(state=tk.NORMAL)
            self.queue_clear_btn.config(state=tk.NORMAL)
            self.queue_listbox.config(state=tk.NORMAL)
            self._refresh_dynamic_controls()
            self._update_convert_state()

    def cancel_conversion(self) -> None:
        if self.is_running:
            self.cancel_event.set()
            self._set_status("cancelling")

    def on_close(self) -> None:
        self._save_settings_safely()
        if not self.is_running:
            self.root.destroy()
            return
        if not messagebox.askyesno(self._t("title_confirm"), self._t("msg_close_running")):
            return
        self.closing = True
        self.cancel_event.set()
        self._set_status("cancelling")
        self._wait_for_worker_then_close()

    def _wait_for_worker_then_close(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            self.root.after(100, self._wait_for_worker_then_close)
            return
        self.root.destroy()

    def _initial_output_dir(self, input_path: str = "") -> str:
        if self.last_output_dir and os.path.isdir(self.last_output_dir):
            return self.last_output_dir
        if input_path and os.path.isfile(input_path):
            return os.path.dirname(input_path)
        return ""

    def _ask_output_target(self, operation: str, input_data: str | tuple[str, ...]) -> str:
        input_path = input_data if isinstance(input_data, str) else ""
        initial_dir = self._initial_output_dir(input_path)
        common = {"initialdir": initial_dir} if initial_dir else {}

        if operation == OP_MERGE:
            return filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[(self._t("filetype_pdf"), "*.pdf")],
                initialfile="merged.pdf",
                **common,
            )
        if operation in (OP_PDF_TO_PPTX, OP_PDF_TO_DOCX, OP_OCR_PDF, OP_OCR_TXT, OP_OCR_DOCX):
            assert isinstance(input_data, str)
            base_name = os.path.splitext(os.path.basename(input_data))[0]
            mapping = {
                OP_PDF_TO_PPTX: (".pptx", self._t("filetype_pptx"), ""),
                OP_PDF_TO_DOCX: (".docx", self._t("filetype_docx"), ""),
                OP_OCR_PDF: (".pdf", self._t("filetype_pdf"), "_ocr"),
                OP_OCR_TXT: (".txt", self._t("filetype_txt"), "_ocr"),
                OP_OCR_DOCX: (".docx", self._t("filetype_docx"), "_ocr"),
            }
            extension, label, suffix = mapping[operation]
            return filedialog.asksaveasfilename(
                defaultextension=extension,
                filetypes=[(label, f"*{extension}")],
                initialfile=f"{base_name}{suffix}{extension}",
                **common,
            )
        if operation in (OP_PDF_TO_PNG, OP_PDF_TO_JPG, OP_SPLIT, OP_BATCH):
            return filedialog.askdirectory(title=self._t("dialog_output_folder"), **common)
        return ""

    def run_conversion(
        self,
        operation: str,
        input_data: str | tuple[str, ...],
        output_target: str,
        options: dict,
    ) -> None:
        def update_progress(percent: int) -> None:
            if not self.closing:
                self.root.after(0, lambda: self.progress_var.set(percent))

        try:
            if operation == OP_PDF_TO_PPTX:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_pptx(
                    input_data,
                    output_target,
                    progress_callback=update_progress,
                    page_range_text=options["page_range_text"],
                    input_password=options["input_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    render_dpi=options["render_dpi"],
                    cancel_event=self.cancel_event,
                )
            elif operation == OP_PDF_TO_DOCX:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_docx(
                    input_data,
                    output_target,
                    progress_callback=update_progress,
                    page_range_text=options["page_range_text"],
                    input_password=options["input_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    cancel_event=self.cancel_event,
                )
            elif operation in (OP_PDF_TO_PNG, OP_PDF_TO_JPG):
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_images(
                    input_data,
                    output_target,
                    image_format="png" if operation == OP_PDF_TO_PNG else "jpg",
                    dpi=options["render_dpi"],
                    progress_callback=update_progress,
                    page_range_text=options["page_range_text"],
                    input_password=options["input_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    jpg_quality=options["jpg_quality"],
                    cancel_event=self.cancel_event,
                )
            elif operation == OP_OCR_PDF:
                assert isinstance(input_data, str)
                success, message = ocr_pdf_to_searchable_pdf(
                    input_data,
                    output_target,
                    progress_callback=update_progress,
                    page_range_text=options["page_range_text"],
                    input_password=options["input_password"],
                    output_password=options["output_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    ocr_language=options["ocr_language"],
                    ocr_mode=options["ocr_mode"],
                    ocr_dpi=options["ocr_dpi"],
                    cancel_event=self.cancel_event,
                )
            elif operation == OP_OCR_TXT:
                assert isinstance(input_data, str)
                success, message = ocr_pdf_to_text(
                    input_data,
                    output_target,
                    progress_callback=update_progress,
                    page_range_text=options["page_range_text"],
                    input_password=options["input_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    ocr_language=options["ocr_language"],
                    ocr_mode=options["ocr_mode"],
                    ocr_dpi=options["ocr_dpi"],
                    cancel_event=self.cancel_event,
                )
            elif operation == OP_OCR_DOCX:
                assert isinstance(input_data, str)
                success, message = ocr_pdf_to_docx(
                    input_data,
                    output_target,
                    progress_callback=update_progress,
                    page_range_text=options["page_range_text"],
                    input_password=options["input_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    ocr_language=options["ocr_language"],
                    ocr_mode=options["ocr_mode"],
                    ocr_dpi=options["ocr_dpi"],
                    cancel_event=self.cancel_event,
                )
            elif operation == OP_MERGE:
                assert isinstance(input_data, tuple)
                success, message = merge_pdfs(
                    input_data,
                    output_target,
                    progress_callback=update_progress,
                    input_password=options["input_password"],
                    output_password=options["output_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    cancel_event=self.cancel_event,
                )
            elif operation == OP_SPLIT:
                assert isinstance(input_data, str)
                success, message = split_pdf(
                    input_data,
                    output_target,
                    progress_callback=update_progress,
                    page_range_text=options["page_range_text"],
                    input_password=options["input_password"],
                    output_password=options["output_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    cancel_event=self.cancel_event,
                )
            elif operation == OP_BATCH:
                assert isinstance(input_data, str)
                success, message = batch_convert_folder(
                    input_data,
                    output_target,
                    options["batch_target_format"],
                    progress_callback=update_progress,
                    page_range_text=options["page_range_text"],
                    input_password=options["input_password"],
                    output_password=options["output_password"],
                    output_conflict_policy=options["output_conflict_policy"],
                    render_dpi=options["render_dpi"],
                    jpg_quality=options["jpg_quality"],
                    write_failure_log=options["write_failure_log"],
                    cancel_event=self.cancel_event,
                )
            else:
                success, message = False, self._t("msg_unknown_operation")
        except Exception as exc:
            success, message = False, str(exc)

        if not self.closing:
            self.root.after(0, lambda: self.conversion_finished(success, message))

    def _localize_result(self, message: str) -> str:
        localized = localize_message(message, self.language_code.get())
        if self.language_code.get() != "ko":
            return localized
        replacements = (
            ("OCR PDF completed!", "검색 가능한 OCR PDF 생성이 완료되었습니다!"),
            ("OCR text extraction completed!", "OCR 텍스트 추출이 완료되었습니다!"),
            ("OCR DOCX conversion completed!", "OCR DOCX 변환이 완료되었습니다!"),
            ("OCR applied to", "OCR 적용:"),
            ("pages.", "페이지."),
            ("OCR DPI must be between 72 and 600.", "OCR DPI는 72~600 범위여야 합니다."),
            ("Missing Tesseract language data:", "Tesseract 언어 데이터가 없습니다:"),
            ("Tesseract OCR engine was not found.", "Tesseract OCR 엔진을 찾지 못했습니다."),
        )
        for source, target in replacements:
            localized = localized.replace(source, target)
        return localized

    def conversion_finished(self, success: bool, message: str) -> None:
        self._set_controls_running(False)
        localized_message = self._localize_result(message)
        if success:
            self._set_status("done")
            messagebox.showinfo(self._t("title_success"), localized_message)
            if self.open_output_folder.get() and self.current_output_target:
                self._open_output_location(self.current_operation, self.current_output_target)
        elif message.startswith(CANCELLED_MESSAGE):
            self._set_status("cancelled")
            messagebox.showinfo(self._t("title_cancelled"), localized_message)
        else:
            self._set_status("error")
            messagebox.showerror(
                self._t("title_error"),
                self._t("msg_error_prefix", message=localized_message),
            )

        self.cancel_event.clear()
        self.progress_var.set(0)
        self.current_output_target = ""
        self.current_operation = ""
        self._save_settings_safely()
        self._update_input_label()
        self._update_convert_state()

    def _editor_saved(self, saved_path: str) -> None:
        self.last_output_dir = os.path.dirname(os.path.abspath(saved_path))
        self._set_status("done")
        self._save_settings_safely()
        if self.open_output_folder.get():
            _open_folder(self.last_output_dir)

    def _remember_output_dir(self, operation: str, target: str) -> None:
        if not target:
            return
        if operation in (OP_PDF_TO_PNG, OP_PDF_TO_JPG, OP_SPLIT, OP_BATCH):
            directory = target
        else:
            directory = os.path.dirname(os.path.abspath(target))
        if directory:
            self.last_output_dir = directory

    def _open_output_location(self, operation: str, target: str) -> None:
        if operation in (OP_PDF_TO_PNG, OP_PDF_TO_JPG, OP_SPLIT, OP_BATCH):
            directory = target
        else:
            directory = os.path.dirname(os.path.abspath(target))
        if directory and os.path.isdir(directory):
            _open_folder(directory)

    def _settings_payload(self) -> dict:
        def safe_int(value: str, fallback: int) -> int:
            try:
                return int(value)
            except (TypeError, ValueError):
                return fallback

        return {
            "language": self.language_code.get(),
            "conflict_policy": self.conflict_policy.get(),
            "render_dpi": safe_int(self.render_dpi.get(), int(DEFAULT_SETTINGS["render_dpi"])),
            "jpg_quality": safe_int(self.jpg_quality.get(), int(DEFAULT_SETTINGS["jpg_quality"])),
            "ocr_language": self.ocr_language.get(),
            "ocr_mode": self.ocr_mode.get(),
            "ocr_dpi": safe_int(self.ocr_dpi.get(), int(DEFAULT_SETTINGS["ocr_dpi"])),
            "open_output_folder": bool(self.open_output_folder.get()),
            "batch_target": self.batch_target_format.get(),
            "write_failure_log": bool(self.write_failure_log.get()),
            "last_output_dir": self.last_output_dir,
        }

    def _save_settings_safely(self) -> None:
        try:
            save_settings(self._settings_payload())
        except OSError:
            pass


def _normalized_queue_path(path: str) -> str:
    return os.path.normcase(os.path.realpath(os.path.abspath(path)))


def _open_folder(path: str) -> None:
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except (OSError, subprocess.SubprocessError):
        pass


def create_root() -> tk.Tk:
    if DND_AVAILABLE and TkinterDnD is not None:
        return TkinterDnD.Tk()
    return tk.Tk()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = create_root()
    app = App(root)
    root.mainloop()
