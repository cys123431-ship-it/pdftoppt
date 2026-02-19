import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import (
    batch_convert_folder,
    convert_pdf_to_docx,
    convert_pdf_to_images,
    convert_pdf_to_pptx,
    merge_pdfs,
    split_pdf,
)

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


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Converter")
        self.root.geometry("700x420")
        self.root.resizable(False, False)

        self.operation = tk.StringVar(value=OP_PDF_TO_PPTX)
        self.page_range = tk.StringVar(value="")
        self.batch_target_format = tk.StringVar(value="PPTX")
        self.selected_input: str | tuple[str, ...] = ""

        self.style = ttk.Style()
        self.style.configure("TButton", padding=6)

        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(main_frame, text="PDF Converter", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=(0, 15))

        operation_frame = ttk.Frame(main_frame)
        operation_frame.pack(fill=tk.X, pady=(0, 10))
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
        input_frame.pack(fill=tk.X, pady=(0, 10))
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

        page_range_frame = ttk.Frame(main_frame)
        page_range_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(page_range_frame, text="Page range:", width=14).pack(side=tk.LEFT)
        self.page_range_entry = ttk.Entry(page_range_frame, textvariable=self.page_range)
        self.page_range_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        page_help = ttk.Label(
            main_frame,
            text="Format: 1-3,5,8-10 (leave empty for all pages).",
            foreground="gray",
        )
        page_help.pack(anchor="w", pady=(0, 10))

        batch_frame = ttk.Frame(main_frame)
        batch_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(batch_frame, text="Batch target:", width=14).pack(side=tk.LEFT)
        self.batch_combo = ttk.Combobox(
            batch_frame,
            textvariable=self.batch_target_format,
            values=BATCH_TARGET_FORMATS,
            state="readonly",
            width=10,
        )
        self.batch_combo.pack(side=tk.LEFT)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=14)

        self.convert_btn = ttk.Button(main_frame, text="Convert", command=self.start_conversion, state=tk.DISABLED)
        self.convert_btn.pack(fill=tk.X)

        self.status_label = ttk.Label(main_frame, text="Ready", foreground="gray")
        self.status_label.pack(pady=(10, 0))

        self.on_operation_changed()

    def _input_placeholder(self) -> str:
        operation = self.operation.get()
        if operation == OP_MERGE:
            return "No PDF files selected"
        if operation == OP_BATCH:
            return "No input folder selected"
        return "No PDF selected"

    def _has_input(self) -> bool:
        if self.operation.get() == OP_MERGE:
            return isinstance(self.selected_input, tuple) and len(self.selected_input) > 0
        return isinstance(self.selected_input, str) and bool(self.selected_input)

    def on_operation_changed(self, _event=None):
        operation = self.operation.get()
        self.selected_input = ()
        self.input_label.config(text=self._input_placeholder())
        self.progress_var.set(0)
        self.status_label.config(text="Ready", foreground="gray")
        self.convert_btn.config(state=tk.DISABLED)

        if operation == OP_MERGE:
            self.select_btn.config(text="Select PDFs")
            self.page_range_entry.config(state=tk.DISABLED)
        elif operation == OP_BATCH:
            self.select_btn.config(text="Select Folder")
            self.page_range_entry.config(state=tk.NORMAL)
        else:
            self.select_btn.config(text="Select PDF")
            self.page_range_entry.config(state=tk.NORMAL)

        if operation == OP_BATCH:
            self.batch_combo.config(state="readonly")
        else:
            self.batch_combo.config(state=tk.DISABLED)

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

    def select_input(self):
        operation = self.operation.get()

        if operation == OP_MERGE:
            paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
            if paths:
                self.selected_input = tuple(paths)
                self.input_label.config(text=f"{len(paths)} PDF files selected")
                self.convert_btn.config(state=tk.NORMAL)
            return

        if operation == OP_BATCH:
            folder = filedialog.askdirectory(title="Select input folder")
            if folder:
                self.selected_input = folder
                self.input_label.config(text=folder)
                self.convert_btn.config(state=tk.NORMAL)
            return

        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.selected_input = file_path
            self.input_label.config(text=os.path.basename(file_path))
            self.convert_btn.config(state=tk.NORMAL)

    def start_conversion(self):
        if not self._has_input():
            return

        operation = self.operation.get()
        page_range_text = self.page_range.get().strip()
        input_data = self.selected_input

        if operation == OP_MERGE and isinstance(input_data, tuple) and len(input_data) < 2:
            messagebox.showwarning("Warning", "Please select at least 2 PDF files to merge.")
            return

        output_target = self._ask_output_target(operation, input_data)
        if not output_target:
            return

        self.convert_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        self.operation_combo.config(state=tk.DISABLED)
        self.page_range_entry.config(state=tk.DISABLED)
        self.batch_combo.config(state=tk.DISABLED)
        self.status_label.config(text="Working...", foreground="blue")

        thread = threading.Thread(
            target=self.run_conversion,
            args=(operation, input_data, output_target, page_range_text),
            daemon=True,
        )
        thread.start()

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
        page_range_text: str,
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
                    page_range_text,
                )
            elif operation == OP_PDF_TO_DOCX:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_docx(
                    input_data,
                    output_target,
                    update_progress,
                    page_range_text,
                )
            elif operation == OP_PDF_TO_PNG:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_images(
                    input_data,
                    output_target,
                    "png",
                    144,
                    update_progress,
                    page_range_text,
                )
            elif operation == OP_PDF_TO_JPG:
                assert isinstance(input_data, str)
                success, message = convert_pdf_to_images(
                    input_data,
                    output_target,
                    "jpg",
                    144,
                    update_progress,
                    page_range_text,
                )
            elif operation == OP_MERGE:
                assert isinstance(input_data, tuple)
                success, message = merge_pdfs(input_data, output_target, update_progress)
            elif operation == OP_SPLIT:
                assert isinstance(input_data, str)
                success, message = split_pdf(
                    input_data,
                    output_target,
                    update_progress,
                    page_range_text,
                )
            elif operation == OP_BATCH:
                assert isinstance(input_data, str)
                success, message = batch_convert_folder(
                    input_data,
                    output_target,
                    self.batch_target_format.get(),
                    update_progress,
                    page_range_text,
                )
            else:
                success, message = False, "Unknown operation."
        except Exception as exc:
            success, message = False, str(exc)

        self.root.after(0, lambda: self.conversion_finished(success, message))

    def conversion_finished(self, success: bool, message: str):
        self.select_btn.config(state=tk.NORMAL)
        self.operation_combo.config(state="readonly")
        if self.operation.get() == OP_MERGE:
            self.page_range_entry.config(state=tk.DISABLED)
        else:
            self.page_range_entry.config(state=tk.NORMAL)
        if self.operation.get() == OP_BATCH:
            self.batch_combo.config(state="readonly")
        else:
            self.batch_combo.config(state=tk.DISABLED)

        if self._has_input():
            self.convert_btn.config(state=tk.NORMAL)
        else:
            self.convert_btn.config(state=tk.DISABLED)

        if success:
            self.status_label.config(text="Done", foreground="green")
            messagebox.showinfo("Success", message)
        else:
            self.status_label.config(text="Error", foreground="red")
            messagebox.showerror("Error", f"An error occurred:\n{message}")

        self.progress_var.set(0)
        self.status_label.config(text="Ready", foreground="gray")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
