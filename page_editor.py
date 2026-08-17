import base64
import os
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Sequence

import fitz

from converter import (
    CONFLICT_OVERWRITE,
    _atomic_replace,
    _make_staging_path,
    _open_pdf_document,
    _resolve_output_path,
    _safe_remove,
    _save_pdf_document,
)


@dataclass
class PagePlanEntry:
    source_index: int
    rotation_delta: int = 0


def apply_page_edits(
    pdf_path: str,
    output_path: str,
    page_plan: Sequence[PagePlanEntry | tuple[int, int]],
    input_password: str = "",
    output_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
) -> tuple[bool, str, str]:
    """Create an edited PDF from an ordered page plan using an atomic save."""
    if not page_plan:
        return False, "At least one page must remain in the PDF.", ""

    source_doc = None
    output_doc = None
    staged_path: str | None = None
    try:
        resolved_path, skipped = _resolve_output_path(output_path, output_conflict_policy)
        if skipped:
            return True, f"Skipped existing file: {output_path}", output_path

        source_doc = _open_pdf_document(pdf_path, input_password)
        normalized: list[tuple[int, int]] = []
        for item in page_plan:
            if isinstance(item, PagePlanEntry):
                page_index, rotation_delta = item.source_index, item.rotation_delta
            else:
                page_index, rotation_delta = int(item[0]), int(item[1])
            if page_index < 0 or page_index >= len(source_doc):
                raise ValueError(f"Page index out of bounds: {page_index}")
            if rotation_delta % 90 != 0:
                raise ValueError("Page rotation must be a multiple of 90 degrees.")
            normalized.append((page_index, rotation_delta % 360))

        output_doc = fitz.open()
        for page_index, rotation_delta in normalized:
            output_doc.insert_pdf(source_doc, from_page=page_index, to_page=page_index)
            output_page = output_doc[-1]
            output_page.set_rotation((output_page.rotation + rotation_delta) % 360)

        staged_path = _make_staging_path(resolved_path, suffix=".pdf")
        _save_pdf_document(output_doc, staged_path, output_password)

        output_doc.close()
        output_doc = None
        source_doc.close()
        source_doc = None

        _atomic_replace(staged_path, resolved_path)
        staged_path = None
        same_path = os.path.normcase(os.path.abspath(resolved_path)) == os.path.normcase(os.path.abspath(output_path))
        note = "" if same_path else f" Saved as: {resolved_path}"
        return True, f"PDF edit saved successfully!{note}", resolved_path
    except Exception as exc:
        return False, str(exc), ""
    finally:
        _safe_remove(staged_path)
        if output_doc is not None:
            output_doc.close()
        if source_doc is not None:
            source_doc.close()


class PageEditorDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        pdf_path: str,
        input_password: str = "",
        output_password: str = "",
        output_conflict_policy: str = CONFLICT_OVERWRITE,
        language: str = "ko",
        initial_dir: str = "",
        on_saved: Callable[[str], None] | None = None,
    ):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.input_password = input_password
        self.output_password = output_password
        self.output_conflict_policy = output_conflict_policy
        self.language = language
        self.initial_dir = initial_dir
        self.on_saved = on_saved
        self.doc = _open_pdf_document(pdf_path, input_password)
        self.plan = [PagePlanEntry(index) for index in range(len(self.doc))]
        self.preview_image: tk.PhotoImage | None = None
        self.drag_index: int | None = None

        self.title(self._t("title"))
        self.geometry("980x720")
        self.minsize(820, 600)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self._close)

        self._build_ui()
        self._refresh_list(select_index=0)
        self.grab_set()

    def _t(self, key: str) -> str:
        strings = {
            "ko": {
                "title": "PDF 미리보기 / 페이지 편집",
                "pages": "페이지 (드래그해서 순서 변경)",
                "preview": "미리보기",
                "up": "위로",
                "down": "아래로",
                "delete": "삭제",
                "rotate_left": "왼쪽 90°",
                "rotate_right": "오른쪽 90°",
                "save": "다른 이름으로 저장",
                "close": "닫기",
                "page": "페이지",
                "source": "원본",
                "rotation": "회전",
                "last_page": "PDF에는 최소 1페이지가 남아 있어야 합니다.",
                "save_title": "편집된 PDF 저장",
                "saved": "저장 완료",
                "saved_message": "편집된 PDF를 저장했습니다.",
                "skipped_message": "같은 이름의 파일이 이미 있어 저장을 건너뛰었습니다.",
                "error": "오류",
            },
            "en": {
                "title": "PDF Preview / Page Editor",
                "pages": "Pages (drag to reorder)",
                "preview": "Preview",
                "up": "Move Up",
                "down": "Move Down",
                "delete": "Delete",
                "rotate_left": "Rotate Left 90°",
                "rotate_right": "Rotate Right 90°",
                "save": "Save As",
                "close": "Close",
                "page": "Page",
                "source": "source",
                "rotation": "rotation",
                "last_page": "At least one page must remain in the PDF.",
                "save_title": "Save edited PDF",
                "saved": "Saved",
                "saved_message": "Edited PDF saved successfully.",
                "skipped_message": "The existing file was skipped.",
                "error": "Error",
            },
        }
        return strings.get(self.language, strings["en"]).get(key, key)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(1, weight=1)

        ttk.Label(root, text=self._t("pages"), font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6)
        )
        ttk.Label(root, text=self._t("preview"), font=("Segoe UI", 10, "bold")).grid(
            row=0, column=1, sticky="w", pady=(0, 6)
        )

        list_frame = ttk.Frame(root)
        list_frame.grid(row=1, column=0, sticky="ns", padx=(0, 10))
        self.listbox = tk.Listbox(list_frame, width=30, selectmode=tk.SINGLE)
        self.listbox.pack(side=tk.LEFT, fill=tk.Y)
        scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scroll.set)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        self.listbox.bind("<ButtonPress-1>", self._on_drag_start)
        self.listbox.bind("<B1-Motion>", self._on_drag_motion)

        preview_frame = ttk.Frame(root, relief="sunken", borderwidth=1)
        preview_frame.grid(row=1, column=1, sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.preview_label = ttk.Label(preview_frame, anchor="center")
        self.preview_label.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        controls = ttk.Frame(root)
        controls.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        for key, command in (
            ("up", self._move_up),
            ("down", self._move_down),
            ("delete", self._delete),
            ("rotate_left", lambda: self._rotate(-90)),
            ("rotate_right", lambda: self._rotate(90)),
        ):
            ttk.Button(controls, text=self._t(key), command=command).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(controls, text=self._t("close"), command=self._close).pack(side=tk.RIGHT)
        ttk.Button(controls, text=self._t("save"), command=self._save).pack(side=tk.RIGHT, padx=(0, 6))

    def _selected_index(self) -> int | None:
        selection = self.listbox.curselection()
        if not selection:
            return None
        return int(selection[0])

    def _label_for(self, position: int, entry: PagePlanEntry) -> str:
        rotation = entry.rotation_delta % 360
        return f"{self._t('page')} {position + 1}  ({self._t('source')} {entry.source_index + 1}, {self._t('rotation')} {rotation}°)"

    def _refresh_list(self, select_index: int | None = None) -> None:
        self.listbox.delete(0, tk.END)
        for position, entry in enumerate(self.plan):
            self.listbox.insert(tk.END, self._label_for(position, entry))
        if self.plan:
            index = 0 if select_index is None else max(0, min(select_index, len(self.plan) - 1))
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(index)
            self.listbox.activate(index)
            self.listbox.see(index)
            self._render_preview(index)

    def _on_select(self, _event=None) -> None:
        index = self._selected_index()
        if index is not None:
            self._render_preview(index)

    def _render_preview(self, plan_index: int) -> None:
        if plan_index < 0 or plan_index >= len(self.plan):
            return
        entry = self.plan[plan_index]
        page = self.doc[entry.source_index]
        rect = page.rect
        width, height = float(rect.width), float(rect.height)
        if entry.rotation_delta % 180:
            width, height = height, width
        max_width, max_height = 650.0, 520.0
        scale = min(max_width / max(width, 1), max_height / max(height, 1), 1.5)
        matrix = fitz.Matrix(scale, scale).prerotate(entry.rotation_delta)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        encoded = base64.b64encode(pix.tobytes("png")).decode("ascii")
        self.preview_image = tk.PhotoImage(data=encoded)
        self.preview_label.configure(image=self.preview_image)

    def _move_up(self) -> None:
        index = self._selected_index()
        if index is None or index <= 0:
            return
        self.plan[index - 1], self.plan[index] = self.plan[index], self.plan[index - 1]
        self._refresh_list(index - 1)

    def _move_down(self) -> None:
        index = self._selected_index()
        if index is None or index >= len(self.plan) - 1:
            return
        self.plan[index + 1], self.plan[index] = self.plan[index], self.plan[index + 1]
        self._refresh_list(index + 1)

    def _delete(self) -> None:
        index = self._selected_index()
        if index is None:
            return
        if len(self.plan) <= 1:
            messagebox.showwarning(self._t("title"), self._t("last_page"), parent=self)
            return
        self.plan.pop(index)
        self._refresh_list(min(index, len(self.plan) - 1))

    def _rotate(self, delta: int) -> None:
        index = self._selected_index()
        if index is None:
            return
        self.plan[index].rotation_delta = (self.plan[index].rotation_delta + delta) % 360
        self._refresh_list(index)

    def _on_drag_start(self, event) -> None:
        if self.listbox.size() == 0:
            self.drag_index = None
            return
        self.drag_index = self.listbox.nearest(event.y)

    def _on_drag_motion(self, event) -> None:
        if self.drag_index is None or self.listbox.size() == 0:
            return
        target = self.listbox.nearest(event.y)
        if target == self.drag_index or target < 0 or target >= len(self.plan):
            return
        item = self.plan.pop(self.drag_index)
        self.plan.insert(target, item)
        self.drag_index = target
        self._refresh_list(target)

    def _save(self) -> None:
        base = os.path.splitext(os.path.basename(self.pdf_path))[0]
        initial_dir = self.initial_dir or os.path.dirname(self.pdf_path)
        target = filedialog.asksaveasfilename(
            parent=self,
            title=self._t("save_title"),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialdir=initial_dir,
            initialfile=f"{base}_edited.pdf",
        )
        if not target:
            return

        success, message, saved_path = apply_page_edits(
            self.pdf_path,
            target,
            self.plan,
            input_password=self.input_password,
            output_password=self.output_password,
            output_conflict_policy=self.output_conflict_policy,
        )
        if success:
            display_message = self._t("skipped_message") if message.startswith("Skipped") else self._t("saved_message")
            if saved_path:
                display_message = f"{display_message}\n{saved_path}"
            messagebox.showinfo(self._t("saved"), display_message, parent=self)
            if self.on_saved and saved_path:
                self.on_saved(saved_path)
        else:
            messagebox.showerror(self._t("error"), message, parent=self)

    def _close(self) -> None:
        try:
            self.grab_release()
        except tk.TclError:
            pass
        if self.doc is not None:
            self.doc.close()
            self.doc = None
        self.destroy()
