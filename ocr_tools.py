import os
import shutil
import subprocess
import sys
import tempfile
from typing import Optional

import fitz
from docx import Document

from converter import (
    CANCELLED_MESSAGE,
    CONFLICT_OVERWRITE,
    ProgressCallback,
    _atomic_replace,
    _build_output_note,
    _is_cancelled,
    _make_staging_path,
    _open_pdf_document,
    _resolve_output_path,
    _safe_remove,
    _save_pdf_document,
    _set_progress,
    parse_page_range,
)

OCR_MODE_AUTO = "auto"
OCR_MODE_FORCE = "force"
VALID_OCR_MODES = {OCR_MODE_AUTO, OCR_MODE_FORCE}
OCR_LANGUAGES = {"eng", "kor", "kor+eng"}


def _subprocess_creationflags() -> int:
    return int(getattr(subprocess, "CREATE_NO_WINDOW", 0))


def _runtime_root() -> str:
    frozen_root = getattr(sys, "_MEIPASS", None)
    if frozen_root:
        return str(frozen_root)
    return os.path.dirname(os.path.abspath(__file__))


def find_tesseract_executable() -> str:
    """Find bundled or system-installed Tesseract without requiring internet access."""
    candidates: list[str] = []
    explicit = os.environ.get("TESSERACT_CMD", "").strip()
    if explicit:
        candidates.append(explicit)

    candidates.append(os.path.join(_runtime_root(), "tesseract", "tesseract.exe"))
    which = shutil.which("tesseract")
    if which:
        candidates.append(which)

    program_files = [
        os.environ.get("ProgramFiles", ""),
        os.environ.get("ProgramFiles(x86)", ""),
    ]
    for root in program_files:
        if root:
            candidates.append(os.path.join(root, "Tesseract-OCR", "tesseract.exe"))

    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            return os.path.abspath(candidate)
    raise RuntimeError(
        "Tesseract OCR engine was not found. Use the release EXE with bundled OCR, "
        "or install Tesseract 5 and set TESSERACT_CMD."
    )


def find_tessdata_directory(tesseract_executable: Optional[str] = None) -> str:
    executable = tesseract_executable or find_tesseract_executable()
    candidates: list[str] = []

    explicit = os.environ.get("TESSDATA_PREFIX", "").strip()
    if explicit:
        candidates.extend([explicit, os.path.join(explicit, "tessdata")])

    candidates.append(os.path.join(os.path.dirname(executable), "tessdata"))
    candidates.append(os.path.join(_runtime_root(), "tesseract", "tessdata"))

    for candidate in candidates:
        if candidate and os.path.isdir(candidate):
            return os.path.abspath(candidate)
    raise RuntimeError("Tesseract tessdata directory was not found.")


def _language_parts(language: str) -> list[str]:
    normalized = language.strip().lower()
    if normalized not in OCR_LANGUAGES:
        raise ValueError(f"Unsupported OCR language: {language}")
    return [part for part in normalized.split("+") if part]


def ensure_ocr_ready(language: str) -> tuple[str, str]:
    executable = find_tesseract_executable()
    tessdata_dir = find_tessdata_directory(executable)
    missing = [
        lang
        for lang in _language_parts(language)
        if not os.path.isfile(os.path.join(tessdata_dir, f"{lang}.traineddata"))
    ]
    if missing:
        raise RuntimeError(f"Missing Tesseract language data: {', '.join(missing)}")
    return executable, tessdata_dir


def list_ocr_languages() -> set[str]:
    executable = find_tesseract_executable()
    tessdata_dir = find_tessdata_directory(executable)
    command = [executable, "--tessdata-dir", tessdata_dir, "--list-langs"]
    completed = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
        creationflags=_subprocess_creationflags(),
    )
    if completed.returncode != 0:
        return set()
    text = completed.stdout.decode("utf-8", errors="replace")
    return {
        line.strip()
        for line in text.splitlines()
        if line.strip() and not line.lower().startswith("list of available languages")
    }


def page_needs_ocr(page: fitz.Page, ocr_mode: str = OCR_MODE_AUTO, minimum_text_chars: int = 10) -> bool:
    if ocr_mode not in VALID_OCR_MODES:
        raise ValueError(f"Unsupported OCR mode: {ocr_mode}")
    if ocr_mode == OCR_MODE_FORCE:
        return True
    existing = "".join(page.get_text("text").split())
    return len(existing) < minimum_text_chars


def _render_page_to_temp_png(page: fitz.Page, dpi: int) -> str:
    descriptor, path = tempfile.mkstemp(prefix="pdfconverter-ocr-", suffix=".png")
    os.close(descriptor)
    try:
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        pix.save(path)
        return path
    except Exception:
        _safe_remove(path)
        raise


def _run_tesseract(
    image_path: str,
    language: str,
    output_format: str,
    dpi: int,
) -> bytes:
    executable, tessdata_dir = ensure_ocr_ready(language)
    command = [
        executable,
        image_path,
        "stdout",
        "--tessdata-dir",
        tessdata_dir,
        "--oem",
        "1",
        "--psm",
        "3",
        "--dpi",
        str(dpi),
        "-l",
        language,
    ]
    if output_format == "pdf":
        command.append("pdf")

    completed = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
        creationflags=_subprocess_creationflags(),
    )
    if completed.returncode != 0:
        error = completed.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(error or f"Tesseract failed with exit code {completed.returncode}.")
    return completed.stdout


def _ocr_page_text(page: fitz.Page, language: str, dpi: int) -> str:
    temp_image: str | None = None
    try:
        temp_image = _render_page_to_temp_png(page, dpi)
        data = _run_tesseract(temp_image, language, "txt", dpi)
        return data.decode("utf-8", errors="replace").strip()
    finally:
        _safe_remove(temp_image)


def _collect_page_text(
    page: fitz.Page,
    language: str,
    ocr_mode: str,
    dpi: int,
) -> tuple[str, bool]:
    if not page_needs_ocr(page, ocr_mode):
        return page.get_text("text").strip(), False
    return _ocr_page_text(page, language, dpi), True


def ocr_pdf_to_searchable_pdf(
    pdf_path: str,
    output_pdf_path: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
    input_password: str = "",
    output_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    ocr_language: str = "kor+eng",
    ocr_mode: str = OCR_MODE_AUTO,
    ocr_dpi: int = 300,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    """Create a searchable PDF, OCRing only image-only pages in automatic mode."""
    if ocr_dpi < 72 or ocr_dpi > 600:
        return False, "OCR DPI must be between 72 and 600."
    if ocr_mode not in VALID_OCR_MODES:
        return False, f"Unsupported OCR mode: {ocr_mode}"

    source_doc = None
    output_doc = None
    staged_path: str | None = None
    try:
        ensure_ocr_ready(ocr_language)
        resolved_path, skipped = _resolve_output_path(output_pdf_path, output_conflict_policy)
        if skipped:
            return True, f"Skipped existing file: {output_pdf_path}"
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        source_doc = _open_pdf_document(pdf_path, input_password)
        selected_pages = parse_page_range(page_range_text, len(source_doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        output_doc = fitz.open()
        ocr_count = 0
        total = len(selected_pages)
        for position, page_index in enumerate(selected_pages):
            if _is_cancelled(cancel_event):
                return False, CANCELLED_MESSAGE
            page = source_doc[page_index]
            if page_needs_ocr(page, ocr_mode):
                temp_image: str | None = None
                ocr_doc = None
                try:
                    temp_image = _render_page_to_temp_png(page, ocr_dpi)
                    pdf_bytes = _run_tesseract(temp_image, ocr_language, "pdf", ocr_dpi)
                    ocr_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    output_doc.insert_pdf(ocr_doc)
                    ocr_count += 1
                finally:
                    if ocr_doc is not None:
                        ocr_doc.close()
                    _safe_remove(temp_image)
            else:
                output_doc.insert_pdf(source_doc, from_page=page_index, to_page=page_index)
            _set_progress(progress_callback, ((position + 1) / total) * 90)

        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        staged_path = _make_staging_path(resolved_path, suffix=".pdf")
        _save_pdf_document(output_doc, staged_path, output_password)
        output_doc.close()
        output_doc = None
        source_doc.close()
        source_doc = None
        _atomic_replace(staged_path, resolved_path)
        staged_path = None
        _set_progress(progress_callback, 100)
        return (
            True,
            f"OCR PDF completed! OCR applied to {ocr_count}/{total} pages."
            f"{_build_output_note(output_pdf_path, resolved_path)}",
        )
    except Exception as exc:
        return False, str(exc)
    finally:
        _safe_remove(staged_path)
        if output_doc is not None:
            output_doc.close()
        if source_doc is not None:
            source_doc.close()


def ocr_pdf_to_text(
    pdf_path: str,
    output_text_path: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
    input_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    ocr_language: str = "kor+eng",
    ocr_mode: str = OCR_MODE_AUTO,
    ocr_dpi: int = 300,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    if ocr_dpi < 72 or ocr_dpi > 600:
        return False, "OCR DPI must be between 72 and 600."
    source_doc = None
    staged_path: str | None = None
    try:
        ensure_ocr_ready(ocr_language)
        resolved_path, skipped = _resolve_output_path(output_text_path, output_conflict_policy)
        if skipped:
            return True, f"Skipped existing file: {output_text_path}"

        source_doc = _open_pdf_document(pdf_path, input_password)
        selected_pages = parse_page_range(page_range_text, len(source_doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        texts: list[str] = []
        ocr_count = 0
        total = len(selected_pages)
        for position, page_index in enumerate(selected_pages):
            if _is_cancelled(cancel_event):
                return False, CANCELLED_MESSAGE
            text, used_ocr = _collect_page_text(source_doc[page_index], ocr_language, ocr_mode, ocr_dpi)
            texts.append(text)
            ocr_count += int(used_ocr)
            _set_progress(progress_callback, ((position + 1) / total) * 90)

        staged_path = _make_staging_path(resolved_path, suffix=".txt")
        with open(staged_path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write("\n\n\f\n\n".join(texts))
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE
        _atomic_replace(staged_path, resolved_path)
        staged_path = None
        _set_progress(progress_callback, 100)
        return (
            True,
            f"OCR text extraction completed! OCR applied to {ocr_count}/{total} pages."
            f"{_build_output_note(output_text_path, resolved_path)}",
        )
    except Exception as exc:
        return False, str(exc)
    finally:
        _safe_remove(staged_path)
        if source_doc is not None:
            source_doc.close()


def ocr_pdf_to_docx(
    pdf_path: str,
    output_docx_path: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
    input_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    ocr_language: str = "kor+eng",
    ocr_mode: str = OCR_MODE_AUTO,
    ocr_dpi: int = 300,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    if ocr_dpi < 72 or ocr_dpi > 600:
        return False, "OCR DPI must be between 72 and 600."
    source_doc = None
    staged_path: str | None = None
    try:
        ensure_ocr_ready(ocr_language)
        resolved_path, skipped = _resolve_output_path(output_docx_path, output_conflict_policy)
        if skipped:
            return True, f"Skipped existing file: {output_docx_path}"

        source_doc = _open_pdf_document(pdf_path, input_password)
        selected_pages = parse_page_range(page_range_text, len(source_doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        document = Document()
        ocr_count = 0
        total = len(selected_pages)
        for position, page_index in enumerate(selected_pages):
            if _is_cancelled(cancel_event):
                return False, CANCELLED_MESSAGE
            text, used_ocr = _collect_page_text(source_doc[page_index], ocr_language, ocr_mode, ocr_dpi)
            ocr_count += int(used_ocr)
            for paragraph_text in text.splitlines() or [""]:
                document.add_paragraph(paragraph_text)
            if position < total - 1:
                document.add_page_break()
            _set_progress(progress_callback, ((position + 1) / total) * 90)

        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE
        staged_path = _make_staging_path(resolved_path, suffix=".docx")
        document.save(staged_path)
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE
        _atomic_replace(staged_path, resolved_path)
        staged_path = None
        _set_progress(progress_callback, 100)
        return (
            True,
            f"OCR DOCX conversion completed! OCR applied to {ocr_count}/{total} pages."
            f"{_build_output_note(output_docx_path, resolved_path)}",
        )
    except Exception as exc:
        return False, str(exc)
    finally:
        _safe_remove(staged_path)
        if source_doc is not None:
            source_doc.close()
