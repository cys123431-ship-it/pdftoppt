import csv
import io
import multiprocessing
import os
import tempfile
import time
from datetime import datetime
from queue import Empty
from typing import Callable, Optional, Sequence

import fitz  # PyMuPDF
from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches

ProgressCallback = Optional[Callable[[int], None]]

CONFLICT_OVERWRITE = "overwrite"
CONFLICT_SKIP = "skip"
CONFLICT_AUTO_RENAME = "auto_rename"
VALID_CONFLICT_POLICIES = {
    CONFLICT_OVERWRITE,
    CONFLICT_SKIP,
    CONFLICT_AUTO_RENAME,
}

CANCELLED_MESSAGE = "Cancelled by user."


def _set_progress(progress_callback: ProgressCallback, value: float) -> None:
    if progress_callback:
        clamped = max(0, min(100, int(value)))
        progress_callback(clamped)


def _is_cancelled(cancel_event: object | None) -> bool:
    return bool(cancel_event and hasattr(cancel_event, "is_set") and cancel_event.is_set())


def _validate_conflict_policy(output_conflict_policy: str) -> None:
    if output_conflict_policy not in VALID_CONFLICT_POLICIES:
        raise ValueError(f"Unsupported conflict policy: {output_conflict_policy}")


def _resolve_output_path(output_path: str, output_conflict_policy: str) -> tuple[str, bool]:
    _validate_conflict_policy(output_conflict_policy)
    if not os.path.exists(output_path):
        return output_path, False

    if output_conflict_policy == CONFLICT_OVERWRITE:
        return output_path, False

    if output_conflict_policy == CONFLICT_SKIP:
        return "", True

    directory, filename = os.path.split(output_path)
    stem, extension = os.path.splitext(filename)
    suffix = 1
    while True:
        candidate = os.path.join(directory, f"{stem}_{suffix}{extension}")
        if not os.path.exists(candidate):
            return candidate, False
        suffix += 1


def _resolve_output_directory(output_dir: str, output_conflict_policy: str) -> tuple[str, bool]:
    _validate_conflict_policy(output_conflict_policy)
    if not os.path.exists(output_dir):
        return output_dir, False

    if not os.path.isdir(output_dir):
        raise ValueError(f"Output directory path is an existing file: {output_dir}")

    if output_conflict_policy == CONFLICT_OVERWRITE:
        return output_dir, False

    if output_conflict_policy == CONFLICT_SKIP:
        return "", True

    suffix = 1
    while True:
        candidate = f"{output_dir}_{suffix}"
        if not os.path.exists(candidate):
            return candidate, False
        suffix += 1


def _safe_remove(path: str | None) -> None:
    if not path:
        return
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


def _normalized_path(path: str) -> str:
    return os.path.normcase(os.path.realpath(os.path.abspath(path)))


def _same_path(first: str, second: str) -> bool:
    return _normalized_path(first) == _normalized_path(second)


def _make_staging_path(final_path: str, suffix: str | None = None) -> str:
    directory = os.path.dirname(os.path.abspath(final_path)) or os.getcwd()
    os.makedirs(directory, exist_ok=True)
    final_suffix = suffix if suffix is not None else os.path.splitext(final_path)[1]
    file_descriptor, temp_path = tempfile.mkstemp(
        prefix=".pdfconverter-",
        suffix=final_suffix,
        dir=directory,
    )
    os.close(file_descriptor)
    os.remove(temp_path)
    return temp_path


def _atomic_replace(staged_path: str, final_path: str) -> None:
    os.replace(staged_path, final_path)


def _commit_staged_outputs(staged_outputs: Sequence[tuple[str, str]]) -> None:
    """Commit a group of staged files and restore previous outputs if commit fails."""
    backups: dict[str, str] = {}
    committed: list[str] = []
    try:
        for _staged_path, final_path in staged_outputs:
            if os.path.isfile(final_path):
                backup_path = _make_staging_path(final_path, suffix=".bak")
                os.replace(final_path, backup_path)
                backups[final_path] = backup_path

        for staged_path, final_path in staged_outputs:
            os.replace(staged_path, final_path)
            committed.append(final_path)

        for backup_path in backups.values():
            _safe_remove(backup_path)
    except Exception:
        for final_path in committed:
            _safe_remove(final_path)
        for final_path, backup_path in backups.items():
            if os.path.exists(backup_path):
                os.replace(backup_path, final_path)
        for staged_path, _final_path in staged_outputs:
            _safe_remove(staged_path)
        raise


def _build_output_note(original_path: str, resolved_path: str) -> str:
    if _same_path(original_path, resolved_path):
        return ""
    return f" Saved as: {resolved_path}"


def _csv_safe(value: object) -> str:
    text = str(value)
    if text.startswith(("=", "+", "-", "@", "\t", "\r")):
        return "'" + text
    return text


def localize_message(message: str, language: str) -> str:
    """Translate converter result messages for the GUI while keeping the core API stable."""
    if language != "ko":
        return message

    exact = {
        CANCELLED_MESSAGE: "사용자가 작업을 취소했습니다.",
        "Render DPI must be greater than 0.": "렌더 DPI는 0보다 커야 합니다.",
        "DPI must be greater than 0.": "DPI는 0보다 커야 합니다.",
        "JPG quality must be between 1 and 100.": "JPG 품질은 1~100 범위여야 합니다.",
        "The PDF has no pages.": "PDF에 페이지가 없습니다.",
        "Select at least 2 PDF files to merge.": "병합할 PDF를 2개 이상 선택하세요.",
        "Input folder does not exist.": "입력 폴더가 존재하지 않습니다.",
        "No PDF files found in input folder.": "입력 폴더에서 PDF 파일을 찾지 못했습니다.",
        "All images were skipped because output files already exist.": "출력 파일이 이미 존재하여 모든 이미지를 건너뛰었습니다.",
        "All split files were skipped because output files already exist.": "출력 파일이 이미 존재하여 모든 분할 파일을 건너뛰었습니다.",
        "This PDF is password-protected. Enter an input password.": "암호로 보호된 PDF입니다. 입력 PDF 비밀번호를 입력하세요.",
        "Invalid password for encrypted PDF.": "암호화된 PDF의 비밀번호가 올바르지 않습니다.",
        "The output PDF cannot overwrite one of the input PDFs.": "출력 PDF를 입력 PDF 중 하나와 같은 경로에 덮어쓸 수 없습니다.",
        "DOCX conversion worker stopped unexpectedly.": "DOCX 변환 작업이 비정상적으로 종료되었습니다.",
    }
    if message in exact:
        return exact[message]

    replacements = (
        ("Conversion successful!", "변환이 완료되었습니다!"),
        ("Merge successful!", "PDF 병합이 완료되었습니다!"),
        ("Batch conversion successful!", "일괄 변환이 완료되었습니다!"),
        ("Skipped existing file:", "기존 파일을 건너뛰었습니다:"),
        ("Skipped existing output directory:", "기존 출력 폴더를 건너뛰었습니다:"),
        ("Saved as:", "저장 위치:"),
        ("Saved ", "저장: "),
        (" images. Skipped ", "개 이미지, 건너뜀 "),
        (" existing files.", "개 기존 파일."),
        (" image files.", "개 이미지 파일."),
        ("Created ", "생성: "),
        (" split files. Skipped ", "개 분할 파일, 건너뜀 "),
        (" split PDF files.", "개 분할 PDF 파일."),
        ("Unsupported image format:", "지원하지 않는 이미지 형식:"),
        ("Unsupported target format:", "지원하지 않는 대상 형식:"),
        ("Completed with errors.", "오류와 함께 작업이 완료되었습니다."),
        ("Completed with skips.", "일부 파일을 건너뛰고 작업이 완료되었습니다."),
        ("Converted ", "변환: "),
        ("Failed:", "실패:"),
        ("Skipped:", "건너뜀:"),
        ("Failure log:", "실패 로그:"),
        (CANCELLED_MESSAGE, "사용자가 작업을 취소했습니다."),
    )
    localized = message
    for source, target in replacements:
        localized = localized.replace(source, target)
    return localized


def parse_page_range(page_range_text: str, total_pages: int) -> list[int]:
    """Parse ranges such as ``1-3,5,8-10`` into zero-based page indices."""
    if total_pages <= 0:
        return []

    if not page_range_text or not page_range_text.strip():
        return list(range(total_pages))

    selected_pages: set[int] = set()
    for raw_token in page_range_text.split(","):
        token = raw_token.strip()
        if not token:
            continue

        if "-" in token:
            parts = [part.strip() for part in token.split("-", 1)]
            if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
                raise ValueError(f"Invalid page range token: {token}")
            start_page = int(parts[0])
            end_page = int(parts[1])
            if start_page > end_page:
                raise ValueError(f"Range start must be <= end: {token}")
            if start_page < 1 or end_page > total_pages:
                raise ValueError(f"Page range out of bounds: {token} (valid: 1-{total_pages})")
            selected_pages.update(range(start_page - 1, end_page))
            continue

        if not token.isdigit():
            raise ValueError(f"Invalid page number: {token}")
        page_number = int(token)
        if page_number < 1 or page_number > total_pages:
            raise ValueError(f"Page out of bounds: {page_number} (valid: 1-{total_pages})")
        selected_pages.add(page_number - 1)

    if not selected_pages:
        raise ValueError("No pages selected.")
    return sorted(selected_pages)


def _open_pdf_document(pdf_path: str, input_password: str = "") -> fitz.Document:
    document = fitz.open(pdf_path)
    if document.needs_pass:
        if not input_password:
            document.close()
            raise ValueError("This PDF is password-protected. Enter an input password.")
        authenticated = document.authenticate(input_password)
        if authenticated == 0:
            document.close()
            raise ValueError("Invalid password for encrypted PDF.")
    return document


def _save_pdf_document(document: fitz.Document, output_pdf_path: str, output_password: str = "") -> None:
    if output_password:
        encryption_mode = getattr(fitz, "PDF_ENCRYPT_AES_256", fitz.PDF_ENCRYPT_AES_128)
        document.save(
            output_pdf_path,
            encryption=encryption_mode,
            owner_pw=output_password,
            user_pw=output_password,
        )
    else:
        document.save(output_pdf_path)


def _create_temp_pdf_with_selected_pages(
    source_doc: fitz.Document,
    page_indices: Sequence[int],
    cancel_event: object | None = None,
) -> str:
    temp_doc = fitz.open()
    temp_pdf_path: str | None = None
    try:
        for page_index in page_indices:
            if _is_cancelled(cancel_event):
                raise RuntimeError(CANCELLED_MESSAGE)
            temp_doc.insert_pdf(source_doc, from_page=page_index, to_page=page_index)

        file_descriptor, temp_pdf_path = tempfile.mkstemp(prefix="pdfconverter-", suffix=".pdf")
        os.close(file_descriptor)
        try:
            os.chmod(temp_pdf_path, 0o600)
        except OSError:
            pass
        temp_doc.save(temp_pdf_path)
        return temp_pdf_path
    except Exception:
        _safe_remove(temp_pdf_path)
        raise
    finally:
        temp_doc.close()


def _write_batch_failure_log(output_dir: str, rows: Sequence[tuple[str, str, str]]) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(output_dir, f"batch_failures_{timestamp}.csv")
    with open(log_path, "w", newline="", encoding="utf-8-sig") as output_file:
        writer = csv.writer(output_file)
        writer.writerow(["file", "status", "message"])
        for file_name, status, message in rows:
            writer.writerow([_csv_safe(file_name), _csv_safe(status), _csv_safe(message)])
    return log_path


def _fit_picture_to_slide(page: fitz.Page, slide_width: int, slide_height: int) -> tuple[int, int, int, int]:
    page_width = float(page.rect.width)
    page_height = float(page.rect.height)
    if page_width <= 0 or page_height <= 0:
        return 0, 0, slide_width, slide_height

    scale = min(slide_width / page_width, slide_height / page_height)
    width = max(1, int(page_width * scale))
    height = max(1, int(page_height * scale))
    left = max(0, (slide_width - width) // 2)
    top = max(0, (slide_height - height) // 2)
    return left, top, width, height


def _docx_convert_worker(temp_pdf_path: str, output_path: str, result_queue: multiprocessing.Queue) -> None:
    converter = None
    try:
        converter = Converter(temp_pdf_path)
        converter.convert(output_path)
        result_queue.put((True, ""))
    except Exception as exc:
        result_queue.put((False, str(exc)))
    finally:
        if converter is not None:
            converter.close()


def _run_docx_worker(
    temp_pdf_path: str,
    output_path: str,
    cancel_event: object | None,
) -> tuple[bool, str]:
    result_queue: multiprocessing.Queue = multiprocessing.Queue()
    process = multiprocessing.Process(
        target=_docx_convert_worker,
        args=(temp_pdf_path, output_path, result_queue),
        daemon=False,
    )
    process.start()
    try:
        while process.is_alive():
            if _is_cancelled(cancel_event):
                process.terminate()
                process.join(timeout=5)
                return False, CANCELLED_MESSAGE
            process.join(timeout=0.1)
            time.sleep(0.01)

        process.join()
        try:
            return result_queue.get_nowait()
        except Empty:
            if process.exitcode == 0:
                return True, ""
            return False, "DOCX conversion worker stopped unexpectedly."
    finally:
        result_queue.close()
        result_queue.join_thread()


def convert_pdf_to_pptx(
    pdf_path: str,
    pptx_path: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
    input_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    render_dpi: int = 144,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    """Convert selected pages to image-backed PPTX slides without distorting page aspect ratio."""
    doc = None
    staged_path: str | None = None
    try:
        if render_dpi <= 0:
            return False, "Render DPI must be greater than 0."
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        resolved_pptx_path, skipped = _resolve_output_path(pptx_path, output_conflict_policy)
        if skipped:
            return True, f"Skipped existing file: {pptx_path}"

        doc = _open_pdf_document(pdf_path, input_password)
        selected_pages = parse_page_range(page_range_text, len(doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        prs = Presentation()
        first_page = doc[selected_pages[0]]
        prs.slide_width = Inches(first_page.rect.width / 72.0)
        prs.slide_height = Inches(first_page.rect.height / 72.0)

        zoom = render_dpi / 72.0
        total_pages = len(selected_pages)
        for index, page_number in enumerate(selected_pages):
            if _is_cancelled(cancel_event):
                return False, CANCELLED_MESSAGE

            page = doc[page_number]
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            image_stream = io.BytesIO(pix.tobytes("png"))
            left, top, width, height = _fit_picture_to_slide(page, prs.slide_width, prs.slide_height)

            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(image_stream, left, top, width=width, height=height)
            _set_progress(progress_callback, ((index + 1) / total_pages) * 95)

        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        staged_path = _make_staging_path(resolved_pptx_path)
        prs.save(staged_path)
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE
        _atomic_replace(staged_path, resolved_pptx_path)
        staged_path = None
        _set_progress(progress_callback, 100)
        return True, f"Conversion successful!{_build_output_note(pptx_path, resolved_pptx_path)}"
    except Exception as exc:
        return False, str(exc)
    finally:
        _safe_remove(staged_path)
        if doc is not None:
            doc.close()


def convert_pdf_to_docx(
    pdf_path: str,
    docx_path: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
    input_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    """Convert selected pages to DOCX in a cancellable child process."""
    source_doc = None
    temp_pdf_path: str | None = None
    staged_docx_path: str | None = None
    try:
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        resolved_docx_path, skipped = _resolve_output_path(docx_path, output_conflict_policy)
        if skipped:
            return True, f"Skipped existing file: {docx_path}"

        source_doc = _open_pdf_document(pdf_path, input_password)
        selected_pages = parse_page_range(page_range_text, len(source_doc))
        if not selected_pages:
            return False, "The PDF has no pages."
        _set_progress(progress_callback, 10)

        temp_pdf_path = _create_temp_pdf_with_selected_pages(source_doc, selected_pages, cancel_event)
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        _set_progress(progress_callback, 30)
        staged_docx_path = _make_staging_path(resolved_docx_path)
        success, message = _run_docx_worker(temp_pdf_path, staged_docx_path, cancel_event)
        if not success:
            return False, message
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        _set_progress(progress_callback, 95)
        _atomic_replace(staged_docx_path, resolved_docx_path)
        staged_docx_path = None
        _set_progress(progress_callback, 100)
        return True, f"Conversion successful!{_build_output_note(docx_path, resolved_docx_path)}"
    except RuntimeError as exc:
        if str(exc) == CANCELLED_MESSAGE:
            return False, CANCELLED_MESSAGE
        return False, str(exc)
    except Exception as exc:
        return False, str(exc)
    finally:
        _safe_remove(staged_docx_path)
        _safe_remove(temp_pdf_path)
        if source_doc is not None:
            source_doc.close()


def convert_pdf_to_images(
    pdf_path: str,
    output_dir: str,
    image_format: str = "png",
    dpi: int = 144,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
    input_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    jpg_quality: int = 90,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    """Convert selected pages to images transactionally; cancelled jobs leave no partial new outputs."""
    doc = None
    staged_outputs: list[tuple[str, str]] = []
    try:
        if dpi <= 0:
            return False, "DPI must be greater than 0."
        if jpg_quality < 1 or jpg_quality > 100:
            return False, "JPG quality must be between 1 and 100."
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        normalized_format = image_format.lower()
        if normalized_format not in {"png", "jpg", "jpeg"}:
            return False, f"Unsupported image format: {image_format}"

        output_extension = "jpg" if normalized_format in {"jpg", "jpeg"} else "png"
        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]

        doc = _open_pdf_document(pdf_path, input_password)
        selected_pages = parse_page_range(page_range_text, len(doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        zoom = dpi / 72.0
        total_pages = len(selected_pages)
        skipped_count = 0

        for index, page_number in enumerate(selected_pages):
            if _is_cancelled(cancel_event):
                return False, CANCELLED_MESSAGE

            output_name = f"{base_name}_p{page_number + 1:03d}.{output_extension}"
            output_path = os.path.join(output_dir, output_name)
            resolved_path, skipped = _resolve_output_path(output_path, output_conflict_policy)
            if skipped:
                skipped_count += 1
                _set_progress(progress_callback, ((index + 1) / total_pages) * 90)
                continue

            page = doc[page_number]
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            image_bytes = (
                pix.tobytes("jpeg", jpg_quality=jpg_quality)
                if output_extension == "jpg"
                else pix.tobytes("png")
            )
            staged_path = _make_staging_path(resolved_path)
            with open(staged_path, "wb") as output_file:
                output_file.write(image_bytes)
            staged_outputs.append((staged_path, resolved_path))
            _set_progress(progress_callback, ((index + 1) / total_pages) * 90)

        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        if staged_outputs:
            _commit_staged_outputs(staged_outputs)
            staged_outputs.clear()
        _set_progress(progress_callback, 100)

        created_count = len(selected_pages) - skipped_count
        if created_count == 0 and skipped_count > 0:
            return True, "All images were skipped because output files already exist."
        if skipped_count > 0:
            return True, f"Saved {created_count} images. Skipped {skipped_count} existing files."
        return True, f"Saved {created_count} image files."
    except Exception as exc:
        return False, str(exc)
    finally:
        for staged_path, _final_path in staged_outputs:
            _safe_remove(staged_path)
        if doc is not None:
            doc.close()


def merge_pdfs(
    input_pdf_paths: Sequence[str],
    output_pdf_path: str,
    progress_callback: ProgressCallback = None,
    input_password: str = "",
    output_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    """Merge PDFs without deleting any input or previous output before a successful save."""
    if len(input_pdf_paths) < 2:
        return False, "Select at least 2 PDF files to merge."

    merged_doc = fitz.open()
    staged_output_path: str | None = None
    try:
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        resolved_output_path, skipped = _resolve_output_path(output_pdf_path, output_conflict_policy)
        if skipped:
            return True, f"Skipped existing file: {output_pdf_path}"
        if any(_same_path(input_path, resolved_output_path) for input_path in input_pdf_paths):
            return False, "The output PDF cannot overwrite one of the input PDFs."

        total_files = len(input_pdf_paths)
        for index, input_path in enumerate(input_pdf_paths):
            if _is_cancelled(cancel_event):
                return False, CANCELLED_MESSAGE
            source_doc = _open_pdf_document(input_path, input_password)
            try:
                merged_doc.insert_pdf(source_doc)
            finally:
                source_doc.close()
            _set_progress(progress_callback, ((index + 1) / total_files) * 90)

        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        staged_output_path = _make_staging_path(resolved_output_path, suffix=".pdf")
        _save_pdf_document(merged_doc, staged_output_path, output_password)
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE
        _atomic_replace(staged_output_path, resolved_output_path)
        staged_output_path = None
        _set_progress(progress_callback, 100)
        return True, f"Merge successful!{_build_output_note(output_pdf_path, resolved_output_path)}"
    except Exception as exc:
        return False, str(exc)
    finally:
        _safe_remove(staged_output_path)
        merged_doc.close()


def split_pdf(
    pdf_path: str,
    output_dir: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
    input_password: str = "",
    output_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    """Split selected pages transactionally; cancellation leaves previous outputs untouched."""
    source_doc = None
    staged_outputs: list[tuple[str, str]] = []
    try:
        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE

        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        source_doc = _open_pdf_document(pdf_path, input_password)
        selected_pages = parse_page_range(page_range_text, len(source_doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        total_pages = len(selected_pages)
        skipped_count = 0
        for index, page_number in enumerate(selected_pages):
            if _is_cancelled(cancel_event):
                return False, CANCELLED_MESSAGE

            output_name = f"{base_name}_p{page_number + 1:03d}.pdf"
            output_path = os.path.join(output_dir, output_name)
            resolved_output_path, skipped = _resolve_output_path(output_path, output_conflict_policy)
            if skipped:
                skipped_count += 1
                _set_progress(progress_callback, ((index + 1) / total_pages) * 90)
                continue

            staged_path = _make_staging_path(resolved_output_path, suffix=".pdf")
            split_doc = fitz.open()
            try:
                split_doc.insert_pdf(source_doc, from_page=page_number, to_page=page_number)
                _save_pdf_document(split_doc, staged_path, output_password)
            finally:
                split_doc.close()
            staged_outputs.append((staged_path, resolved_output_path))
            _set_progress(progress_callback, ((index + 1) / total_pages) * 90)

        if _is_cancelled(cancel_event):
            return False, CANCELLED_MESSAGE
        if staged_outputs:
            _commit_staged_outputs(staged_outputs)
            staged_outputs.clear()
        _set_progress(progress_callback, 100)

        created_count = len(selected_pages) - skipped_count
        if created_count == 0 and skipped_count > 0:
            return True, "All split files were skipped because output files already exist."
        if skipped_count > 0:
            return True, f"Created {created_count} split files. Skipped {skipped_count} existing files."
        return True, f"Created {created_count} split PDF files."
    except Exception as exc:
        return False, str(exc)
    finally:
        for staged_path, _final_path in staged_outputs:
            _safe_remove(staged_path)
        if source_doc is not None:
            source_doc.close()


def batch_convert_folder(
    input_dir: str,
    output_dir: str,
    target_format: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
    input_password: str = "",
    output_password: str = "",
    output_conflict_policy: str = CONFLICT_OVERWRITE,
    render_dpi: int = 144,
    jpg_quality: int = 90,
    write_failure_log: bool = True,
    cancel_event: object | None = None,
) -> tuple[bool, str]:
    """Batch convert all PDFs in a folder to PPTX, DOCX, PNG or JPG."""
    try:
        if not os.path.isdir(input_dir):
            return False, "Input folder does not exist."
        os.makedirs(output_dir, exist_ok=True)

        pdf_files = sorted(
            filename
            for filename in os.listdir(input_dir)
            if filename.lower().endswith(".pdf") and os.path.isfile(os.path.join(input_dir, filename))
        )
        if not pdf_files:
            return False, "No PDF files found in input folder."

        normalized_target = target_format.upper()
        if normalized_target not in {"PPTX", "DOCX", "PNG", "JPG"}:
            return False, f"Unsupported target format: {target_format}"

        total_files = len(pdf_files)
        converted_count = 0
        skipped_count = 0
        failed_count = 0
        failure_rows: list[tuple[str, str, str]] = []
        cancelled = False

        for file_index, filename in enumerate(pdf_files):
            if _is_cancelled(cancel_event):
                cancelled = True
                failure_rows.append((filename, "cancelled", CANCELLED_MESSAGE))
                break

            input_pdf_path = os.path.join(input_dir, filename)
            base_name = os.path.splitext(filename)[0]

            def file_progress(child_percent: int, index: int = file_index) -> None:
                overall = ((index + (child_percent / 100.0)) / total_files) * 100.0
                _set_progress(progress_callback, overall)

            if normalized_target == "PPTX":
                success, message = convert_pdf_to_pptx(
                    input_pdf_path,
                    os.path.join(output_dir, f"{base_name}.pptx"),
                    progress_callback=file_progress,
                    page_range_text=page_range_text,
                    input_password=input_password,
                    output_conflict_policy=output_conflict_policy,
                    render_dpi=render_dpi,
                    cancel_event=cancel_event,
                )
            elif normalized_target == "DOCX":
                success, message = convert_pdf_to_docx(
                    input_pdf_path,
                    os.path.join(output_dir, f"{base_name}.docx"),
                    progress_callback=file_progress,
                    page_range_text=page_range_text,
                    input_password=input_password,
                    output_conflict_policy=output_conflict_policy,
                    cancel_event=cancel_event,
                )
            else:
                output_subdir = os.path.join(output_dir, base_name)
                resolved_subdir, dir_skipped = _resolve_output_directory(output_subdir, output_conflict_policy)
                if dir_skipped:
                    success, message = True, f"Skipped existing output directory: {output_subdir}"
                else:
                    os.makedirs(resolved_subdir, exist_ok=True)
                    success, message = convert_pdf_to_images(
                        input_pdf_path,
                        resolved_subdir,
                        image_format=normalized_target.lower(),
                        dpi=render_dpi,
                        progress_callback=file_progress,
                        page_range_text=page_range_text,
                        input_password=input_password,
                        output_conflict_policy=output_conflict_policy,
                        jpg_quality=jpg_quality,
                        cancel_event=cancel_event,
                    )

            if not success and message == CANCELLED_MESSAGE:
                cancelled = True
                failure_rows.append((filename, "cancelled", message))
                break

            if success and message.startswith("Skipped"):
                skipped_count += 1
                failure_rows.append((filename, "skipped", message))
            elif success:
                converted_count += 1
            else:
                failed_count += 1
                failure_rows.append((filename, "failed", message))

            _set_progress(progress_callback, ((file_index + 1) / total_files) * 100)

        log_suffix = ""
        if write_failure_log and failure_rows:
            log_path = _write_batch_failure_log(output_dir, failure_rows)
            log_suffix = f" Failure log: {log_path}"

        if cancelled:
            return (
                False,
                f"{CANCELLED_MESSAGE} Converted {converted_count}/{total_files} files. "
                f"Failed: {failed_count}, Skipped: {skipped_count}.{log_suffix}",
            )
        if failed_count > 0:
            return (
                False,
                f"Completed with errors. Converted {converted_count}/{total_files}, "
                f"Failed: {failed_count}, Skipped: {skipped_count}.{log_suffix}",
            )
        if skipped_count > 0:
            return (
                True,
                f"Completed with skips. Converted {converted_count}/{total_files}, "
                f"Skipped: {skipped_count}.{log_suffix}",
            )
        return True, f"Batch conversion successful! Converted {converted_count} files."
    except Exception as exc:
        return False, str(exc)
