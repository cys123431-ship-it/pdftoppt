import io
import os
import tempfile
from typing import Callable, Optional, Sequence

import fitz  # PyMuPDF
from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches

ProgressCallback = Optional[Callable[[int], None]]


def _set_progress(progress_callback: ProgressCallback, value: float) -> None:
    if progress_callback:
        clamped = max(0, min(100, int(value)))
        progress_callback(clamped)


def parse_page_range(page_range_text: str, total_pages: int) -> list[int]:
    """
    Parses page ranges like "1-3,5,8-10" into zero-based page indices.
    Empty input means all pages.
    """
    if total_pages <= 0:
        return []

    if not page_range_text or not page_range_text.strip():
        return list(range(total_pages))

    selected_pages: set[int] = set()
    tokens = page_range_text.split(",")

    for token in tokens:
        token = token.strip()
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


def _create_temp_pdf_with_selected_pages(pdf_path: str, page_indices: Sequence[int]) -> str:
    source_doc = fitz.open(pdf_path)
    temp_doc = fitz.open()

    try:
        for page_index in page_indices:
            temp_doc.insert_pdf(source_doc, from_page=page_index, to_page=page_index)

        file_descriptor, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
        os.close(file_descriptor)
        temp_doc.save(temp_pdf_path)
        return temp_pdf_path
    finally:
        temp_doc.close()
        source_doc.close()


def convert_pdf_to_pptx(
    pdf_path: str,
    pptx_path: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
) -> tuple[bool, str]:
    """Converts selected pages from PDF to PPTX (one page image per slide)."""
    doc = None
    try:
        doc = fitz.open(pdf_path)
        selected_pages = parse_page_range(page_range_text, len(doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        prs = Presentation()
        first_page = doc[selected_pages[0]]
        prs.slide_width = Inches(first_page.rect.width / 72.0)
        prs.slide_height = Inches(first_page.rect.height / 72.0)

        total_pages = len(selected_pages)
        for index, page_number in enumerate(selected_pages):
            page = doc[page_number]
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))

            img_data = pix.tobytes("png")
            image_stream = io.BytesIO(img_data)

            blank_slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_slide_layout)
            slide.shapes.add_picture(
                image_stream,
                0,
                0,
                width=prs.slide_width,
                height=prs.slide_height,
            )
            _set_progress(progress_callback, ((index + 1) / total_pages) * 100)

        prs.save(pptx_path)
        return True, "Conversion successful!"
    except Exception as exc:
        return False, str(exc)
    finally:
        if doc is not None:
            doc.close()


def convert_pdf_to_docx(
    pdf_path: str,
    docx_path: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
) -> tuple[bool, str]:
    """Converts selected pages from PDF to DOCX."""
    cv = None
    temp_pdf_path = None

    try:
        source_doc = fitz.open(pdf_path)
        try:
            total_pages = len(source_doc)
            selected_pages = parse_page_range(page_range_text, total_pages)
        finally:
            source_doc.close()

        if not selected_pages:
            return False, "The PDF has no pages."

        all_pages = list(range(total_pages))
        input_pdf_path = pdf_path
        if selected_pages != all_pages:
            temp_pdf_path = _create_temp_pdf_with_selected_pages(pdf_path, selected_pages)
            input_pdf_path = temp_pdf_path

        _set_progress(progress_callback, 10)
        cv = Converter(input_pdf_path)
        cv.convert(docx_path)
        _set_progress(progress_callback, 100)
        return True, "Conversion successful!"
    except Exception as exc:
        return False, str(exc)
    finally:
        if cv is not None:
            cv.close()
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)


def convert_pdf_to_images(
    pdf_path: str,
    output_dir: str,
    image_format: str = "png",
    dpi: int = 144,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
) -> tuple[bool, str]:
    """Converts selected PDF pages to PNG or JPG files."""
    doc = None
    try:
        if dpi <= 0:
            return False, "DPI must be greater than 0."

        normalized_format = image_format.lower()
        if normalized_format not in {"png", "jpg", "jpeg"}:
            return False, f"Unsupported image format: {image_format}"

        output_extension = "jpg" if normalized_format in {"jpg", "jpeg"} else "png"
        pixmap_format = "jpeg" if output_extension == "jpg" else "png"

        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]

        doc = fitz.open(pdf_path)
        selected_pages = parse_page_range(page_range_text, len(doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        zoom = dpi / 72.0
        total_pages = len(selected_pages)

        for index, page_number in enumerate(selected_pages):
            page = doc[page_number]
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))

            output_name = f"{base_name}_p{page_number + 1:03d}.{output_extension}"
            output_path = os.path.join(output_dir, output_name)
            image_bytes = pix.tobytes(pixmap_format)
            with open(output_path, "wb") as output_file:
                output_file.write(image_bytes)

            _set_progress(progress_callback, ((index + 1) / total_pages) * 100)

        return True, f"Saved {total_pages} image files."
    except Exception as exc:
        return False, str(exc)
    finally:
        if doc is not None:
            doc.close()


def merge_pdfs(
    input_pdf_paths: Sequence[str],
    output_pdf_path: str,
    progress_callback: ProgressCallback = None,
) -> tuple[bool, str]:
    """Merges multiple PDFs into one file."""
    if len(input_pdf_paths) < 2:
        return False, "Select at least 2 PDF files to merge."

    merged_doc = fitz.open()
    try:
        total_files = len(input_pdf_paths)
        for index, input_path in enumerate(input_pdf_paths):
            source_doc = fitz.open(input_path)
            try:
                merged_doc.insert_pdf(source_doc)
            finally:
                source_doc.close()

            _set_progress(progress_callback, ((index + 1) / total_files) * 100)

        merged_doc.save(output_pdf_path)
        return True, "Merge successful!"
    except Exception as exc:
        return False, str(exc)
    finally:
        merged_doc.close()


def split_pdf(
    pdf_path: str,
    output_dir: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
) -> tuple[bool, str]:
    """Splits selected pages into one-page PDF files."""
    source_doc = None
    try:
        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]

        source_doc = fitz.open(pdf_path)
        selected_pages = parse_page_range(page_range_text, len(source_doc))
        if not selected_pages:
            return False, "The PDF has no pages."

        total_pages = len(selected_pages)
        for index, page_number in enumerate(selected_pages):
            split_doc = fitz.open()
            try:
                split_doc.insert_pdf(source_doc, from_page=page_number, to_page=page_number)
                output_name = f"{base_name}_p{page_number + 1:03d}.pdf"
                output_path = os.path.join(output_dir, output_name)
                split_doc.save(output_path)
            finally:
                split_doc.close()

            _set_progress(progress_callback, ((index + 1) / total_pages) * 100)

        return True, f"Created {total_pages} split PDF files."
    except Exception as exc:
        return False, str(exc)
    finally:
        if source_doc is not None:
            source_doc.close()


def batch_convert_folder(
    input_dir: str,
    output_dir: str,
    target_format: str,
    progress_callback: ProgressCallback = None,
    page_range_text: str = "",
) -> tuple[bool, str]:
    """
    Batch converts all PDFs in a folder into target format.
    Supported target formats: PPTX, DOCX, PNG, JPG
    """
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
        success_count = 0
        failures: list[str] = []

        for file_index, filename in enumerate(pdf_files):
            input_pdf_path = os.path.join(input_dir, filename)
            base_name = os.path.splitext(filename)[0]

            def file_progress(child_percent: int, index: int = file_index) -> None:
                overall = ((index + (child_percent / 100.0)) / total_files) * 100.0
                _set_progress(progress_callback, overall)

            if normalized_target == "PPTX":
                output_path = os.path.join(output_dir, f"{base_name}.pptx")
                success, message = convert_pdf_to_pptx(
                    input_pdf_path,
                    output_path,
                    file_progress,
                    page_range_text,
                )
            elif normalized_target == "DOCX":
                output_path = os.path.join(output_dir, f"{base_name}.docx")
                success, message = convert_pdf_to_docx(
                    input_pdf_path,
                    output_path,
                    file_progress,
                    page_range_text,
                )
            else:
                per_file_output_dir = os.path.join(output_dir, base_name)
                success, message = convert_pdf_to_images(
                    input_pdf_path,
                    per_file_output_dir,
                    normalized_target.lower(),
                    144,
                    file_progress,
                    page_range_text,
                )

            if success:
                success_count += 1
            else:
                failures.append(f"{filename}: {message}")

            _set_progress(progress_callback, ((file_index + 1) / total_files) * 100)

        if failures:
            preview = "; ".join(failures[:3])
            if len(failures) > 3:
                preview += f"; and {len(failures) - 3} more"
            return False, f"Converted {success_count}/{total_files}. Failed: {preview}"

        return True, f"Batch conversion successful! Converted {success_count} files."
    except Exception as exc:
        return False, str(exc)
