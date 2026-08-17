import os
import tempfile
import unittest
from unittest.mock import patch

import fitz

from ocr_tools import (
    OCR_MODE_AUTO,
    OCR_MODE_FORCE,
    ocr_pdf_to_searchable_pdf,
    ocr_pdf_to_text,
    page_needs_ocr,
)


def make_pdf(path: str, text: str | None) -> None:
    doc = fitz.open()
    try:
        page = doc.new_page(width=400, height=250)
        if text:
            page.insert_text((40, 80), text, fontsize=18)
        doc.save(path)
    finally:
        doc.close()


def mock_ocr_pdf_bytes(text: str = "MOCK OCR") -> bytes:
    doc = fitz.open()
    try:
        page = doc.new_page(width=400, height=250)
        page.insert_text((40, 80), text, fontsize=18)
        return doc.tobytes()
    finally:
        doc.close()


class OcrToolsTests(unittest.TestCase):
    def test_page_needs_ocr_auto_and_force(self):
        doc = fitz.open()
        try:
            text_page = doc.new_page(width=400, height=250)
            text_page.insert_text((40, 80), "This page already has useful text.")
            self.assertFalse(page_needs_ocr(text_page, OCR_MODE_AUTO))
            self.assertTrue(page_needs_ocr(text_page, OCR_MODE_FORCE))

            blank = doc.new_page(width=400, height=250)
            self.assertTrue(page_needs_ocr(blank, OCR_MODE_AUTO))
        finally:
            doc.close()

    @patch("ocr_tools.ensure_ocr_ready", return_value=("fake", "fake-data"))
    @patch("ocr_tools._run_tesseract")
    def test_auto_text_export_skips_ocr_for_text_pdf(self, mocked_run, _mocked_ready):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "text.pdf")
            target = os.path.join(tmp, "text.txt")
            make_pdf(source, "ALREADY SEARCHABLE TEXT")

            success, message = ocr_pdf_to_text(
                source,
                target,
                ocr_language="eng",
                ocr_mode=OCR_MODE_AUTO,
            )
            self.assertTrue(success, message)
            mocked_run.assert_not_called()
            with open(target, "r", encoding="utf-8") as handle:
                self.assertIn("ALREADY SEARCHABLE TEXT", handle.read())

    @patch("ocr_tools.ensure_ocr_ready", return_value=("fake", "fake-data"))
    @patch("ocr_tools._run_tesseract", return_value=b"MOCK OCR TEXT")
    def test_blank_pdf_uses_ocr_for_text_export(self, mocked_run, _mocked_ready):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "scan.pdf")
            target = os.path.join(tmp, "scan.txt")
            make_pdf(source, None)

            success, message = ocr_pdf_to_text(
                source,
                target,
                ocr_language="eng",
                ocr_mode=OCR_MODE_AUTO,
                ocr_dpi=150,
            )
            self.assertTrue(success, message)
            mocked_run.assert_called_once()
            with open(target, "r", encoding="utf-8") as handle:
                self.assertIn("MOCK OCR TEXT", handle.read())

    @patch("ocr_tools.ensure_ocr_ready", return_value=("fake", "fake-data"))
    def test_searchable_pdf_contains_mocked_ocr_text(self, _mocked_ready):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "scan.pdf")
            target = os.path.join(tmp, "searchable.pdf")
            make_pdf(source, None)

            with patch("ocr_tools._run_tesseract", return_value=mock_ocr_pdf_bytes("FOUND OCR")):
                success, message = ocr_pdf_to_searchable_pdf(
                    source,
                    target,
                    ocr_language="eng",
                    ocr_mode=OCR_MODE_AUTO,
                    ocr_dpi=150,
                )
            self.assertTrue(success, message)
            result = fitz.open(target)
            try:
                self.assertIn("FOUND OCR", result[0].get_text("text"))
            finally:
                result.close()


if __name__ == "__main__":
    unittest.main()
